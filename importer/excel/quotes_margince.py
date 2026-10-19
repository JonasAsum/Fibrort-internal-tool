"""Margince adapter for excel/quotes.py: the quote log as deals with offers.

One enquiry (a quote number) is one DEAL on its account; every revision is one OFFER under it, with the
price as one line item. The CRM mints its own offer number and cannot take the Excel one, and the offer
has no source id, so the Excel quote number is kept in a deal custom field ("Quote number") and in the
offer's intro text; that field is how a second run finds the deals it wrote.

What is deliberately NOT done:
  - A won quote never becomes a won deal or an accepted offer: the ERP orders already feed the won
    revenue (excel/orders_margince.py), so a win here would count the sale twice. A won quote stays an
    open deal and says "Won order" in its description and its status field.
  - No exchange rate is applied. The log names no currency; `currency` (default EUR, matched against
    the ERP orders: one linked order is exactly 8.0 CNY per quote-log unit) is written on every deal and
    offer, and the dry run says so.
  - Sending an offer freezes an FX rate and is refused when the CRM has none: that offer stays a draft.

  preview(client, book)  reads the CRM only; what a commit would write
  commit(client, book)   writes; a second run writes nothing
"""
from zoneinfo import ZoneInfo

from common.client import ApiError
from common.fields import ensure_field
from excel.orders_margince import AccountRun, _match_result, company_ids
from visits import pipeline, when
from visits.reader import Parsed

SOURCE_SYSTEM = "quotelog"
FIELD_SOURCE = "quote log"
DEFAULT_CURRENCY = "EUR"
TZ = ZoneInfo("Asia/Shanghai")
NAME_MAX = 200
# attr -> (label, custom field type)
FIELDS = {"quote_no": ("Quote number", "text"), "status": ("Quote status", "text"),
          "end_customer": ("End customer", "text"), "product_group": ("Product group", "text"),
          "erp_order": ("ERP order number", "text")}


def minor(amount) -> int | None:
    return None if amount is None else round(amount * 100)


def deal_name(quotes) -> str:
    q = quotes[-1]
    what = q.product_type or q.product_group or "Quote"
    text = f"{q.customer} - {what} ({q.deal_key})"
    return text if len(text) <= NAME_MAX else text[:NAME_MAX - 1].rstrip() + "…"


def deal_description(quotes) -> str:
    q = quotes[-1]
    lines = [f"Quote log {q.deal_key}: {len(quotes)} revision{'s' if len(quotes) != 1 else ''}."]
    if q.status:
        lines.append(f"Status in the log: {q.status}" + (f" (lost index {q.lost_index})" if q.lost_index else ""))
    if q.bucket == "won":
        lines.append("Won order per the quote log. Kept open here on purpose: the ERP orders carry the won revenue.")
    for label, value in (("End customer", q.end_customer), ("Contact", q.contact), ("Product", q.product_type),
                         ("Article", q.article), ("Representative", q.representative), ("Clerk", q.clerk),
                         ("Lead time", q.lead_time), ("RFQ date", q.rfq_date), ("Order date", q.order_date),
                         ("ERP order", q.erp_order)):
        if value:
            lines.append(f"{label}: {value}")
    lines.append(f"Source: {q.file}, sheet {q.sheet}, row {q.row}")
    return "\n".join(lines)


def lost_reason(q) -> str:
    return q.status if q.status else (f"Lost index {q.lost_index}" if q.lost_index else "Lost (quote log)")


def seat_names(name: str) -> list[str]:
    """'Liu, Hao' and 'Betty Wu/Ella Lin' to the names a seat could carry, the first person first."""
    first = name.split("/")[0].strip()
    if "," in first:
        last, given = (p.strip() for p in first.split(",", 1))
        return [f"{given} {last}", f"{last} {given}"]
    return [first] if first else []


def offer_body(q, currency) -> dict:
    intro = [f"Quote log {q.quote_no or q.deal_key}, revision {q.revision}"]
    for label, value in (("RFQ", q.rfq_date), ("Quote date", q.quote_date), ("Lead time", q.lead_time)):
        if value:
            intro.append(f"{label}: {value}")
    body = {"currency": currency, "source": "import", "intro_text": "\n".join(intro)}
    unit = q.price if q.price else (round(q.total / q.quantity, 2) if q.total and q.quantity else q.total)
    if unit:
        body["line_items"] = [{"description": q.product_type or q.product_group or "Quote log item",
                               "quantity": q.quantity if q.quantity and q.quantity > 0 else 1,
                               "unit_price_minor": minor(unit)}]
    return body


def remarks_note(quotes, deal_id) -> dict | None:
    text = "\n".join(dict.fromkeys(q.remark for q in quotes if q.remark))
    if not text:
        return None
    key = quotes[0].deal_key
    day = max((q.quote_date for q in quotes if q.quote_date), default=None)
    note = {"source": "import", "source_system": SOURCE_SYSTEM, "source_id": f"quote:{key}:remarks", "kind": "note",
            "subject": "Follow-up remarks (quote log)", "body": text,
            "links": [{"entity_type": "deal", "entity_id": deal_id}]}
    if day:
        note["occurred_at"] = when.day_instant(day, TZ).isoformat()
    return note


class QuoteRun(AccountRun):
    """Company writes for the quote log: a quoted company is a prospect, not a customer, so only the
    other-language name is filled in."""

    def enrich_company(self, company_id, company) -> int:
        current = self.client.get_company(company_id)
        other = pipeline.other_language_name(current.get("display_name") or "", company.written_names)
        if other and not current.get("legal_name"):
            self.client.patch(f"/v1/companies/{company_id}", {"legal_name": other})
            return 1
        return 0


class _Lookup:
    """What the CRM holds, read once: the quote log's deals by quote number, the seats, the notes."""

    def __init__(self, client, create_fields: bool):
        self.columns = {}
        for attr, (label, type_) in FIELDS.items():
            if create_fields:
                self.columns[attr] = ensure_field(client, "deal", label, type_, None, FIELD_SOURCE)["column_name"]
            else:
                found = next((f for f in client.custom_fields("deal") if f["label"].lower() == label.lower()
                              and f["type"] == type_), None)
                if found:
                    self.columns[attr] = found["column_name"]
        column = self.columns.get("quote_no")
        self.deals = {d[column]: d for d in client.list_deals() if column and d.get(column)}
        self.seats = {when.fold_person(n): i for n, i in client.seats_by_name().items()}
        self.notes = {a.get("source_id") for a in client.paginate("/v1/activities", {"limit": 200})
                      if a.get("source_system") == SOURCE_SYSTEM}
        self.offers = {}
        self.client = client

    def seat(self, quotes):
        for q in reversed(quotes):
            for name in seat_names(q.responsible):
                if when.fold_person(name) in self.seats:
                    return self.seats[when.fold_person(name)]
        return None

    def held_offers(self, key) -> int:
        deal = self.deals.get(key)
        if deal is None:
            return 0
        if key not in self.offers:
            self.offers[key] = len(self.client.deal_offers(deal["id"]))
        return self.offers[key]


def plan(client, book, currency=DEFAULT_CURRENCY) -> dict:
    """What a commit would write, from reads only."""
    ids = company_ids(client, _match_result(book))
    lookup = _Lookup(client, create_fields=False)
    deals = book.deals()
    new = [k for k in deals if k not in lookup.deals]
    offers_new = sum(len(qs) - lookup.held_offers(k) for k, qs in deals.items())
    return {"currency": currency, "companies_new": len(book.accepted_orgs()) - len(ids), "companies_present": len(ids),
            "deals_new": len(new), "deals_present": len(deals) - len(new), "offers_new": offers_new,
            "notes_new": sum(1 for k, qs in deals.items() if any(q.remark for q in qs)
                             and f"quote:{k}:remarks" not in lookup.notes),
            "fields_missing": [FIELDS[a][0] for a in FIELDS if a not in lookup.columns],
            "quotes_waiting_for_a_company_answer": len(book.quotes) - len(book.accepted())}


def commit(client, book, currency=DEFAULT_CURRENCY, on_progress=None) -> dict:
    """Writes the deals, offers and remark notes the CRM does not hold yet; a second run writes nothing."""
    emit = on_progress or (lambda step, **info: None)
    result = _match_result(book)
    run = QuoteRun(client, Parsed(visits=[], weeks=[]), result, pipeline.Options(source_system=SOURCE_SYSTEM), emit)
    run.import_companies()
    lookup = _Lookup(client, create_fields=True)
    pipeline_id, stage_id = client.default_stage()
    out = {"companies": run.report.companies_report["disposition"], "deals_created": 0, "deals_present": 0,
           "deals_closed_lost": 0, "offers_created": 0, "offers_sent": 0, "offers_rejected": 0, "offers_left_draft": 0,
           "notes_created": 0, "failed": []}
    deals = book.deals()
    emit("quotes_start", total=len(deals))
    for i, (key, quotes) in enumerate(deals.items()):
        last = quotes[-1]
        company_id = run.company_ids.get(last.company_key)
        if company_id is None:
            out["failed"].append({"id": key, "error": "no account in the CRM"})
            continue
        try:
            deal = lookup.deals.get(key)
            if deal is None:
                body = {"name": deal_name(quotes), "pipeline_id": pipeline_id, "stage_id": stage_id,
                        "company_id": company_id, "source": "import", "source_system": SOURCE_SYSTEM,
                        "description": deal_description(quotes), "currency": currency,
                        lookup.columns["quote_no"]: key}
                if last.amount:
                    body["amount_minor"] = minor(last.amount)
                if seat := lookup.seat(quotes):
                    body["owner_id"] = seat
                for attr in ("status", "end_customer", "product_group", "erp_order"):
                    if getattr(last, attr):
                        body[lookup.columns[attr]] = getattr(last, attr)
                deal = lookup.deals[key] = client.create_deal(body)
                out["deals_created"] += 1
                if last.bucket == "lost":
                    try:
                        client.patch(f"/v1/deals/{deal['id']}", {"status": "lost", "lost_reason": lost_reason(last)})
                        out["deals_closed_lost"] += 1
                    except ApiError:
                        pass  # stays open; the status is in its description
            else:
                out["deals_present"] += 1
            held = lookup.held_offers(key)
            for n, q in enumerate(quotes):
                if n < held:
                    continue
                offer = client.create_offer(deal["id"], offer_body(q, currency))
                out["offers_created"] += 1
                if n == len(quotes) - 1 and q.bucket in ("open", "lost", "won"):  # only the live revision is sent
                    try:
                        client.offer_action(offer["id"], "send")
                        out["offers_sent"] += 1
                        if q.bucket == "lost":
                            client.offer_action(offer["id"], "reject", {"reason": lost_reason(q)})
                            out["offers_rejected"] += 1
                    except ApiError:
                        out["offers_left_draft"] += 1
            note = remarks_note(quotes, deal["id"])
            if note and note["source_id"] not in lookup.notes:
                client.log_activity(note)
                lookup.notes.add(note["source_id"])
                out["notes_created"] += 1
        except ApiError as exc:
            out["failed"].append({"id": key, "error": f"HTTP {exc.response.status_code}"})
        emit("quotes_progress", done=i + 1, total=len(deals), created=out["deals_created"])
    return out
