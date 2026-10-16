"""Margince adapter for excel/orders.py: read-only order history per account.

The CRM has no object for "an order that lives in the ERP". Won deals would be wrong here: they feed
the forecast and the won-revenue figures that the visits and the mails already fill, and an ERP order is
the same sale again. So every order is one NOTE on its account's timeline (source_system `erp-orders`,
source_id `order:<document no>`), worded as a copy of the ERP, and every account with orders gets one
overview note that sets the orders (the actual) beside its open deals (the pipeline). Nothing here
touches a deal, and no exchange rate is ever applied: amounts are shown in the currency the export
proves (CNY for the Chinese entity) or marked "currency not stated".

  plan(client, book)    reads the CRM only; counts what commit would write
  commit(client, book)  writes; a second run writes nothing, a later export only updates what changed
"""
import re
from collections import defaultdict
from datetime import time
from zoneinfo import ZoneInfo

from common.client import ApiError
from excel.split import fold_key
from visits import pipeline, when
from visits.match import Company, MatchResult
from visits.reader import Parsed

from .orders import COMPANY_OK

SOURCE_SYSTEM = "erp-orders"
OVERVIEW_ID = "orders-summary:"
TZ = ZoneInfo("Asia/Shanghai")
NOON = time(12, 0)
EXPORT_MARK = re.compile(r"^ERP export: (.+)$", re.M)
COUNTED = re.compile(r"^Orders counted: (.*)$", re.M)
MAX_ITEMS = 30
STATUS_LABEL = {"open": "Open", "to_ship": "To ship", "in_progress": "In progress", "completed": "Completed",
                "cancelled": "Cancelled", "other": ""}
READ_ONLY = ("Read-only history copied from the ERP sales-order export. The ERP is the source of truth: "
             "do not edit this note, a later import overwrites it.")


def money(x) -> str:
    return f"{x:,.2f}"


# ---- what gets written, independent of the network ----

def _value_lines(order) -> list[str]:
    if order.cny_gross is not None:
        lines = [f"Order value: CNY {money(order.cny_net)} excl. tax, CNY {money(order.cny_gross)} incl. tax"
                 if order.cny_net is not None else f"Order value: CNY {money(order.cny_gross)} incl. tax"]
        if order.currency != "CNY":
            lines.append(f"Ordered at {money(order.gross)} incl. tax in the order's own currency "
                         "(the export does not name it; no exchange rate is applied here).")
        return lines
    cur = order.currency
    head = f"Order value ({cur})" if cur else "Order value (currency not stated in the export)"
    return [f"{head}: {money(order.net)} excl. tax, {money(order.gross)} incl. tax"]


def _item_line(item) -> str:
    qty = f" x{item.quantity:g}" + (f" {item.unit}" if item.unit else "") if item.quantity is not None else ""
    amount = f" - {money(item.gross_amount)} incl. tax" if item.gross_amount is not None else ""
    return f"- {' '.join(x for x in (item.product_code, item.product_name) if x) or 'item'}{qty}{amount}"


def order_subject(order) -> str:
    value = ""
    if order.cny_gross is not None:
        value = f" - CNY {money(order.cny_gross)} incl. tax"
    elif order.currency:
        value = f" - {order.currency} {money(order.gross)} incl. tax"
    status = STATUS_LABEL[order.status_key] or order.status or ""
    return f"ERP order {order.document_no}" + (f" - {status}" if status else "") + value


def order_body(order) -> str:
    facts = [f"Order date: {order.date_iso}" if order.date_iso else "Order date: not stated"]
    if order.salesperson:
        facts.append(f"Sales rep: {order.salesperson}")
    if order.entity:
        facts.append(f"Selling company: {order.entity}")
    out = [" | ".join(facts), *_value_lines(order)]
    if order.status:
        out.append(f"Status in the ERP: {order.status}")
    if order.payment_status or order.collected:
        pay = order.payment_status or "payment status not stated"
        out.append(f"Payment: {pay}" + (f", collected {money(order.collected)}" if order.collected else ""))
    if order.planned_ship:
        out.append(f"Planned shipping: {order.planned_ship}")
    out.append(f"Products ({len(order.items)} line{'s' if len(order.items) != 1 else ''}):")
    out += [_item_line(i) for i in order.items[:MAX_ITEMS]]
    if len(order.items) > MAX_ITEMS:
        out.append(f"- and {len(order.items) - MAX_ITEMS} more lines")
    out += ["", f"{READ_ONLY} Source: {order.file}.", f"ERP export: {order.export_iso}"]
    return "\n".join(out)


def order_payload(order, company_id: str) -> dict:
    payload = {
        "source": "import", "source_system": SOURCE_SYSTEM, "source_id": order.source_id,
        "kind": "note", "subject": order_subject(order), "body": order_body(order),
        "occurred_at": when.day_instant(order.date_iso, TZ, NOON).isoformat(),
        "links": [{"entity_type": "company", "entity_id": company_id}],
        "raw": {"erp_order": {"document_no": order.document_no, "status": order.status, "entity": order.entity,
                              "file": order.file, "sheet": order.sheet, "export": order.export_iso}},
    }
    if order.salesperson:
        payload["source_author_name"] = order.salesperson
    return payload


def _year(order) -> str:
    return (order.date_iso or "")[:4] or "undated"


def _cny(order):
    return order.cny_net if order.cny_net is not None else order.cny_gross


def overview_body(orders, deals) -> str:
    """The actual (orders by order date) against the pipeline (open deals), per currency, never converted.
    orders: this account's Orders; deals: [(currency, amount_minor)] of its open deals."""
    live = sorted((o for o in orders if o.status_key != "cancelled"), key=lambda o: o.document_no)
    out = ["Orders in the ERP exports against open deals in the CRM. Read-only; refreshed by the order import.", ""]
    by_year = defaultdict(list)
    for o in live:
        by_year[_year(o)].append(o)
    for year in sorted(by_year):
        group = by_year[year]
        cny = [o for o in group if _cny(o) is not None]
        text = f"{year}: {len(group)} order{'s' if len(group) != 1 else ''}"
        if cny:
            tax = "excl. tax" if all(o.cny_net is not None for o in cny) else "incl. or excl. tax as the export gives it"
            unfinished = [o for o in cny if o.status_key != "completed"]
            text += f", CNY {money(sum(_cny(o) for o in cny))} {tax}"
            text += f"; {len(unfinished)} not completed" + (f" (CNY {money(sum(_cny(o) for o in unfinished))})" if unfinished else "")
        rest = len(group) - len(cny)
        if rest:
            text += ("; " if cny else ", ") + f"{rest} in a currency the export does not name (not summed, see the order notes)"
        out.append(text)
    pipe = defaultdict(lambda: [0, 0])
    for currency, minor in deals:
        pipe[currency or "no currency"][0] += 1
        pipe[currency or "no currency"][1] += minor
    if pipe:
        out.append("Open deals: " + "; ".join(f"{n} in {cur}: {money(total / 100)}" for cur, (n, total) in sorted(pipe.items())))
    else:
        out.append("Open deals: none in the CRM.")
    out += ["Amounts in different currencies are not added up or converted.", "",
            "Orders counted: " + ", ".join(o.document_no for o in live),
            f"ERP export: {max((o.export_iso for o in orders), default='')}"]
    return "\n".join(out)


def overview_subject() -> str:
    return "ERP orders against open deals"


# ---- the CRM side ----

def _match_result(book) -> MatchResult:
    result = MatchResult()
    for key, org in book.accepted_orgs().items():
        known = org.match if org.match.get("score", 0) >= COMPANY_OK else {}
        result.companies[key] = Company(
            key=key, display_name=org.name, city="", written_names=list(org.written_names),
            list_names=[n for n in (known.get("cn"), known.get("en")) if n])
    return result


class AccountRun(pipeline._Run):
    """The visit pipeline's company writes for customers with no visit behind them: an order proves the
    company is a customer; the other language's name is filled in where the CRM has none."""

    def enrich_company(self, company_id, company) -> int:
        current, patch = self.client.get_company(company_id), {}
        types = current.get("relationship_types") or []
        if "customer" not in types:
            patch["relationship_types"] = [*types, "customer"]
        if current.get("lifecycle") in (None, "unknown"):
            patch["lifecycle"] = "customer"
        other = pipeline.other_language_name(current.get("display_name") or "", company.written_names)
        if other and not current.get("legal_name"):
            patch["legal_name"] = other
        if not patch:
            return 0
        self.client.patch(f"/v1/companies/{company_id}", patch)
        return 1


def open_deals(client) -> dict[str, list]:
    """company id -> [(currency, amount_minor)] of its open deals (those whose amount the caller may read)."""
    out = defaultdict(list)
    for d in client.paginate("/v1/deals", {"status": "open", "limit": 200}):
        if d.get("company_id") and d.get("amount_minor") is not None:
            out[d["company_id"]].append((d.get("currency"), d["amount_minor"]))
    return out


def existing_notes(client) -> dict[str, dict]:
    return {a["source_id"]: a for a in client.paginate("/v1/activities", {"limit": 200})
            if a.get("source_system") == SOURCE_SYSTEM and a.get("source_id")}


def _stamp(body, mark=EXPORT_MARK) -> str:
    m = mark.search(body or "")
    return m.group(1).strip() if m else ""


def _counted(body) -> set:
    m = COUNTED.search(body or "")
    return {x.strip() for x in m.group(1).split(",") if x.strip()} if m else set()


def decide(found: dict | None, payload: dict, mark=EXPORT_MARK) -> str:
    """create | update | same | newer: what to do with a note given the one the CRM holds."""
    if found is None:
        return "create"
    if (found.get("body") or "") == payload["body"] and (found.get("subject") or "") == payload["subject"]:
        return "same"
    held, incoming = _stamp(found.get("body"), mark), _stamp(payload["body"], mark)
    if held and incoming and held > incoming:
        return "newer"  # an older export never overwrites a newer one
    if _counted(found.get("body")) - _counted(payload["body"]):
        return "newer"  # an overview that knows orders this run lacks (a run over fewer exports) stays
    return "update"


def upsert_note(client, found: dict | None, payload: dict, mark=EXPORT_MARK) -> str:
    """Writes payload as a new note, or updates the one found; returns what was done (see decide)."""
    action = decide(found, payload, mark)
    if action == "create":
        client.log_activity(payload)
    elif action == "update":
        client.patch(f"/v1/activities/{found['id']}", {"subject": payload["subject"], "body": payload["body"]})
    return action


def company_ids(client, result: MatchResult) -> dict[str, str]:
    by_name = {fold_key(n): i for n, i in client.company_id_by_name().items()}
    found = {}
    for key, company in result.companies.items():
        hit = next((n for n in pipeline._crm_names(company) if fold_key(n) in by_name), None)
        if hit:
            found[key] = by_name[fold_key(hit)]
    return found


def plan(client, book) -> dict:
    """What a commit would write, from reads only."""
    result = _match_result(book)
    ids = company_ids(client, result)
    held, deals = existing_notes(client), open_deals(client)
    counts = {"create": 0, "update": 0, "same": 0, "newer": 0}
    summaries = {"create": 0, "update": 0, "same": 0, "newer": 0}
    per_company = defaultdict(list)
    for order in book.accepted():
        per_company[order.company_key].append(order)
        counts[decide(held.get(order.source_id), order_payload(order, ids.get(order.company_key, "?")))] += 1
    for key, orders in per_company.items():
        cid = ids.get(key)
        payload = _overview_payload(key, orders, deals.get(cid, []), cid or "?")
        summaries[decide(held.get(payload["source_id"]), payload)] += 1
    return {"companies_new": len(result.companies) - len(ids), "companies_present": len(ids),
            "orders": counts, "overviews": summaries, "orders_waiting_for_a_company_answer": len(book.orders) - len(book.accepted())}


def _overview_payload(key, orders, deals, company_id) -> dict:
    latest = max((o.date_iso for o in orders if o.date_iso), default=None)
    payload = {
        "source": "import", "source_system": SOURCE_SYSTEM, "source_id": OVERVIEW_ID + key, "kind": "note",
        "subject": overview_subject(), "body": overview_body(orders, deals),
        "links": [{"entity_type": "company", "entity_id": company_id}],
    }
    if latest:
        payload["occurred_at"] = when.day_instant(latest, TZ, NOON).isoformat()
    return payload


def commit(client, book, on_progress=None) -> dict:
    """Writes the accepted orders and one overview per account; a second run writes nothing."""
    emit = on_progress or (lambda step, **info: None)
    result = _match_result(book)
    run = AccountRun(client, Parsed(visits=[], weeks=[]), result, pipeline.Options(source_system=SOURCE_SYSTEM), emit)
    run.import_companies()
    held, deals = existing_notes(client), open_deals(client)
    done = {"create": 0, "update": 0, "same": 0, "newer": 0}
    overview = {"create": 0, "update": 0, "same": 0, "newer": 0}
    errors, per_company = [], defaultdict(list)
    accepted = book.accepted()
    emit("orders_start", total=len(accepted))
    for i, order in enumerate(accepted):
        cid = run.company_ids.get(order.company_key)
        if cid is None:
            errors.append(f"{order.document_no}: no account in the CRM")
            continue
        per_company[order.company_key].append(order)
        try:
            done[upsert_note(client, held.get(order.source_id), order_payload(order, cid))] += 1
        except ApiError as e:
            errors.append(f"{order.document_no}: {e}")
        emit("orders_progress", done=i + 1, total=len(accepted))
    for key, orders in per_company.items():
        cid = run.company_ids[key]
        payload = _overview_payload(key, orders, deals.get(cid, []), cid)
        try:
            overview[upsert_note(client, held.get(payload["source_id"]), payload)] += 1
        except ApiError as e:
            errors.append(f"overview {key}: {e}")
    r = run.report
    return {"companies": r.companies_report["disposition"], "companies_enriched": r.companies_updated,
            "orders_created": done["create"], "orders_updated": done["update"], "orders_unchanged": done["same"],
            "orders_kept_newer": done["newer"], "overviews_created": overview["create"],
            "overviews_updated": overview["update"], "overviews_unchanged": overview["same"],
            "orders_waiting_for_a_company_answer": len(book.orders) - len(accepted), "errors": errors}
