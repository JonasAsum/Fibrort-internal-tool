"""Turns extracted.json (+ the merged readings) into the list of records to write.

No network here: `parse` and `preview` mode print these counts, `commit` writes them.

enrichment.json (written by the reading pass, all keys optional):
  deals:    [{key, name, company, status: open|won|lost, amount_minor, currency, expected_close, message_ids, note}]
  tasks:    [{id, subject, body, due, company, assignee, message_id, done}]
  notes:    [{id, subject, body, company, date, thread_key}]
  routing:  {message_id: {company, deal}}      overrides where a message belongs
  contacts: {email: {title, phones: [e164]}}
Every entry should carry `evidence` (a Message-ID or attachment sha); it is kept in the record's raw data.

Since the reading layer (READING.md, reading.py) the entries come from staging/enrichment/<case>.json, only the
items that were agreed or reviewed (load_enrichment). Everything a reading contributes is kept apart from what
code found - deals with origin "reading", plan.routing, plan.contact_updates, tasks and notes - so the importer
can write it under the AI agent's passport. A deal code made from a contract is never changed by a reading;
a reading only adds links and a note to it.

The reading pass (a model) may only REFERENCE what code extracted. validate_enrichment() rejects an entry that
names a company that is not kept, a contact email that was never seen, or states an order number, email or phone
that does not appear verbatim in the mails or documents. Company and person names never come from it: corrections
to those go into staging/decisions.json, which you write yourself.
"""
import json
import os
import re
from dataclasses import dataclass, field

CUSTOMER_TYPE_LABEL = {"dealer": "Dealer", "end_user": "End user", "manufacturer": "Manufacturer"}
DEAL_DOC_TYPES = {"order_confirmation": "order_no", "sales_contract": "contract_no"}


def fold(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


@dataclass
class Plan:
    companies: list[dict] = field(default_factory=list)
    contacts: list[dict] = field(default_factory=list)
    products: dict[str, dict] = field(default_factory=dict)
    deals: dict[str, dict] = field(default_factory=dict)
    emails: list[dict] = field(default_factory=list)
    tasks: list[dict] = field(default_factory=list)
    notes: list[dict] = field(default_factory=list)
    attachments: list[dict] = field(default_factory=list)
    routing: list[dict] = field(default_factory=list)  # read from the mails: {message_id, company, deals, evidence}
    contact_updates: list[dict] = field(default_factory=list)  # read from the mails: {email, title, evidence}
    reading: dict = field(default_factory=dict)  # how many read items are in / waiting for review / rejected
    warnings: list[str] = field(default_factory=list)

    def counts(self) -> dict:
        return {
            "companies": len(self.companies), "contacts": len(self.contacts),
            "contacts_with_phone": sum(1 for c in self.contacts if c["phones"]),
            "products": len(self.products), "deals": len(self.deals),
            "deals_with_amount": sum(1 for d in self.deals.values() if d.get("amount_minor")),
            "offers": sum(1 for d in self.deals.values() if d.get("lines")),
            "emails": len(self.emails), "tasks": len(self.tasks), "notes": len(self.notes),
            "attachments": len(self.attachments),
            "attachment_bytes": sum(a["size"] for a in self.attachments),
            "read_deals": sum(1 for d in self.deals.values() if d.get("origin") == "reading"),
            "routed_emails": len(self.routing), "title_updates": len(self.contact_updates),
        }


REF_RE = re.compile(r"\b(?:SO|QU|OC|PO|OBR)\s?\d{6,}\b|\b0024\d{6}\b", re.I)
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
PHONE_RE = re.compile(r"\+\d{8,15}")


def _corpus(extracted: dict) -> str:
    parts = []
    for m in extracted["messages"]:
        parts += [m["subject"], m["body"]] + [p["email"] for p in [m["from"]] + m["to"] + m["cc"]]
        parts += [a["filename"] for a in m["attachments"]]
    parts += [d.get("excerpt", "") for d in extracted["documents"].values()]
    return "\n".join(parts)


def _digits(s: str) -> str:
    return re.sub(r"\D", "", s)


def validate_enrichment(enrichment: dict, extracted: dict, kept_names: set[str]) -> tuple[dict, list[str]]:
    """Drops every enrichment entry that is not grounded in the extracted data. Returns (clean, rejections)."""
    corpus = _corpus(extracted)
    corpus_folded = re.sub(r"\s+", "", corpus).upper()
    corpus_digits = _digits(corpus)
    known_emails = {e for c in extracted["contacts"] for e in c["emails"]}
    known_msgs = {m["message_id"] for m in extracted["messages"]}
    rejected, clean = [], {"deals": [], "tasks": [], "notes": [], "routing": {}, "contacts": {}}

    def ungrounded(text: str) -> list[str]:
        bad = [r for r in REF_RE.findall(text or "") if re.sub(r"\s", "", r).upper() not in corpus_folded]
        bad += [e for e in EMAIL_RE.findall(text or "") if e.lower() not in corpus.lower()]
        bad += [p for p in PHONE_RE.findall(text or "") if _digits(p)[-8:] not in corpus_digits]
        return bad

    if enrichment.get("companies"):
        rejected.append(f"companies: {len(enrichment['companies'])} entries ignored - company names and legal names "
                        f"come from the extraction; corrections go into staging/decisions.json")
    for kind in ("deals", "tasks", "notes"):
        for e in enrichment.get(kind, []):
            ref = e.get("key") or e.get("id")
            if e.get("company") and e["company"] not in kept_names:
                rejected.append(f"{kind} {ref}: company {e['company']!r} is not a kept company")
                continue
            if kind == "deals" and re.sub(r"\s", "", e["key"]).upper() not in corpus_folded and not e[
                    "key"].startswith(("RFQ-", "INQ-")):
                rejected.append(f"deals {ref}: the number {e['key']} appears in no mail or document")
                continue
            bad = ungrounded(" ".join(str(e.get(f) or "") for f in ("note", "subject", "body", "name")))
            if bad:
                rejected.append(f"{kind} {ref}: not found in the mails: {bad[:3]}")
                continue
            clean[kind].append(e)
    for mid, r in enrichment.get("routing", {}).items():
        if mid not in known_msgs or (r.get("company") and r["company"] not in kept_names):
            rejected.append(f"routing {mid[:30]}: unknown message or company")
            continue
        clean["routing"][mid] = r
    for email, o in enrichment.get("contacts", {}).items():
        if email not in known_emails:
            rejected.append(f"contacts {email}: never seen in the mails")
            continue
        if o.get("full_name") or o.get("company"):
            rejected.append(f"contacts {email}: name/company ignored - those come from headers and signatures")
        phones = [p for p in o.get("phones", []) if _digits(p)[-8:] in corpus_digits]
        if len(phones) < len(o.get("phones", [])):
            rejected.append(f"contacts {email}: phone(s) not in any mail dropped")
        title = o.get("title") if o.get("title") and o["title"] in corpus else None
        if phones or title:
            clean["contacts"][email] = {**o, "phones": phones, "title": title}
    return clean, rejected


def load_enrichment(staging: str) -> tuple[dict | None, dict]:
    """(entries for build_plan, counts). The merged readings in staging/enrichment/ when there are any (only
    agreed or reviewed items); otherwise the hand-written staging/enrichment.json of the Acme pilot."""
    from .reading import load_for_plan  # reading imports this module
    enrichment, counts = load_for_plan(staging)
    if enrichment is None and os.path.exists(os.path.join(staging, "enrichment.json")):
        with open(os.path.join(staging, "enrichment.json")) as f:
            enrichment, counts = json.load(f), {"hand_written_enrichment_json": True}
    return enrichment, counts


def build_plan(extracted: dict, enrichment: dict | None = None, only_company: str | None = None) -> Plan:
    plan = Plan()
    docs = extracted["documents"]
    companies = {c["name"]: dict(c) for c in extracted["companies"] if c["keep"]}
    enrichment, rejected = validate_enrichment(enrichment or {}, extracted, set(companies))
    plan.warnings += [f"enrichment rejected - {r}" for r in rejected]
    routing = enrichment["routing"]
    contact_over = enrichment["contacts"]
    party_owner = {p: c["name"] for c in companies.values() for p in c.get("parties") or []}

    # Mails are written where code filed them. A reading's routing is kept apart (plan.routing) and applied
    # afterwards by the AI agent as a relink, so the CRM shows who moved it.
    messages = [m for m in extracted["messages"] if m["import"]]
    routed: dict[str, dict] = {}
    for mid, r in routing.items():
        routed[mid] = {"message_id": mid, "company": r.get("company"), "deals": [r["deal"]] if r.get("deal") else [],
                       "evidence": r.get("evidence"), "reading": r.get("reading")}
    for m in messages:
        m["deal_key"] = None

    # deals from documents
    for m in messages:
        for a in m["attachments"]:
            d = docs.get(a["sha256"])
            parsed = d.get("parsed") if d else None
            if not parsed or parsed["type"] not in DEAL_DOC_TYPES:
                continue
            key = parsed.get(DEAL_DOC_TYPES[parsed["type"]])
            if not key:
                continue
            party = parsed.get("counterparty", "")
            # the extraction already tied every contract party to its company (Apex's contracts mailed inside
            # the Acme folder belong to Apex)
            company = party_owner.get(party) or m["company"]
            deal = plan.deals.setdefault(key, {"key": key, "name": None, "company": company, "status": "won",
                                               "currency": parsed.get("currency", "CNY"), "amount_minor": None,
                                               "expected_close": None, "lines": [], "docs": [], "messages": [],
                                               "note": None})
            deal["docs"].append({"sha256": a["sha256"], "filename": a["filename"], "type": parsed["type"]})
            deal["messages"].append(m["message_id"])
            if parsed.get("total_minor") and not deal["amount_minor"]:
                deal["amount_minor"] = parsed["total_minor"]
            if parsed.get("lines") and not deal["lines"]:
                deal["lines"] = parsed["lines"]
            dates = [ln.get("delivery_date") for ln in parsed.get("lines", []) if ln.get("delivery_date")]
            if dates and not deal["expected_close"]:
                deal["expected_close"] = max(dates)
            if parsed["type"] == "sales_contract" and not deal["company"]:
                deal["company"] = party_owner.get(parsed.get("counterparty", ""))
            if parsed["type"] == "order_confirmation" and parsed.get("customer_no"):
                deal["customer_no"] = parsed["customer_no"]
    for deal in plan.deals.values():
        deal["name"] = f"{deal['key']} — {deal['company'] or 'unassigned'}"
        deal["origin"] = "document"
    read_notes = []
    for e in enrichment.get("deals", []):
        for mid in e.get("message_ids", []):  # the mails a reading tied to the deal: linked by the agent
            routed.setdefault(mid, {"message_id": mid, "company": None, "deals": [], "evidence": e.get("evidence"),
                                    "reading": e.get("reading")})
            if e["key"] not in routed[mid]["deals"]:
                routed[mid]["deals"].append(e["key"])
        existing = plan.deals.get(e["key"])
        if existing and existing["docs"]:
            # code made this deal from its contract; a reading adds links and what it says, never other values
            if e.get("note"):
                read_notes.append({"id": f"note:deal:{e['key']}", "type": "deal", "subject": f"{e['key']}: what the mails say",
                                   "body": e["note"], "company": existing["company"], "deal": e["key"],
                                   "evidence": e.get("evidence"), "reading": e.get("reading")})
            continue
        deal = plan.deals.setdefault(e["key"], {"key": e["key"], "name": None, "company": None, "status": "open",
                                                "currency": e.get("currency", "CNY"), "amount_minor": None,
                                                "expected_close": None, "lines": [], "docs": [], "messages": [],
                                                "note": None, "origin": "reading"})
        for k in ("company", "status", "amount_minor", "currency", "expected_close", "note", "lines"):
            if e.get(k) is not None:
                deal[k] = e[k]
        deal["name"] = e.get("name") or f"{deal['key']} — {deal['company'] or 'unassigned'}"
        deal["evidence"] = e.get("evidence")
        deal["reading"] = e.get("reading")
    deals_of_message: dict[str, set] = {}
    for key, deal in plan.deals.items():
        for mid in deal["messages"]:
            deals_of_message.setdefault(mid, set()).add(key)
    deal_by_message = {mid: next(iter(keys)) for mid, keys in deals_of_message.items() if len(keys) == 1}
    for m in messages:
        m["deal_key"] = m["deal_key"] or deal_by_message.get(m["message_id"])
        if not m["deal_key"]:
            # a mail naming several orders (a batch announcement) belongs to none of them in particular
            cands = {k for k in m.get("refs", {}).get("sales_order", []) + m.get("refs", {}).get("quotation", [])
                     if k in plan.deals}
            if len(cands) == 1:
                m["deal_key"] = cands.pop()
        deal = plan.deals.get(m["deal_key"])
        if deal and deal["company"] and fold(deal["company"]) != fold(m["company"] or ""):
            m["company"] = deal["company"]

    if only_company:
        messages = [m for m in messages if fold(m["company"]) == fold(only_company)]
        plan.deals = {k: d for k, d in plan.deals.items() if fold(d["company"]) == fold(only_company)}

    # products from deal lines (a product only a reading names is created by the agent)
    for deal in sorted(plan.deals.values(), key=lambda d: d["origin"] != "document"):
        for ln in deal["lines"]:
            model = ln.get("model")
            if not model:
                continue
            prod = plan.products.setdefault(model, {"sku": model, "name": model, "unit_price_minor": ln.get(
                "unit_price_minor") or 0, "currency": deal["currency"], "origin": deal["origin"]})
            if ln.get("unit_price_minor") and prod["origin"] == deal["origin"]:
                prod["unit_price_minor"] = ln["unit_price_minor"]

    # companies actually used
    used = {m["company"] for m in messages if m["company"]} | {d["company"] for d in plan.deals.values() if d["company"]}
    contact_companies = set()
    for email, over in contact_over.items():  # titles a reading found: filled in by the agent, never overwriting
        if over.get("title") or over.get("phones"):
            plan.contact_updates.append({"email": email, "title": over.get("title"), "phones": over.get("phones") or [],
                                         "evidence": over.get("evidence"), "reading": over.get("reading")})
    for c in extracted["contacts"]:
        if not c["keep"]:
            continue
        if only_company and fold(c["company"]) != fold(only_company) and not c["internal"]:
            continue
        plan.contacts.append(c)
        if c["company"]:
            contact_companies.add(c["company"])
    plan.companies = [companies[n] for n in sorted(used | contact_companies, key=fold) if n in companies]
    if not only_company:  # a kept company with no mail left (all routed elsewhere) is still worth having
        have = {c["name"] for c in plan.companies}
        plan.companies += [c for n, c in sorted(companies.items()) if n not in have]
    for c in plan.companies:
        c["relationship_label"] = c.get("relationship")
        c["customer_type_label"] = CUSTOMER_TYPE_LABEL.get(c.get("customer_type") or "")
    conflicts = [c["name"] for c in plan.companies if any((c.get("name_conflicts") or {}).values())]
    if conflicts:
        plan.warnings.append(f"legal name left empty, the mails disagree (see relevance.csv): {conflicts}")
    dropped = [c["name"] for c in extracted["companies"] if not c["keep"]]
    plan.warnings.append(f"{len(dropped)} companies not imported (reasons in staging/relevance.csv): "
                         f"{', '.join(sorted(dropped)[:12])}{' ...' if len(dropped) > 12 else ''}")

    # emails + attachments
    seen_sha: set[str] = set()
    for m in messages:
        plan.emails.append(m)
        for a in m["attachments"]:
            if a["stored_as"] and not a["decoration"] and a["sha256"] not in seen_sha:
                seen_sha.add(a["sha256"])
                plan.attachments.append({**a, "message_id": m["message_id"], "deal_key": m["deal_key"]})
    plan.tasks = enrichment["tasks"]
    plan.notes = enrichment["notes"] + read_notes
    imported = {m["message_id"] for m in plan.emails}
    plan.routing = [r for r in routed.values() if r["message_id"] in imported]
    undated = sum(1 for m in plan.emails if not m["date_iso"])
    if undated:
        plan.warnings.append(f"{undated} emails have no parseable date and will be skipped on commit")
    nodeal = [d["key"] for d in plan.deals.values() if not d["amount_minor"]]
    if nodeal:
        plan.warnings.append(f"deals without an amount: {', '.join(sorted(nodeal))}")
    return plan
