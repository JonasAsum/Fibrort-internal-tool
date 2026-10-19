"""Margince adapter for excel/receivables.py: open and overdue AR as a read-only number on the account page.

The CRM's own finance card is a mirror of an accounting connector and cannot be written through the API,
so the two figures go into company custom fields, created once and reused: "AR open (CNY)" and "AR overdue
(CNY)" (currency fields, whole fen) and "AR as-of date" (date). The as-of date is part of the value: a
newer snapshot replaces the figures, an older one never overwrites a newer one, and a customer that is
missing from a newer snapshot has paid, so its figures go to zero on that date. If the CRM refuses to create
custom fields, each account gets one note (source_system `erp-receivables`) that is updated in place instead.

  plan(client, book)    reads the CRM only; counts what commit would write
  commit(client, book)  writes; a second run writes nothing
"""
import re
from decimal import ROUND_HALF_UP, Decimal
from zoneinfo import ZoneInfo
from datetime import time

from common.client import ApiError
from visits import pipeline, when
from visits.match import Company, MatchResult
from visits.reader import Parsed

from .orders_margince import AccountRun, company_ids, money, upsert_note
from .receivables import COMPANY_OK

SOURCE_SYSTEM = "erp-receivables"
NOTE_ID = "ar:"
SUFFIX = " (receivables)"
ASOF_MARK = re.compile(r"^As of: (.+)$", re.M)
FIELDS = (("open", "AR open (CNY)", "currency"), ("overdue", "AR overdue (CNY)", "currency"),
          ("as_of", "AR as-of date", "date"))
READ_ONLY = ("Read-only figure copied from the ERP receivables list. The ERP is the source of truth: "
             "do not edit it here, a later import overwrites it.")
TZ = ZoneInfo("Asia/Shanghai")


class ReceivablesError(RuntimeError):
    """The snapshot cannot be written as it is (undated, or in a currency the CRM fields do not hold)."""


def minor(amount: Decimal) -> int:
    return int((amount * 100).quantize(Decimal(1), ROUND_HALF_UP))


def stored_minor(value):
    """A stored currency value as whole fen, or None when it is empty or in a shape this does not know."""
    if value is None or value == "":
        return None
    if isinstance(value, dict):
        value = value.get("amount_minor", value.get("minor"))
    try:
        return int(Decimal(str(value)))
    except Exception:
        return None


def _day(value) -> str:
    return str(value or "")[:10]


def _check(book):
    if not book.as_of or book.currency != "CNY":
        raise ReceivablesError("the snapshot has no date or does not name CNY as its currency; nothing is written")


def _match_result(book) -> MatchResult:
    result = MatchResult()
    for c in book.accepted():
        known = c.match if c.match.get("score", 0) >= COMPANY_OK else {}
        result.companies[c.key] = Company(
            key=c.key, display_name=c.name, city="", written_names=list(c.written_names),
            list_names=[n for n in (known.get("cn"), known.get("en")) if n])
    return result


def _find_fields(client) -> dict:
    """The active company fields this importer uses, by role: {open|overdue|as_of: custom field}."""
    existing = {f["label"].lower(): f for f in client.custom_fields("company")}
    found = {}
    for role, label, type_ in FIELDS:
        for candidate in (label, label + SUFFIX):  # the suffixed one when a person's field of another type holds the label
            f = existing.get(candidate.lower())
            if f and f.get("type") == type_ and (type_ != "currency" or f.get("currency") == "CNY"):
                found[role] = f
                break
    return found


def ensure_fields(client) -> dict | None:
    """The three fields, created where missing; None when the CRM refuses a schema change (note mode)."""
    found = _find_fields(client)
    taken = {f["label"].lower() for f in client.custom_fields("company")}
    for role, label, type_ in FIELDS:
        if role in found:
            continue
        body = {"object": "company", "label": label + SUFFIX if label.lower() in taken else label,
                "type": type_}
        if type_ == "currency":
            body["currency"] = "CNY"
        try:
            found[role] = client.create_custom_field(body)
        except ApiError as e:
            if e.response.status_code in (400, 403, 409, 422, 501):
                return None
            raise
    return found


def decide(current: dict, cols: dict, target: dict) -> tuple[str, dict]:
    """(write | same | newer, the patch) for one account, given what the CRM holds."""
    held = _day(current.get(cols["as_of"]))
    if held and held > target["as_of"]:
        return "newer", {}
    patch = {}
    for role in ("open", "overdue"):
        if stored_minor(current.get(cols[role])) != target[role]:
            patch[cols[role]] = target[role]
    if held != target["as_of"]:
        patch[cols["as_of"]] = target["as_of"]
    return ("write" if patch else "same"), patch


def targets(book) -> dict:
    """customer key -> the figures to hold for it (whole fen, the snapshot's date)."""
    return {c.key: {"open": minor(c.open), "overdue": minor(c.overdue), "as_of": book.as_of}
            for c in book.accepted()}


def _paid(client, cols, book, ids_in_snapshot) -> dict:
    """Accounts that hold a figure from an older snapshot but are not in this one: paid, so zero.
    Skipped while any customer of the snapshot waits for a company answer (it may be one of them)."""
    if any(c.queued for c in book.customers.values()):
        return {}
    out = {}
    for co in client.paginate("/v1/companies", {"limit": 200}):
        held = _day(co.get(cols["as_of"]))
        if co["id"] in ids_in_snapshot or not held or held >= book.as_of:
            continue
        if stored_minor(co.get(cols["open"])) or stored_minor(co.get(cols["overdue"])):
            out[co["id"]] = {"open": 0, "overdue": 0, "as_of": book.as_of}
    return out


def plan(client, book) -> dict:
    """What a commit would write, from reads only."""
    _check(book)
    result = _match_result(book)
    ids = company_ids(client, result)
    cols = {role: f["column_name"] for role, f in _find_fields(client).items()}
    counts = {"write": 0, "same": 0, "newer": 0, "zero": 0}
    if len(cols) == len(FIELDS):
        want = targets(book)
        for key, target in want.items():
            if key in ids:
                counts[decide(client.get_company(ids[key]), cols, target)[0]] += 1
            else:
                counts["write"] += 1
        counts["zero"] = len(_paid(client, cols, book, set(ids.values())))
    else:
        counts["write"] = len(book.accepted())
    return {"mode": "fields" if len(cols) == len(FIELDS) else "fields (to be created)",
            "companies_new": len(result.companies) - len(ids), "companies_present": len(ids),
            "accounts": counts, "customers_waiting_for_a_company_answer": len(book.customers) - len(book.accepted())}


def commit(client, book, on_progress=None) -> dict:
    """Writes the two figures and the date onto each accepted customer's account; a second run writes nothing."""
    _check(book)
    emit = on_progress or (lambda step, **info: None)
    result = _match_result(book)
    run = AccountRun(client, Parsed(visits=[], weeks=[]), result, pipeline.Options(source_system=SOURCE_SYSTEM), emit)
    run.import_companies()
    fields = ensure_fields(client)
    report = {"companies": run.report.companies_report["disposition"], "errors": [],
              "customers_waiting_for_a_company_answer": len(book.customers) - len(book.accepted())}
    if fields is None:
        notes = _commit_notes(client, book, run)
        report["errors"] += notes.pop("errors")
        report.update(mode="notes", **notes)
        return report
    cols = {role: f["column_name"] for role, f in fields.items()}
    counts = {"write": 0, "same": 0, "newer": 0, "zero": 0}
    emit("accounts_start", total=len(book.accepted()))
    work = []
    for key, target in targets(book).items():
        if key in run.company_ids:
            work.append((run.company_ids[key], target, "write"))
        else:
            report["errors"].append(f"{key}: no account in the CRM")
    for cid, target in _paid(client, cols, book, set(run.company_ids.values())).items():
        work.append((cid, target, "zero"))
    for i, (cid, target, why) in enumerate(work):
        action, patch = decide(client.get_company(cid), cols, target)
        if action == "write":
            try:
                client.patch(f"/v1/companies/{cid}", patch)
                after = client.get_company(cid)
                bad = [k for k, v in patch.items()
                       if (stored_minor(after.get(k)) != v if k != cols["as_of"] else _day(after.get(k)) != v)]
                if bad:
                    report["errors"].append(f"{cid}: the CRM did not keep {', '.join(bad)}")
            except ApiError as e:
                report["errors"].append(f"{cid}: {e}")
                continue
            counts["zero" if why == "zero" else "write"] += 1
        else:
            counts[action] += 1
        emit("accounts_progress", done=i + 1, total=len(work))
    report.update(mode="fields", accounts_written=counts["write"], accounts_unchanged=counts["same"],
                  accounts_kept_newer=counts["newer"], accounts_zeroed=counts["zero"],
                  fields=[f["label"] for f in fields.values()])
    return report


def note_payload(customer, book, company_id) -> dict:
    body = "\n".join([f"Open receivables: CNY {money(customer.open)}",
                      f"Overdue: CNY {money(customer.overdue)}",
                      f"Open invoices: {customer.invoices}, overdue: {customer.overdue_invoices}"
                      + (f", oldest due date: {customer.oldest_due}" if customer.oldest_due else ""), "",
                      f"{READ_ONLY} Source: {book.source}.", f"As of: {book.as_of}"])
    return {"source": "import", "source_system": SOURCE_SYSTEM, "source_id": NOTE_ID + customer.key,
            "kind": "note", "subject": "Receivables (ERP)", "body": body,
            "occurred_at": when.day_instant(book.as_of, TZ, time(12, 0)).isoformat(),
            "links": [{"entity_type": "company", "entity_id": company_id}]}


def _commit_notes(client, book, run) -> dict:
    held = {a["source_id"]: a for a in client.paginate("/v1/activities", {"limit": 200})
            if a.get("source_system") == SOURCE_SYSTEM and a.get("source_id")}
    counts = {"create": 0, "update": 0, "same": 0, "newer": 0}
    errors = []
    for customer in book.accepted():
        cid = run.company_ids.get(customer.key)
        if cid is None:
            errors.append(f"{customer.key}: no account in the CRM")
            continue
        payload = note_payload(customer, book, cid)
        try:
            counts[upsert_note(client, held.get(payload["source_id"]), payload, ASOF_MARK)] += 1
        except ApiError as e:
            errors.append(f"{customer.key}: {e}")
    return {"notes_created": counts["create"], "notes_updated": counts["update"], "notes_unchanged": counts["same"],
            "notes_kept_newer": counts["newer"], "errors": errors}
