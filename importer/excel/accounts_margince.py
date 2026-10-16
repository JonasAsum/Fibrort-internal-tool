"""Margince adapter for excel/accounts.py: the only place that knows the CRM for customer lists.

Where each part of a list lands:

  name, address, province, city, industry   company (display name = Chinese name, else English; the other
                                            language goes in legal_name, which the CRM search reads)
  ERP customer code, sales region, product  company custom fields (the code is how a re-run finds the
  area, dealer, supporting sales, ERP rep,  account again; ERP stays the source of truth for it)
  first order
  application, competitor, expected annual  the same custom fields, as OPPORTUNITY data. Not a deal: a deal
  volume                                    needs a pipeline stage and a money amount, and a list that
                                            gives units per year would force invented ones into the forecast.
  customer / prospect / competitor          lifecycle and relationship_types
  contact name + phone (analysis lists)     a contact employed at the company
  free remarks (visit purpose, exhibition   a note linked to the company
  intel, "Do Not use" flag)
  exhibition attendees, registration form   a lead (custom fields: phone, location, industry, application,
                                            details); a lead whose company the CRM already holds as an
                                            account is skipped, not duplicated

Nothing here overwrites: a company the CRM already holds (found by ERP code, else by name) only gets the
blanks filled; a value that differs is counted and listed as a conflict. Every write carries
source_system + a content-based source_id (leads, notes), or is found again by code/name (companies,
contacts), so a re-run writes nothing new.

  plan_counts(result)            counts from the read alone
  preview(client, result)        reads the CRM only, says what a commit would do
  commit(client, result, ...)    writes. Only ever point it at a local test CRM unless told otherwise.
"""
import hashlib
from dataclasses import asdict, dataclass, field
from datetime import datetime
from zoneinfo import ZoneInfo

from common.client import ApiError
from common.fields import ensure_field
from visits import when
from excel.split import fold_key

from .accounts import Account, AccountsResult, Lead
from .companies import norm

SOURCE_SYSTEM = "customer-lists"
FIELD_SOURCE = "customer lists"  # suffix of a label another importer already used with another type

# (Account attribute, label, type)
COMPANY_FIELDS = [
    ("code", "ERP customer code", "text"), ("region", "Sales region", "text"),
    ("product_area", "Product area", "text"), ("dealer", "Dealer", "text"),
    ("supporting", "Supporting sales", "text"), ("erp_rep", "ERP sales rep", "text"),
    ("first_order", "First order", "text"), ("application", "Application", "text"),
    ("competitor", "Competitor", "text"), ("forecast", "Expected annual volume", "number"),
]
# (Lead attribute, label, type)
LEAD_FIELDS = [
    ("phone", "Phone", "text"), ("location", "Location", "text"), ("industry", "Industry", "text"),
    ("application", "Application", "text"), ("details", "Details", "text"),
]
MOBILE_PREFIX = "+861"


@dataclass
class AccountsReport:
    companies_created: int = 0
    companies_updated: int = 0
    companies_unchanged: int = 0
    companies_ambiguous: list = field(default_factory=list)  # source ids the CRM holds twice
    conflicts: list = field(default_factory=list)  # {"account", "field", "crm", "list"}
    contacts_created: int = 0
    contacts_present: int = 0
    emails_held_elsewhere: list = field(default_factory=list)  # {"id", "contact", "emails"}: address already on another CRM contact
    notes_created: int = 0
    notes_idempotent: int = 0
    leads_created: int = 0
    leads_present: int = 0
    leads_already_accounts: int = 0
    fields_created: list = field(default_factory=list)
    failed: list = field(default_factory=list)  # {"id", "error"}: one bad record never stops the run


# ---- what gets written, independent of the network ----

def _vocabulary(client, object_, specs, create: bool, created: list | None = None) -> dict:
    """{attribute: custom field column}. Without `create` only fields that already exist are found."""
    columns = {}
    for attr, label, type_ in specs:
        if create:
            before = {f["label"] for f in client.custom_fields(object_)}
            found = ensure_field(client, object_, label, type_, None, FIELD_SOURCE)
            if found["label"] not in before and created is not None:
                created.append(f"{object_}:{found['label']}")
            columns[attr] = found["column_name"]
        else:
            found = next((f for f in client.custom_fields(object_) if f["label"].lower() == label.lower()
                          and f["type"] == type_), None)
            if found:
                columns[attr] = found["column_name"]
    return columns


def _same_value(current, wanted) -> bool:
    if isinstance(wanted, (int, float)) and not isinstance(wanted, bool):
        try:
            return float(current) == float(wanted)
        except (TypeError, ValueError):
            return False
    return " ".join(str(current).split()).casefold() == " ".join(str(wanted).split()).casefold()


def _number(value):
    return int(value) if isinstance(value, float) and value.is_integer() else value


def company_patch(acc: Account, current: dict, columns: dict) -> tuple[dict, list]:
    """(patch, conflicts): what to add to a company the CRM holds. Blanks are filled, nothing is
    overwritten; a value that differs is a conflict and stays as it is."""
    patch, conflicts = {}, []

    def fill(label, key, wanted, held):
        if wanted in ("", None):
            return
        if held in (None, ""):
            patch[key] = _number(wanted)
        elif not _same_value(held, wanted):
            conflicts.append({"account": acc.source_id, "field": label, "crm": held, "list": wanted})

    for attr, label, _ in COMPANY_FIELDS:
        if attr in columns:
            fill(label, columns[attr], getattr(acc, attr), current.get(columns[attr]))
    fill("Industry", "industry", acc.industry, current.get("industry"))
    fill("Legal name", "legal_name", acc.other_language_name(), current.get("legal_name"))
    address = dict(current.get("address") or {})
    for key, wanted in (("line1", acc.address), ("city", acc.city), ("region", acc.province)):
        if wanted and not address.get(key):
            address[key] = wanted
        elif wanted and not _same_value(address[key], wanted):
            conflicts.append({"account": acc.source_id, "field": f"Address {key}", "crm": address[key], "list": wanted})
    if address != (current.get("address") or {}):
        patch["address"] = address
    if acc.lifecycle and current.get("lifecycle") in (None, "unknown"):
        patch["lifecycle"] = acc.lifecycle
    held = list(current.get("relationship_types") or [])
    if missing := [t for t in acc.relationships if t not in held]:
        patch["relationship_types"] = held + missing
    return patch, conflicts


def lead_body(lead: Lead) -> dict:
    return {"company_name": lead.company, "full_name": lead.contact or None,
            "title": "Sales representative" if lead.contact and lead.origin == "exhibition" else None,
            "source": "import", "source_system": SOURCE_SYSTEM, "source_id": lead.source_id}


def note_id(acc: Account, text: str) -> str:
    return f"note:{acc.source_id}:{hashlib.sha1(' '.join(text.lower().split()).encode()).hexdigest()[:10]}"


def note_body(acc: Account, text: str, company_id: str, when_: datetime) -> dict:
    return {"source": "import", "source_system": SOURCE_SYSTEM, "source_id": note_id(acc, text), "kind": "note",
            "subject": "Customer list: " + (text if len(text) <= 70 else text[:67] + "..."), "body": text,
            "occurred_at": when_.isoformat(), "links": [{"entity_type": "company", "entity_id": company_id}]}


def contact_body(person) -> dict:
    """The CRM contact of a Person: every +86 / international number of the list (a mobile first), every e-mail
    address. A number without a country prefix was never placed by the reader, so it is not here."""
    phones = []
    for number in person.phones:
        if number not in (p["phone"] for p in phones):
            phones.append({"phone": number, "phone_type": "mobile" if number.startswith(MOBILE_PREFIX) else "work",
                           "is_primary": not phones, "position": len(phones)})
    body = {"full_name": person.name, "source": "import", "source_system": SOURCE_SYSTEM, "phones": phones}
    if person.emails:
        body["emails"] = [{"email": e, "email_type": "work", "is_primary": i == 0, "position": i}
                          for i, e in enumerate(person.emails)]
    return body


def plan_counts(result: AccountsResult) -> dict:
    """What a commit would write at most, counted from the read alone."""
    accounts, leads = result.accepted_accounts(), result.accepted_leads()
    return {
        "accounts": len(accounts),
        "accounts_with_code": sum(bool(a.code) for a in accounts),
        "customers": sum(a.lifecycle == "customer" for a in accounts),
        "prospects": sum(a.lifecycle == "prospect" for a in accounts),
        "competitors": sum("competitor" in a.relationships for a in accounts),
        "with_opportunity_data": sum(bool(a.competitor or a.application or a.forecast is not None) for a in accounts),
        "contacts": sum(len(a.contacts) for a in accounts),
        "notes": sum(len(a.notes) for a in accounts),
        "leads": len(leads),
        "queued": (len(result.accounts) - len(accounts)) + (len(result.leads) - len(leads)),
    }


# ---- the CRM side ----

class CrmIndex:
    """The companies the CRM holds, findable by ERP code and by name."""

    def __init__(self, companies, code_column):
        self.by_id, self.code_column = {}, code_column
        self.by_code, self.by_name = {}, {}
        for c in companies:
            self.add(c)

    @staticmethod
    def _key(name) -> str:
        return norm(name) or " ".join(str(name or "").split()).lower()

    def add(self, company: dict) -> None:
        cid = company["id"]
        self.by_id[cid] = company
        code = company.get(self.code_column) if self.code_column else None
        if code:
            self.by_code.setdefault(str(code).upper().lstrip("0"), set()).add(cid)
        for name in (company.get("display_name"), company.get("legal_name")):
            if name:
                self.by_name.setdefault(self._key(name), set()).add(cid)

    def find(self, acc: Account) -> tuple[str | None, bool]:
        """(company id, ambiguous): the one company the CRM holds for this account, if exactly one."""
        if acc.code:
            ids = self.by_code.get(acc.code, set())
            if len(ids) == 1:
                return next(iter(ids)), False
            if len(ids) > 1:
                return None, True
        ids = set()
        for name in acc.names:
            ids |= self.by_name.get(self._key(name), set())
        if acc.code:  # a company that already carries another ERP code is another customer
            ids = {i for i in ids if not self.by_id[i].get(self.code_column) or
                   str(self.by_id[i][self.code_column]).upper().lstrip("0") == acc.code}
        return (next(iter(ids)), False) if len(ids) == 1 else (None, len(ids) > 1)

    def has_name(self, name: str) -> bool:
        return self._key(name) in self.by_name


def _crm_state(client, with_fields: bool):
    columns = _vocabulary(client, "company", COMPANY_FIELDS, with_fields)
    companies = list(client.paginate("/v1/companies", {"limit": 200}))
    return columns, CrmIndex(companies, columns.get("code"))


def _lead_ids(client) -> set:
    return {l.get("source_id") for l in client.paginate("/v1/leads", {"limit": 200})
            if l.get("source_system") == SOURCE_SYSTEM}


def _note_ids(client) -> set:
    return {a.get("source_id") for a in client.paginate("/v1/activities", {"limit": 200})
            if a.get("source_system") == SOURCE_SYSTEM}


def preview(client, result: AccountsResult) -> dict:
    """Reads only. What a commit would do: companies to create or fill in, the conflicts it would leave
    alone, contacts, notes and leads that are new."""
    columns, index = _crm_state(client, with_fields=False)
    out = plan_counts(result)
    out.update(companies_new=0, companies_present=0, companies_to_fill=0, companies_ambiguous=0, conflicts=0,
               contacts_new=0, notes_new=0, leads_new=0, leads_already_accounts=0,
               fields_missing=[f"company:{label}" for attr, label, _ in COMPANY_FIELDS if attr not in columns])
    held_notes = _note_ids(client) if any(a.notes for a in result.accepted_accounts()) else set()
    for acc in result.accepted_accounts():
        cid, ambiguous = index.find(acc)
        if ambiguous:
            out["companies_ambiguous"] += 1
            continue
        known = set()
        if cid is None:
            out["companies_new"] += 1
        else:
            out["companies_present"] += 1
            patch, conflicts = company_patch(acc, index.by_id[cid], columns)
            out["companies_to_fill"] += bool(patch)
            out["conflicts"] += len(conflicts)
            if acc.contacts:
                known = {fold_key(c["full_name"]) for c in client.contacts_for_company(cid)}
        out["contacts_new"] += sum(fold_key(p.name) not in known for p in acc.contacts)
        out["notes_new"] += sum(note_id(acc, t) not in held_notes for t in acc.notes)
    held_leads = _lead_ids(client)
    for lead in result.accepted_leads():
        if lead.source_id in held_leads:
            continue
        if index.has_name(lead.company):
            out["leads_already_accounts"] += 1
        else:
            out["leads_new"] += 1
    return out


def _is_duplicate_email(exc: ApiError) -> bool:
    try:
        return exc.response.status_code == 409 and exc.response.json().get("code") == "duplicate_email"
    except ValueError:
        return False


class _Run:
    def __init__(self, client, result, emit):
        self.client, self.result, self.emit = client, result, emit
        self.report = AccountsReport()
        self.tz = ZoneInfo("Asia/Shanghai")
        self.now = datetime.now(self.tz).replace(microsecond=0)

    def vocabulary(self):
        self.emit("vocabulary")
        self.columns = _vocabulary(self.client, "company", COMPANY_FIELDS, True, self.report.fields_created)
        self.lead_columns = _vocabulary(self.client, "lead", LEAD_FIELDS, True, self.report.fields_created)
        self.seats = {when.fold_person(n): i for n, i in self.client.seats_by_name().items()}

    def create_company(self, acc: Account) -> dict:
        body = {"display_name": acc.display(), "source": "import", "source_system": SOURCE_SYSTEM}
        if (other := acc.other_language_name()):
            body["legal_name"] = other
        if acc.industry:
            body["industry"] = acc.industry
        address = {k: v for k, v in (("line1", acc.address), ("city", acc.city), ("region", acc.province)) if v}
        if address:
            body["address"] = address
        seat = self.seats.get(when.fold_person(acc.erp_rep)) if acc.erp_rep else None
        if seat:
            body["owner_id"] = seat
        try:
            return self.client.post("/v1/companies", json=body)
        except ApiError as exc:
            if seat and exc.response.status_code == 422:  # a seat that cannot be handed work
                return self.client.post("/v1/companies", json={k: v for k, v in body.items() if k != "owner_id"})
            raise

    def account(self, acc: Account, index: CrmIndex, known_notes: set) -> None:
        cid, ambiguous = index.find(acc)
        if ambiguous:
            self.report.companies_ambiguous.append(acc.source_id)
            return
        created = cid is None
        if created:
            company = self.create_company(acc)
            cid = company["id"]
            self.report.companies_created += 1
        else:
            company = index.by_id[cid]
        patch, conflicts = company_patch(acc, company, self.columns)
        self.report.conflicts += conflicts
        if patch:
            company = self.client.patch(f"/v1/companies/{cid}", patch) or {**company, **patch}
            self.report.companies_updated += not created
        else:
            self.report.companies_unchanged += not created
        index.add({**company, "id": cid})
        self.contacts(acc, cid)
        for text in acc.notes:
            sid = note_id(acc, text)
            if sid in known_notes:
                self.report.notes_idempotent += 1
                continue
            made, _ = self.client.log_activity(note_body(acc, text, cid, self.now))
            known_notes.add(sid)
            self.report.notes_created += made
            self.report.notes_idempotent += not made

    def _create_contact(self, acc: Account, person) -> dict:
        """Creates the contact. An e-mail address the CRM already holds on another contact is not stolen from it:
        the contact is created without that address and the case is listed in the report."""
        body = contact_body(person)
        try:
            return self.client.create_contact(body)
        except ApiError as exc:
            if not (body.get("emails") and _is_duplicate_email(exc)):
                raise
        self.report.emails_held_elsewhere.append({"id": acc.source_id, "contact": person.name, "emails": list(person.emails)})
        return self.client.create_contact({k: v for k, v in body.items() if k != "emails"})

    def contacts(self, acc: Account, company_id: str) -> None:
        if not acc.contacts:
            return
        held = {fold_key(c["full_name"]) for c in self.client.contacts_for_company(company_id)}
        for person in acc.contacts:
            name = person.name
            if fold_key(name) in held:
                self.report.contacts_present += 1
                continue
            try:
                contact = self._create_contact(acc, person)
                try:
                    self.client.link_employment(contact["id"], company_id)
                except ApiError as exc:
                    if exc.response.status_code != 409:
                        raise
            except ApiError as exc:
                self.report.failed.append({"id": f"{acc.source_id}#contact", "error": str(exc)[:300]})
                continue
            held.add(fold_key(name))
            self.report.contacts_created += 1

    def lead(self, lead: Lead, index: CrmIndex, held: set) -> None:
        if lead.source_id in held:
            self.report.leads_present += 1
            return
        if index.has_name(lead.company):
            self.report.leads_already_accounts += 1
            return
        try:
            status, made = self.client.post_status("/v1/leads", json=lead_body(lead))
        except ApiError as exc:
            if exc.response.status_code != 409:  # the CRM already holds this lead
                raise
            status = 409
        if status == 201:
            self.report.leads_created += 1
            values = {self.lead_columns[a]: getattr(lead, a) for a in self.lead_columns if getattr(lead, a)}
            if values:
                self.client.patch(f"/v1/leads/{made['id']}", values)
        else:
            self.report.leads_present += 1
        held.add(lead.source_id)

    def run(self) -> AccountsReport:
        self.vocabulary()
        self.emit("index")
        index = CrmIndex(self.client.paginate("/v1/companies", {"limit": 200}), self.columns.get("code"))
        known_notes = _note_ids(self.client) if any(a.notes for a in self.result.accepted_accounts()) else set()
        accounts = self.result.accepted_accounts()
        for i, acc in enumerate(accounts):
            try:
                self.account(acc, index, known_notes)
            except ApiError as exc:
                self.report.failed.append({"id": acc.source_id, "error": str(exc)[:300]})
            self.emit("accounts_progress", done=i + 1, total=len(accounts))
        held = _lead_ids(self.client)
        for lead in self.result.accepted_leads():
            try:
                self.lead(lead, index, held)
            except ApiError as exc:
                self.report.failed.append({"id": lead.source_id, "error": str(exc)[:300]})
        return self.report


def commit(client, result: AccountsResult, on_progress=None) -> dict:
    """Writes the accepted accounts, their contacts and notes, and the accepted leads. Returns the report as a dict."""
    emit = on_progress or (lambda step, **info: None)
    return asdict(_Run(client, result, emit).run())
