"""Writes a parsed, matched visit list into Margince, keeping every cell.

Where each column lands (README.md has the same table for a reader):

  customer, city            company (CSV import: name, address.city)
  potential, customer type  company custom fields "Potential", "Customer type",
                            plus lifecycle (existing -> customer, new -> prospect)
                            when the CRM does not know it yet
  owner (the rep)           author of every record written, the meeting's host
                            and the company owner when the rep holds a seat
  name, phone, office,      contact with real phone numbers and email addresses,
  email                     employed at the company
  date, purpose, notes      one meeting per visit (a note when there is nobody to
                            meet), timed on the visit's day, with the whole row
                            kept verbatim in the activity's `raw`
  green / yellow notes,     one task each, due on the date the deadline phrase
  deadline column           names, assigned to whoever it says will follow up
  key plan                  a "Week plan: planned for <day>" line on each visit
  key plan + summary        one weekly-report note per sheet, dated Friday of
                            that week and linked to no company

Every write carries `source_system` + a source_id built from what the record
says, not where it sat: `visit:<owner>:<date>:<customer>:<n>` (n = which of that
day's visits to that customer), tasks `<visit id>#T<hash of the note text>`, week
notes `week:<owner>:<year>-W<cw>`. A later file that repeats earlier weeks, or the
same file saved again, converges instead of duplicating. Records written under the
old file:sheet:row ids are found by those ids and by the row each one stored.
"""
import csv
import hashlib
import io
import re
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo

from . import when
from common.client import ApiError, MargenceClient
from common.fields import ensure_field
from excel.split import fold_key, phone_e164
from .match import Company, Contact, MatchResult
from .reader import Note, Parsed, Visit, Week

MEETING_STATUS = {"done": "held", "cancelled": "canceled", "internal": "held", "exhibition": "held"}
LIFECYCLE_BY_TYPE = {"existing": "customer", "new": "prospect"}

# (Company attribute, label, type, picklist options). Created on first run
# when the workspace has no active field with that label.
COMPANY_FIELDS = [
    ("potential", "Potential", "picklist", ["A", "B", "C"]),
    ("customer_type", "Customer type", "picklist", ["new", "existing", "dealer"]),
    ("sales_reps", "Sales reps", "text", None),
    ("first_visit", "First visit", "date", None),
    ("last_visit", "Last visit", "date", None),
]

# source_key must name one of the MAPPED columns (createImportRun rejects
# anything else with `unmapped_column`) - there is no separate id column, so
# the file's identifying column is the one already mapped to the natural key.
COMPANY_SOURCE_KEY = "name"
COMPANY_MAPPING = {"name": "display_name", "city": "address.city"}


@dataclass
class Options:
    source_system: str = "visitlist"
    timezone: str = "Asia/Shanghai"
    tasks: bool = True
    # Tasks due before this day are written already done: a historical file
    # would otherwise hand a rep a year of overdue reminders. None = the
    # Monday of the newest week in the file, i.e. "now" as the file sees it.
    close_tasks_due_before: str | None = None


@dataclass
class PipelineReport:
    companies_report: dict
    companies_updated: int = 0
    contacts_created: int = 0
    contacts_updated: int = 0
    activities_created: int = 0
    activities_idempotent: int = 0
    activities_skipped_no_date: list[str] = field(default_factory=list)
    tasks_created: int = 0
    tasks_closed: int = 0
    tasks_idempotent: int = 0
    weeks_created: int = 0
    weeks_idempotent: int = 0
    fields_created: list[str] = field(default_factory=list)
    reps_with_seat: dict[str, str] = field(default_factory=dict)
    reps_without_seat: list[str] = field(default_factory=list)


# ---- what gets written, independent of the network ----

def companies_csv(result: MatchResult, only: set[str] | None = None) -> str:
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=["name", "city"])
    writer.writeheader()
    writer.writerows({"name": c.display_name, "city": c.city} for k, c in result.companies.items()
                     if only is None or k in only)
    return buf.getvalue()


_CJK = re.compile(r"[\u4e00-\u9fff]")


def _crm_names(company: Company) -> list[str]:
    """Every name the company may already be filed under in the CRM, the Excel's spelling first."""
    out: list[str] = []
    for name in [company.display_name, *company.written_names, *company.list_names]:
        if name and name not in out:
            out.append(name)
    return out


def other_language_name(shown: str, written_names: list[str]) -> str:
    """The name the Excel files wrote in the other language than `shown` (Chinese against Latin), or "".
    Never translated: only a spelling that actually stands in a file."""
    for name in written_names:
        if bool(_CJK.search(name)) != bool(_CJK.search(shown)):
            return name
    return ""


def normalize_phone(raw: str) -> str:
    """The shared rule (excel.split.phone_e164): Chinese numbers without a country code become +86,
    a number with its own country code keeps it; anything it cannot place is returned as written."""
    return phone_e164(raw) or raw.strip()


def _existing_id(exc: ApiError, code: str) -> str | None:
    try:
        body = exc.response.json()
    except ValueError:
        return None
    if exc.response.status_code == 409 and body.get("code") == code:
        return (body.get("details") or {}).get("existing_id")
    return None


def contact_phones(contact: Contact) -> list[dict]:
    phones = []
    for raw, kind in [(p, "mobile") for p in contact.phones] + [(p, "work") for p in contact.office_phones]:
        number = normalize_phone(raw)
        # The CRM only takes numbers with a country prefix; one we cannot place (a
        # local number without area code, a WeChat ID) stays in the visit's text and raw row.
        if not number.startswith("+"):
            continue
        if number not in (p["phone"] for p in phones):
            phones.append({"phone": number, "phone_type": kind, "is_primary": not phones, "position": len(phones)})
    return phones


_EMAIL = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")


def contact_emails(contact: Contact) -> list[dict]:
    """Every address the cells hold (a cell can carry two, or text around one); text
    that holds no address stays in the visit's raw row only."""
    found = []
    for cell in contact.emails:
        for e in _EMAIL.findall(cell or ""):
            if e.lower() not in (f.lower() for f in found):
                found.append(e)
    return [{"email": e, "email_type": "work", "is_primary": i == 0, "position": i}
            for i, e in enumerate(found)]


def task_id(visit: Visit, note: Note) -> str:
    text = " ".join(note.text.lower().split())
    return f"{visit.source_id}#T{hashlib.sha1(text.encode()).hexdigest()[:10]}"


def _task_subject(note: Note) -> str:
    return ("Waiting — " if note.highlight == "yellow" else "") + note.text


_WEEK_SUBJECT = re.compile(r"^Weekly report CW(\d+) (\d{4}) — (.*)$")


class Existing:
    """The activities the CRM already holds for this source_system, findable by the id the
    current code would give them and by the older file:sheet:row ids."""

    def __init__(self, activities):
        self.by_id: dict[str, dict] = {}
        self.by_content: dict[str, dict] = {}
        self.tasks: dict[str, list[dict]] = defaultdict(list)  # visit id -> its tasks
        rows = []
        for a in activities:
            sid = a.get("source_id") or ""
            self.by_id[sid] = a
            if a.get("kind") == "task":
                if "#" in sid:
                    self.tasks[sid.split("#")[0]].append(a)
                continue
            raw = a.get("raw") or {}
            if row := raw.get("visit_list_row"):
                rows.append((row.get("workbook", ""), row.get("sheet", ""), row.get("excel_row_start", 0), row, a))
            elif week := raw.get("visit_list_week"):
                self.by_content.setdefault(f"week:{fold_key(week.get('owner'))}:{week.get('year')}-W{week.get('cw')}", a)
            elif m := _WEEK_SUBJECT.match(a.get("subject") or ""):
                self.by_content.setdefault(f"week:{fold_key(m.group(3))}:{m.group(2)}-W{m.group(1)}", a)
        seen: dict = {}
        for workbook, sheet, _, row, a in sorted(rows, key=lambda r: r[:3]):
            if not row.get("date_iso"):
                continue
            key = (workbook, sheet, fold_key(row.get("owner")), row["date_iso"], fold_key(row.get("customer_raw")))
            n = seen.get(key, 0)
            seen[key] = n + 1
            self.by_content.setdefault(f"visit:{key[2]}:{key[3]}:{key[4]}:{n}", a)

    def find(self, *ids: str) -> dict | None:
        for i in ids:
            if i in self.by_id:
                return self.by_id[i]
        return self.by_content.get(ids[0])

    def task(self, visit: Visit, note: Note, found: dict | None) -> dict | None:
        """The task already written for this note, under the new id or an old row-based one."""
        if hit := self.by_id.get(task_id(visit, note)):
            return hit
        for prefix in {visit.visit_key, (found or {}).get("source_id", "")}:
            for t in self.tasks.get(prefix, ()):
                if t.get("subject") == _task_subject(note):
                    return t
        return None

    def add(self, source_id: str, activity: dict) -> None:
        self.by_id[source_id] = activity


def task_notes(visit: Visit) -> list[Note]:
    """The note lines that ask somebody to do something: highlighted, or
    carrying a deadline. "Cancelled by customer." is an outcome, not a task."""
    return [n for n in visit.notes
            if (n.highlight or n.deadline) and "cancel" not in n.text.lower()]


def _fmt_day(day: date) -> str:
    return day.strftime("%a %d %b %Y")


def _note_line(note: Note, visit_day: date | None) -> str:
    tag = {"green": "[action] ", "yellow": "[waiting] "}.get(note.highlight, "")
    line = tag + note.text
    due = when.resolve_due(note.deadline, visit_day) if visit_day and note.deadline else None
    if note.deadline:
        line += f" — {('due ' + _fmt_day(due) + ', ') if due else ''}\"{note.deadline}\""
    return line


def visit_subject(visit: Visit, company: Company | None) -> str:
    name = company.display_name if company else visit.customer_raw
    where = f" ({visit.city})" if visit.city else ""
    if visit.kind == "exhibition":
        return f"Exhibition — {visit.customer_raw}{where}"
    if visit.kind == "internal":
        return f"Internal — {visit.purpose or visit.city}"
    prefix = "Cancelled visit" if visit.status == "cancelled" else "Visit"
    return f"{prefix} — {name}{where}"


def visit_body(visit: Visit, company: Company | None, planned: list[str] = ()) -> str:
    visit_day = date.fromisoformat(visit.date_iso) if visit.date_iso else None
    lines = [visit.purpose] if visit.purpose else []
    reach = [b for b in (visit.phone_raw, visit.office_phone and f"office {visit.office_phone}", visit.email) if b]
    if visit.contact_raw:
        lines.append(f"Met: {visit.contact_raw}" + (f" ({', '.join(reach)})" if reach else ""))
    facts = [f"Location: {visit.city}" if visit.city else "",
             f"ERP customer code: {visit.extra['customer_code']}" if visit.extra.get("customer_code") else ""]
    if visit.potential:
        facts.append(f"Potential: {visit.potential}")
    if visit.customer_type_raw:
        facts.append(f"Customer: {visit.customer_type_raw}")
    written = visit.extra.get("customer_as_written") or visit.customer_raw
    if company and written != company.display_name:
        facts.append(f"written as \"{written}\"")
    if any(facts):
        lines.append(" · ".join(f for f in facts if f))
    if planned:
        days = " and ".join(_fmt_day(date.fromisoformat(d)) for d in planned)
        on_plan = visit.date_iso in planned
        lines.append(f"Week plan: planned for {days}" + ("" if on_plan or not visit_day
                                                          else f", visited {_fmt_day(visit_day)}"))
    if visit.notes:
        lines.append("")
        lines.append("Next steps:")
        lines.extend("• " + _note_line(n, visit_day) for n in visit.notes)
    lines.append("")
    lines.append(f"From {visit.owner or 'the'}'s visit list, CW{visit.cw} ({visit.date_raw}, time of day "
                 f"not recorded) — {visit.workbook} › {visit.sheet} › rows "
                 f"{visit.excel_row_start}–{visit.excel_row_end}")
    return "\n".join(lines).strip()


def week_body(week: Week) -> str:
    lines = ["Key plan"]
    for day in week.days:
        lines.append(f"• {_fmt_day(date.fromisoformat(day.date_iso))}: {day.plan or '—'}")
    if week.summary:
        lines.append("")
        lines.append("Summary")
        lines.extend("• " + s for s in week.summary)
    lines.append("")
    lines.append(f"From {week.owner}'s visit list — {week.workbook} › {week.sheet}")
    return "\n".join(lines)


def planned_days(week: Week, result: MatchResult) -> dict[str, list[str]]:
    """Company key -> the days the key plan ("City: Name; City: Name") put
    it on, so each visit can say whether it went as planned."""
    plan: dict[str, list[str]] = {}
    for day in week.days:
        for entry in day.plan.split(";"):
            city, sep, name = entry.partition(":")
            key = result.resolve_company(name.strip(), city.strip()) if sep else None
            if key and day.date_iso not in plan.setdefault(key, []):
                plan[key].append(day.date_iso)
    return plan


def default_task_cutoff(parsed: Parsed) -> str | None:
    mondays = [w.monday_iso for w in parsed.weeks]
    return max(mondays) if mondays else None


def preview_companies(client: MargenceClient, result: MatchResult) -> dict:
    """The company CSV run a commit would make, previewed and never approved."""
    by_name = {fold_key(n) for n in client.company_id_by_name()}
    missing = {k for k, c in result.companies.items() if not any(fold_key(n) in by_name for n in _crm_names(c))}
    if not missing:
        return {"disposition": {"already_present": len(result.companies)}, "rows_read": 0, "issues": []}
    source_ref = client.upload_csv("company", companies_csv(result, missing), "companies.csv")
    run_id = client.create_import_run("company", source_ref, COMPANY_MAPPING, COMPANY_SOURCE_KEY, "create")
    report = client.wait_for_report(run_id)
    report["disposition"]["already_present"] = len(result.companies) - len(missing)
    return report


def plan_counts(parsed: Parsed, result: MatchResult) -> dict:
    """What a commit would write, counted locally: the preview's half that
    needs no server."""
    contacts = list(result.contacts.values())
    return {
        "visits": len(parsed.visits),
        "meetings": sum(1 for v in parsed.visits if v.date_iso and result.links.get(v.visit_key, (None, None))[1]),
        "notes": sum(1 for v in parsed.visits if v.date_iso and not result.links.get(v.visit_key, (None, None))[1]),
        "contacts": len(contacts),
        "contacts_with_phone": sum(1 for c in contacts if c.phones or c.office_phones),
        "contacts_with_email": sum(1 for c in contacts if c.emails),
        "tasks": sum(len(task_notes(v)) for v in parsed.visits if v.date_iso),
        "weeks": len(parsed.weeks),
        "task_cutoff": default_task_cutoff(parsed),
    }


# ---- the run ----

class _Run:
    def __init__(self, client: MargenceClient, parsed: Parsed, result: MatchResult, opts: Options, emit):
        self.client, self.parsed, self.result, self.opts, self.emit = client, parsed, result, opts, emit
        self.tz = ZoneInfo(opts.timezone)
        self.report = PipelineReport(companies_report={})
        self.started = datetime.now().astimezone()
        self.company_ids: dict[str, str] = {}  # company key -> id
        self.contact_ids: dict[str, str] = {}  # contact key -> id
        self.seats: dict[str, str] = {}  # folded rep name -> seat id
        self.field_columns: dict[str, str] = {}  # Company attribute -> cf_ column
        self.me = client.get("/v1/me")["user"]["id"]
        self.plans = {(w.workbook, w.sheet): planned_days(w, result) for w in parsed.weeks}
        self.existing = Existing(())

    def seat_for(self, name: str) -> str | None:
        return self.seats.get(when.fold_person(name)) if name else None

    def ensure_vocabulary(self) -> None:
        self.emit("vocabulary")
        self.seats = {when.fold_person(n): i for n, i in self.client.seats_by_name().items()}
        reps = sorted({v.owner for v in self.parsed.visits if v.owner})
        self.report.reps_with_seat = {r: self.seat_for(r) for r in reps if self.seat_for(r)}
        self.report.reps_without_seat = [r for r in reps if not self.seat_for(r)]
        for attr, label, type_, options in COMPANY_FIELDS:
            before = {f["label"] for f in self.client.custom_fields("company")}
            found = ensure_field(self.client, "company", label, type_, options, "visits")
            if found["label"] not in before:
                self.report.fields_created.append(found["label"])
            self.field_columns[attr] = found["column_name"]

    def company_ids_by_name(self) -> dict[str, str]:
        by_name = {fold_key(n): i for n, i in self.client.company_id_by_name().items()}
        found = {}
        for k, c in self.result.companies.items():
            hit = next((n for n in _crm_names(c) if fold_key(n) in by_name), None)
            if hit:
                found[k] = by_name[fold_key(hit)]
        return found

    def import_companies(self) -> None:
        """Only names the CRM does not hold EXACTLY go through the CSV import,
        and they go with on_duplicate=create. The importer's own duplicate
        check is a fuzzy name ladder built to queue pairs for a human, and on
        Chinese company names - which share "Intelligent Equipment", "Tech",
        "Automation" - it calls most of a visit list a duplicate of the rest;
        the spellings of one company were already folded by match.py. The
        near-misses still reach the CRM's duplicate review queue."""
        self.emit("companies_import_start")
        self.company_ids = self.company_ids_by_name()
        missing = set(self.result.companies) - set(self.company_ids)
        self.report.companies_report = {"disposition": {"already_present": len(self.company_ids)},
                                        "rows_read": 0, "issues": []}
        if missing:
            report = self.client.run_csv_import(
                "company", companies_csv(self.result, missing), "companies.csv", mapping=COMPANY_MAPPING,
                source_key=COMPANY_SOURCE_KEY, on_duplicate="create")
            report["disposition"]["already_present"] = len(self.company_ids)
            self.report.companies_report = report
            self.company_ids = self.company_ids_by_name()
        self.emit("companies_import_done", report=self.report.companies_report)
        for i, (key, company) in enumerate(self.result.companies.items()):
            if key in self.company_ids:
                self.report.companies_updated += self.enrich_company(self.company_ids[key], company)
            self.emit("companies_enrich", done=i + 1, total=len(self.result.companies))

    def enrich_company(self, company_id: str, company: Company) -> int:
        current = self.client.get_company(company_id)
        values = {"potential": company.potential, "customer_type": company.customer_type,
                  "sales_reps": ", ".join(company.owners), "first_visit": company.first_visit,
                  "last_visit": company.last_visit}
        patch = {}
        # An older file must not overwrite what a newer one already said.
        stored_last = current.get(self.field_columns["last_visit"]) or ""
        if company.last_visit >= stored_last:
            patch = {self.field_columns[a]: v for a, v in values.items() if v}
        stored_first = current.get(self.field_columns["first_visit"])
        if stored_first and stored_first < company.first_visit:
            patch.pop(self.field_columns["first_visit"], None)
        if current.get("lifecycle") in (None, "unknown") and company.customer_type in LIFECYCLE_BY_TYPE:
            patch["lifecycle"] = LIFECYCLE_BY_TYPE[company.customer_type]
        rep_seat = self.seat_for(company.owners[0]) if company.owners else None
        created_at = current.get("created_at")
        created_now = bool(created_at) and datetime.fromisoformat(created_at) >= self.started
        if rep_seat and current.get("owner_id") == self.me and created_now:
            patch["owner_id"] = rep_seat
        # The other language's name goes next to the one the CRM shows ("legal name", which the CRM
        # search reads); one a person typed is never replaced.
        other = other_language_name(current.get("display_name") or "", company.written_names)
        if other and not current.get("legal_name"):
            patch["legal_name"] = other
        patch = {k: v for k, v in patch.items() if current.get(k) != v}
        if not patch:
            return 0
        self.client.patch(f"/v1/companies/{company_id}", patch)
        return 1

    def import_contacts(self) -> None:
        self.emit("contacts_start", total=len(self.result.contacts))
        by_company: dict[str, dict[str, str]] = {}
        for i, (key, contact) in enumerate(self.result.contacts.items()):
            company_id = self.company_ids.get(contact.company_key)
            if not company_id:
                continue
            if company_id not in by_company:
                by_company[company_id] = {fold_key(c["full_name"]): c["contact_id"]
                                          for c in self.client.contacts_for_company(company_id)}
            existing_id = by_company[company_id].get(fold_key(contact.full_name))
            if existing_id:
                self.contact_ids[key] = existing_id
                self.report.contacts_updated += self.add_reachability(existing_id, contact)
            else:
                self.contact_ids[key], created = self.create_contact(contact, company_id)
                by_company[company_id][fold_key(contact.full_name)] = self.contact_ids[key]
                self.report.contacts_created += created
                self.report.contacts_updated += not created  # an existing person, found by email
            self.emit("contacts_progress", done=i + 1, total=len(self.result.contacts))

    def create_contact(self, contact: Contact, company_id: str) -> tuple[str, bool]:
        """(contact id, whether a new record was made)."""
        body = {"full_name": contact.full_name, "source": "import", "source_system": self.opts.source_system,
                "phones": contact_phones(contact), "emails": contact_emails(contact)}
        if contact.owners:
            body["source_author_name"] = contact.owners[0]
            if seat := self.seat_for(contact.owners[0]):
                body["owner_id"] = seat
        created = True
        try:
            contact_id = self.client.create_contact(body)["id"]
        except ApiError as exc:
            existing = _existing_id(exc, "duplicate_email")
            if not existing:
                raise
            held = self.client.get_contact(existing)
            if fold_key(held.get("full_name")) == fold_key(contact.full_name):
                # The same person met at two companies (or written two ways): that record is
                # the contact, linked here too, and it learns the numbers this sheet adds.
                contact_id, created = existing, False
                self.add_reachability(existing, contact)
            else:
                # Somebody else's address (a shared mailbox, a colleague's): this person is
                # created without it; the address stays in the visit's text.
                taken = {e["email"].lower() for e in held.get("emails") or []}
                free = [e for e in body["emails"] if e["email"].lower() not in taken]
                body["emails"] = [{**e, "is_primary": i == 0, "position": i} for i, e in enumerate(free)]
                try:
                    contact_id = self.client.create_contact(body)["id"]
                except ApiError as again:
                    if not _existing_id(again, "duplicate_email"):
                        raise
                    body["emails"] = []
                    contact_id = self.client.create_contact(body)["id"]
        try:
            self.client.link_employment(contact_id, company_id)
        except ApiError as exc:
            if exc.response.status_code != 409:  # already employed there: nothing to add
                raise
        return contact_id, created

    def add_reachability(self, contact_id: str, contact: Contact) -> int:
        """Adds the numbers and addresses the sheet knows and the record does
        not; never removes or reorders what is already there."""
        current = self.client.get_contact(contact_id)
        patch = {}
        have_phones = [{k: p[k] for k in ("phone", "phone_type", "is_primary", "position") if k in p}
                       for p in current.get("phones") or []]
        new_phones = [p for p in contact_phones(contact) if p["phone"] not in {h["phone"] for h in have_phones}]
        if new_phones:
            patch["phones"] = have_phones + [{**p, "is_primary": not have_phones and i == 0,
                                              "position": len(have_phones) + i} for i, p in enumerate(new_phones)]
        have_emails = [{k: e[k] for k in ("email", "email_type", "is_primary", "position") if k in e}
                       for e in current.get("emails") or []]
        known = {h["email"].lower() for h in have_emails}
        new_emails = [e for e in contact_emails(contact) if e["email"].lower() not in known]
        if new_emails:
            patch["emails"] = have_emails + [{**e, "is_primary": not have_emails and i == 0,
                                              "position": len(have_emails) + i} for i, e in enumerate(new_emails)]
        if not patch:
            return 0
        try:
            self.client.patch(f"/v1/contacts/{contact_id}", patch)
        except ApiError as exc:
            if not _existing_id(exc, "duplicate_email"):
                raise
            # Another contact holds one of the addresses (this person was made without it):
            # keep the numbers, leave the address in the visit's text, and do not ask again as a change.
            patch.pop("emails", None)
            if not patch:
                return 0
            self.client.patch(f"/v1/contacts/{contact_id}", patch)
        return 1

    def log(self, payload: dict) -> tuple[bool, dict]:
        try:
            return self.client.log_activity(payload)
        except ApiError as e:
            # The source_id is already held under another kind (an earlier run
            # filed this visit as a note before its contact was known). The
            # held record wins; asking again without the meeting-only fields
            # returns it as the idempotent replay it is.
            if e.response.status_code != 422 or "meeting_status" not in payload:
                raise
            retry = {k: v for k, v in payload.items() if k not in ("meeting_status", "host_user_id")}
            return self.client.log_activity(retry)

    def load_existing(self) -> None:
        self.emit("existing_activities")
        self.existing = Existing(a for a in self.client.paginate("/v1/activities", {"limit": 200})
                                 if a.get("source_system") == self.opts.source_system)

    def write(self, payload: dict, found: dict | None) -> tuple[bool, dict]:
        """(created, activity): nothing is sent when the CRM already holds the record."""
        if found:
            return False, found
        created, activity = self.log(payload)
        self.existing.add(payload["source_id"], activity)
        return created, activity

    def import_visits(self) -> None:
        datable = [v for v in self.parsed.visits if v.date_iso]
        self.report.activities_skipped_no_date = [v.visit_key for v in self.parsed.visits if not v.date_iso]
        cutoff = self.opts.close_tasks_due_before or default_task_cutoff(self.parsed)
        self.emit("activities_start", total=len(datable))
        for i, visit in enumerate(datable):
            company_key, contact_key = self.result.links.get(visit.visit_key, (None, None))
            company = self.result.companies.get(company_key) if company_key else None
            company_id = self.company_ids.get(company_key) if company_key else None
            contact_id = self.contact_ids.get(contact_key) if contact_key else None
            found = self.existing.find(visit.source_id, visit.visit_key)
            created, _ = self.write(self.visit_payload(visit, company, company_id, contact_id), found)
            self.report.activities_created += created
            self.report.activities_idempotent += not created
            if self.opts.tasks:
                for note in task_notes(visit):
                    self.log_task(visit, note, company, company_id, contact_id, cutoff, found)
            self.emit("activity_progress", done=i + 1, total=len(datable),
                      created=self.report.activities_created, idempotent=self.report.activities_idempotent)

    def common(self, source_id: str, author: str) -> dict:
        body = {"source": "import", "source_system": self.opts.source_system, "source_id": source_id}
        if author:
            body["source_author_name"] = author
            if seat := self.seat_for(author):
                body["source_author_id"] = seat
        return body

    def visit_payload(self, visit: Visit, company: Company | None, company_id: str | None,
                      contact_id: str | None) -> dict:
        # A meeting is WITH A HUMAN: the API refuses a company link on one
        # (activitylinks.go, CompanyMeetingError) and reaches the company
        # through the contact's employer. With nobody to meet, the visit is
        # filed as a note, which can name the company directly.
        kind = "meeting" if contact_id else "note"
        links = ([{"entity_type": "contact", "entity_id": contact_id}] if contact_id else
                 [{"entity_type": "company", "entity_id": company_id}] if company_id else [])
        payload = {
            **self.common(visit.source_id, visit.owner),
            "kind": kind,
            "subject": visit_subject(visit, company),
            "body": visit_body(visit, company, self.plans.get((visit.workbook, visit.sheet), {})
                               .get(company.key if company else "", [])),
            "occurred_at": when.visit_instant(visit, self.tz).isoformat(),
            "links": links,
            "raw": {"visit_list_row": asdict(visit), "legacy_key": visit.visit_key, "status": visit.status,
                    "company": company.display_name if company else None,
                    "time_of_day_recorded": False},
        }
        if kind == "meeting":
            payload["meeting_status"] = MEETING_STATUS.get(visit.status, "held")
            if seat := self.seat_for(visit.owner):
                payload["host_user_id"] = seat
        return payload

    def task_payload(self, visit: Visit, note: Note, company: Company | None, company_id: str | None,
                     contact_id: str | None) -> dict:
        visit_day = date.fromisoformat(visit.date_iso)
        due = when.resolve_due(note.deadline, visit_day) or when.resolve_due(note.text, visit_day)
        follower = when.follow_up_owner(note.deadline) or visit.owner
        about = company.display_name if company else visit.customer_raw
        lines = [f"From the visit to {about} on {_fmt_day(visit_day)} ({visit.owner}, CW{visit.cw})."]
        if note.highlight == "yellow":
            lines.append("Marked yellow on the sheet: waiting on the customer.")
        if note.deadline:
            lines.append(f"Deadline as written: \"{note.deadline}\"")
        payload = {
            **self.common(task_id(visit, note), visit.owner),
            "kind": "task",
            "subject": _task_subject(note),
            "body": "\n".join(lines),
            "occurred_at": when.visit_instant(visit, self.tz).isoformat(),
            "links": [l for l in ([{"entity_type": "contact", "entity_id": contact_id}] if contact_id else []) +
                      ([{"entity_type": "company", "entity_id": company_id}] if company_id else [])],
        }
        if due:
            payload["due_at"] = when.day_instant(due.isoformat(), self.tz).isoformat()
        if seat := self.seat_for(follower):
            payload["assignee_id"] = seat
        return payload

    def log_task(self, visit: Visit, note: Note, company: Company | None, company_id: str | None,
                 contact_id: str | None, cutoff: str | None, found: dict | None = None) -> None:
        visit_day = date.fromisoformat(visit.date_iso)
        due = when.resolve_due(note.deadline, visit_day) or when.resolve_due(note.text, visit_day)
        payload = self.task_payload(visit, note, company, company_id, contact_id)
        created, task = self.write(payload, self.existing.task(visit, note, found))
        if not created:
            self.report.tasks_idempotent += 1
            return
        self.report.tasks_created += 1
        settled_by = (due or visit_day).isoformat()
        if cutoff and settled_by < cutoff and not task.get("is_done"):
            self.client.patch(f"/v1/activities/{task['id']}", {"is_done": True})
            self.report.tasks_closed += 1

    def import_weeks(self) -> None:
        self.emit("weeks_start", total=len(self.parsed.weeks))
        for i, week in enumerate(self.parsed.weeks):
            # The rep's own report on the whole week, so it is linked to no
            # company: on a company's page it would read as news about that
            # company. Each visit already says which day the plan put it on.
            # Dated at the end of the week it reports on.
            friday = (date.fromisoformat(week.monday_iso) + timedelta(days=4)).isoformat()
            created, _ = self.write({
                **self.common(week.source_id, week.owner),
                "kind": "note",
                "subject": f"Weekly report CW{week.cw} {week.year} — {week.owner}",
                "body": week_body(week),
                "occurred_at": when.day_instant(friday, self.tz, time(17, 0)).isoformat(),
                "links": [],
                "raw": {"visit_list_week": asdict(week), "legacy_key": week.week_key},
            }, self.existing.find(week.source_id, week.week_key))
            self.report.weeks_created += created
            self.report.weeks_idempotent += not created
            self.emit("weeks_progress", done=i + 1, total=len(self.parsed.weeks))


def run(client: MargenceClient, parsed: Parsed, result: MatchResult, opts: Options | None = None,
        on_progress=None) -> PipelineReport:
    """on_progress(step, **info), called at each stage - the web UI's job
    tracker is the only caller that passes one; the CLI doesn't."""
    def emit(step, **info):
        if on_progress:
            on_progress(step, **info)

    r = _Run(client, parsed, result, opts or Options(), emit)
    r.ensure_vocabulary()
    r.import_companies()
    r.import_contacts()
    r.load_existing()
    r.import_visits()
    r.import_weeks()
    emit("done", report=r.report)
    return r.report
