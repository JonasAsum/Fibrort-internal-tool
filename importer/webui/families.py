"""One adapter per list family (the visit reports are handled in server.py, which owns their state).
Each adapter turns a family's module (plain records first, CRM mapping second) into the same few
calls the web UI needs: read, summary counts, notes, company questions, flags (one-click include
or leave out), a dry run, and a commit whose report has one shape for every family.

Nothing here writes to the CRM except `commit`.
"""
import os
from collections import Counter
from decimal import Decimal

from excel import accounts, accounts_margince, contacts, contacts_margince, orders, orders_margince
from excel import quotes, quotes_margince, receivables, receivables_margince, todos, todos_margince
from excel.companies import norm as company_norm

# Plain words for the review reasons that are notes or skipped rows (company questions are asked, not listed).
REASON_TEXT = {
    "section": "heading rows that are not people (skipped)",
    "section_heading": "heading rows (skipped)",
    "phone_no_country": "phone numbers without a country code left out (no local number is guessed)",
    "bad_email": "e-mail addresses that are not valid left out",
    "no_company": "rows without a company name (skipped)",
    "no_contact_name": "rows without a person's name (skipped)",
    "withheld_sensitive": "text that looks like bank or tax details withheld (stays out of the CRM)",
    "id_number_in_text_field": "long ID-like numbers in text fields dropped (tax and registration numbers stay out)",
    "differs_between_copies": "values that differ between two copies of the same list (the richer copy is used)",
    "do_not_use": "accounts the ERP marks 'Do Not use' (imported with a note, not as customers)",
    "code_cell_has_text": "customer-code cells that hold text instead of a code",
    "conflict": "values in the files that differ from each other or from the CRM (the CRM value is kept)",
    "name_differs": "one ERP code written with different names in two lists",
    "same_name_other_code": "one company under two ERP codes (two accounts, told apart by the code)",
    "owner_unknown": "owners that are not one of the sales reps (task left unassigned)",
    "owner_placeholder": "owner left as a placeholder such as 'XXX' (task left unassigned)",
    "it_project_list": "the JET-IT list looks like an IT project list, not sales: check it",
    "no_text": "rows without any text (skipped)",
    "no_tasks": "lists with no task in them (blank template)",
    "unknown_status": "status words that are not known (task left open)",
    "ambiguous_due": "due dates that could be read two ways (task has no due date, text kept)",
    "due_decimal": "due dates written like 5.25 (read as month.day)",
    "ambiguous_date": "dates that could be read two ways (left out)",
    "unparsed_date": "dates that could not be read (left out)",
    "company_ambiguous": "task texts that name several possible customers (not linked)",
    "it_list_excluded": "rows of the IT list left out at your request",
    "currency_not_stated": "orders where the export names no currency (never added up)",
    "no_date": "orders without a date (not imported)",
    "blank_order": "orders with no content (skipped)",
    "group_company": "orders and receivables of a FIBRO group company (intercompany, not a customer)",
    "no_due_date": "invoices without a due date (counted as open, never overdue)",
    "invoice_doubtful": "customers held out: an invoice without a number, or one number listed twice",
    "order_doubtful": "orders held out: lines that do not add up cleanly, or a FIBRO group company",
}


def money(x) -> str:
    return f"{Decimal(x):,.2f}"


def notes_from(review, skip=("company",)) -> list[str]:
    count = Counter(i.reason for i in review if i.reason not in skip)
    return [f"{n} {REASON_TEXT.get(r, r.replace('_', ' '))}" for r, n in sorted(count.items(), key=lambda t: (-t[1], t[0]))]


SUGGEST_MIN = 0.85  # a top candidate this close is offered as a one-click "same company"


def with_suggestion(q: dict, master) -> dict:
    """Adds q["suggest"] (the top candidate) when it is close enough to be one click, never a silent merge."""
    top = (q.get("candidates") or [None])[0]
    q["suggest"] = top if top and top["score"] >= SUGGEST_MIN and top["code"] and master.has_code(top["code"]) else None
    return q


def questions_from(review, unit: str) -> list[dict]:
    """One question per distinct company name the family could not place ("same company?")."""
    from excel import run as excel_run
    master = excel_run._master(excel_run.IMPORTER_DIR)
    asked: dict[str, dict] = {}
    for item in review:
        if item.reason != "company":
            continue
        q = asked.setdefault(company_norm(item.value) or item.value, {
            "kind": "company", "name": item.value, "count": 0, "unit": unit, "candidates": [
                {"code": c[0], "name": c[1], "en": (master.get(c[0]) or {}).get("en", ""), "score": round(c[2], 2)}
                for c in (item.candidates or [])[:3]]})
        q["count"] += 1
    return [with_suggestion(q, master) for q in asked.values()]


def rows(*pairs) -> list[dict]:
    """[(label, value)] -> [{"label", "value"}]; None values are left out."""
    return [{"label": label, "value": value} for label, value in pairs if value is not None]


def report(written, present, failed=0, skipped=0, lines=(), notes=(), errors=()) -> dict:
    """The one shape of a commit result: what was written, was already there, was skipped, failed."""
    status = "failed" if failed and not written else "partial" if failed else "done" if written else "nothing_new"
    return {"status": status, "written": written, "present": present, "skipped": skipped, "failed": failed,
            "rows": rows(*lines), "notes": list(notes), "errors": list(errors)[:10]}


class Adapter:
    key = ""
    unit = "rows"
    flags: tuple = ()  # (id, label, help)
    left_out_by_default: tuple = ()  # flagged groups that are not known to be right: out until a person ticks them in

    def read(self, paths):
        raise NotImplementedError

    def apply(self, full, left_out: set):
        """The result with the flagged groups in `left_out` removed."""
        return full

    def flag_counts(self, full) -> dict:
        return {}

    def counts(self, result) -> list[dict]:
        raise NotImplementedError

    def writable(self, result) -> int:
        """How many records are ready (a commit with 0 has nothing to do)."""
        raise NotImplementedError

    def waiting(self, result) -> int:
        return 0

    def notes(self, result) -> list[str]:
        return notes_from(result.review)

    def questions(self, result) -> list[dict]:
        return questions_from(result.review, self.unit)

    def file_stats(self, result) -> dict:
        """{saved path: (records, unit)}"""
        return {}

    def file_notes(self, result) -> dict:
        """{saved path: (status, plain words)} for a file whose records are all or partly covered by another file."""
        return {}

    def skipped(self, result) -> dict:
        """{saved path: why it was not read}"""
        return {s["file"]: s["error"] for s in result.skipped_files}

    def dry_run(self, client, result) -> list[dict]:
        raise NotImplementedError

    def commit(self, client, result, emit) -> dict:
        raise NotImplementedError


class Accounts(Adapter):
    key = "accounts"
    unit = "list rows"

    def read(self, paths):
        return accounts.read_paths(paths)

    def notes(self, r):
        # columns in the file that the importer does not read: shown so nothing is silently missing
        out = notes_from(r.review)
        for s in r.per_file_stats:
            if s.unread:
                out.append(f"Not read from {os.path.basename(s.path)}: {', '.join(s.unread)}")
        return out

    def writable(self, r):
        return len(r.accepted_accounts()) + len(r.accepted_leads())

    def waiting(self, r):
        return (len(r.accounts) - len(r.accepted_accounts())) + (len(r.leads) - len(r.accepted_leads()))

    def counts(self, r):
        p = accounts_margince.plan_counts(r)
        return rows(("Accounts", p["accounts"]), ("Customers", p["customers"]), ("Prospects", p["prospects"]),
                    ("Competitors", p["competitors"]), ("With competitor, forecast or application", p["with_opportunity_data"]),
                    ("Contacts", p["contacts"]), ("Notes", p["notes"]), ("Leads", p["leads"]),
                    ("Waiting for a company answer", p["queued"]))

    def file_stats(self, r):
        return {s.path: (s.rows, "rows") for s in r.per_file_stats}

    def dry_run(self, client, r):
        p = accounts_margince.preview(client, r)
        lines = [("New accounts", p["companies_new"]), ("Accounts already in the CRM", p["companies_present"]),
                 ("Existing accounts that get empty fields filled", p["companies_to_fill"]),
                 ("Differing values left as they are", p["conflicts"]), ("Accounts the CRM holds twice (skipped)", p["companies_ambiguous"]),
                 ("New contacts", p["contacts_new"]), ("New notes", p["notes_new"]), ("New leads", p["leads_new"]),
                 ("Leads that are already accounts", p["leads_already_accounts"])]
        if p["fields_missing"]:
            lines.append(("Custom fields that will be created", ", ".join(p["fields_missing"])))
        return rows(*lines)

    def commit(self, client, r, emit):
        c = accounts_margince.commit(client, r, emit)
        written = c["companies_created"] + c["companies_updated"] + c["contacts_created"] + c["notes_created"] + c["leads_created"]
        present = c["companies_unchanged"] + c["contacts_present"] + c["notes_idempotent"] + c["leads_present"]
        notes = [f"{len(c['conflicts'])} differing values were left as they are in the CRM"] if c["conflicts"] else []
        if c["companies_ambiguous"]:
            notes.append(f"{len(c['companies_ambiguous'])} accounts exist twice in the CRM and were skipped")
        return report(written, present, len(c["failed"]), lines=[
            ("Accounts created", c["companies_created"]), ("Accounts filled in", c["companies_updated"]),
            ("Accounts already complete", c["companies_unchanged"]), ("Contacts created", c["contacts_created"]),
            ("Notes created", c["notes_created"]), ("Leads created", c["leads_created"]),
            ("Custom fields created", ", ".join(c["fields_created"]) or None)],
            notes=notes, errors=[f"{f['id']}: {f['error']}" for f in c["failed"]])


def _contact_group(origin: str) -> str:
    return "finance" if origin.startswith("Finance list") else "internal" if "internal" in origin else ""


class Contacts(Adapter):
    key = "contacts"
    unit = "list rows"
    flags = (("finance", "Finance-list contacts and companies",
              "From the Chinese receivable/payable lists: customers and suppliers. Tagged 'Finance list' in the CRM."),
             ("internal", "Internal FIBRO colleagues",
              "From the handover list: kept on their own account '... (internal)', tagged 'Handover list: internal colleague'."))

    def read(self, paths):
        return contacts.read_paths(paths)

    def apply(self, full, left_out):
        if not left_out:
            return full
        keep = [c for c in full.contacts if _contact_group(c.origin) not in left_out]
        used = {c.company_key for c in keep}
        orgs = {k: o for k, o in full.orgs.items() if k in used or _contact_group(o.origin) not in left_out}
        dropped = {n for k, o in full.orgs.items() if k not in orgs for n in o.written_names}
        review = [i for i in full.review if not (i.reason == "company" and i.value in dropped)]
        return contacts.ContactsResult(contacts=keep, orgs=orgs, review=review, skipped_files=full.skipped_files,
                                       per_file_stats=full.per_file_stats)

    def flag_counts(self, full):
        n = Counter(_contact_group(c.origin) for c in full.contacts)
        return {"finance": n["finance"], "internal": n["internal"]}

    def writable(self, r):
        return len(r.accepted_contacts()) + len(r.accepted_orgs())

    def waiting(self, r):
        return len(r.contacts) - len(r.accepted_contacts())

    def counts(self, r):
        s = contacts.summarize(r)
        return rows(("Contacts", s["contacts_accepted"]), ("Companies", s["companies_accepted"]),
                    ("Waiting for a company answer", s["contacts_queued"]))

    def file_stats(self, r):
        return {s.path: (0 if s.blank_rows >= s.rows else s.rows, "rows") for s in r.per_file_stats}

    def dry_run(self, client, r):
        plan, extras = contacts_margince.contacts_plan(r)
        p = contacts_margince.contacts_preview(client, plan)
        return rows(("New companies", p["companies_new"]), ("Companies already in the CRM", p["companies_present"]),
                    ("New contacts", p["contacts_new"]), ("Contacts already in the CRM", p["contacts_present"]),
                    ("Job titles that will be set", sum("title" in e for e in extras.values())))

    def commit(self, client, r, emit):
        c = contacts_margince.contacts_commit(client, r, on_progress=emit)
        d = c["companies"] or {}
        created = d.get("created", 0)
        written = created + c["companies_enriched"] + c["contacts_created"] + c["contacts_updated"] + c["titles_set"] + c["fields_set"]
        notes = [f"{c['contacts_without_company']} contacts have no company in the CRM and were not written"] if c["contacts_without_company"] else []
        return report(written, d.get("already_present", 0), 0, lines=[
            ("Companies created", created), ("Companies enriched", c["companies_enriched"]),
            ("Contacts created", c["contacts_created"]), ("Contacts given phones or e-mails", c["contacts_updated"]),
            ("Job titles set", c["titles_set"]), ("Other fields set (origin, department, remark)", c["fields_set"])], notes=notes)


class Todos(Adapter):
    key = "todos"
    unit = "task rows"
    flags = (("it_list", "JET-IT project list",
              "Looks like an IT project list, not sales. Tasks carry the tag 'needs-check: IT project list'."),
             ("template", "Sales-topics template tasks",
              "A template that names no sales rep. Tasks carry the tag 'needs-check: topics template'."))
    TAGS = {"it_list": todos.IT_TAG, "template": todos.TEMPLATE_TAG}
    left_out_by_default = ("it_list", "template")

    def read(self, paths):
        return todos.read_paths(paths)

    def apply(self, full, left_out):
        if not left_out:
            return full
        gone = {self.TAGS[f] for f in left_out}
        return todos.TodosResult(todos=[t for t in full.todos if not gone & set(t.tags)], review=full.review,
                                 skipped_files=full.skipped_files, per_file_stats=full.per_file_stats)

    def flag_counts(self, full):
        return {f: sum(tag in t.tags for t in full.todos) for f, tag in self.TAGS.items()}

    def writable(self, r):
        return len(r.todos)

    def counts(self, r):
        s = todos.summarize(r)
        return rows(("Tasks", s["tasks"]), ("Done", s["done"]), ("Open", s["open"]), ("With a due date", s["with_due"]),
                    ("With a known sales rep as owner", s["with_owner"]), ("Linked to a known customer", s["linked_to_company"]),
                    ("Tagged 'needs-check'", s["needs_check"]))

    def questions(self, r):
        return []  # tasks are linked only when the match is certain; nothing is asked

    def file_stats(self, r):
        return {s.path: (s.tasks, "tasks") for s in r.per_file_stats}

    def dry_run(self, client, r):
        p = todos_margince.preview(client, r)
        lines = [("New tasks", p["new"]), ("Tasks already in the CRM", p["present"]),
                 ("New tasks that are already done", p["would_close"]), ("Tasks linked to an account", p["linked"]),
                 ("Customer named but not in the CRM (task not linked)", p["company_not_in_crm"]),
                 ("Tasks assigned to a rep", p["assigned"]), ("Tasks without an assignee", p["unassigned"])]
        if p["reps_without_seat"]:
            lines.append(("Reps without a CRM seat", ", ".join(p["reps_without_seat"])))
        return rows(*lines)

    def commit(self, client, r, emit):
        c = todos_margince.commit(client, r, on_progress=emit)
        return report(c["tasks_created"], c["tasks_present"], len(c["failed"]), lines=[
            ("Tasks created", c["tasks_created"]), ("Tasks already in the CRM", c["tasks_present"]),
            ("Created as done", c["tasks_closed"]), ("Linked to an account", c["linked"]),
            ("Assigned to a rep", c["assigned"])], errors=[f"{f['source_id']}: HTTP {f['status']}" for f in c["failed"]])


class Orders(Adapter):
    key = "orders"
    unit = "orders"
    flags = (("no_currency", "Orders with no currency named",
              "The export names no currency for these (the US entity). They are never added up with the others."),)

    def read(self, paths):
        return orders.build_book(paths)

    def apply(self, full, left_out):
        if not left_out:
            return full
        kept = [o for o in full.orders if o.currency is not None]
        used = {o.company_key for o in kept}
        return orders.OrdersBook(orders=kept, orgs={k: g for k, g in full.orgs.items() if k in used},
                                 review=[i for i in full.review if i.reason != "company" or any(
                                     i.value in g.written_names for g in full.orgs.values() if g.key in used)],
                                 files=full.files, skipped_files=full.skipped_files)

    def flag_counts(self, full):
        return {"no_currency": sum(o.currency is None for o in full.orders)}

    def writable(self, r):
        return len(r.accepted())

    def waiting(self, r):
        return len(r.orders) - len(r.accepted())

    def counts(self, r):
        s = orders.summarize(r)
        return rows(("Orders", s["orders_accepted"]), ("Accounts with orders", s["companies_accepted"]),
                    ("Order lines read", s["lines"]), ("Waiting for a company answer", s["orders_queued"]),
                    ("With no currency named", s["currency_not_stated"]))

    def notes(self, r):
        kept = [f"{f.name}: {f.reason}" for f in r.excluded()]
        return notes_from(r.review) + kept

    def file_stats(self, r):
        return {f.path: (f.documents, "orders") for f in r.files if f.kind == "orders"}

    def file_notes(self, r):
        out = {}
        for f in r.files:
            if f.kind != "orders" or not f.documents or f.kept_documents == f.documents:
                continue
            out[f.path] = (("superseded", f"All {f.documents} orders in it are also in a newer export of this upload, which is used instead.")
                           if not f.kept_documents else
                           ("ready", f"{f.kept_documents} of its {f.documents} orders go in; the other {f.documents - f.kept_documents} "
                                     "are also in a newer export, which is used instead."))
        return out

    def dry_run(self, client, r):
        p = orders_margince.plan(client, r)
        o, v = p["orders"], p["overviews"]
        return rows(("New accounts", p["companies_new"]), ("Accounts already in the CRM", p["companies_present"]),
                    ("New order notes", o["create"]), ("Order notes that change", o["update"]),
                    ("Order notes already up to date", o["same"] + o["newer"]),
                    ("New account overview notes", v["create"]), ("Overview notes that change", v["update"]),
                    ("Orders waiting for a company answer", p["orders_waiting_for_a_company_answer"]))

    def commit(self, client, r, emit):
        c = orders_margince.commit(client, r, on_progress=emit)
        d = c["companies"] or {}
        written = d.get("created", 0) + c["orders_created"] + c["orders_updated"] + c["overviews_created"] + c["overviews_updated"]
        present = c["orders_unchanged"] + c["orders_kept_newer"] + c["overviews_unchanged"]
        return report(written, present, len(c["errors"]), lines=[
            ("Order notes created", c["orders_created"]), ("Order notes updated", c["orders_updated"]),
            ("Order notes already up to date", c["orders_unchanged"] + c["orders_kept_newer"]),
            ("Account overviews created", c["overviews_created"]), ("Account overviews updated", c["overviews_updated"])],
            errors=c["errors"])


class Receivables(Adapter):
    key = "receivables"
    unit = "customers"

    def read(self, paths):
        return receivables.build_book(paths)

    def writable(self, r):
        return len(r.accepted())

    def waiting(self, r):
        return len(r.customers) - len(r.accepted())

    def counts(self, r):
        s = receivables.summarize(r)
        ccy = s["currency"] or ""
        return rows(("Customers with open invoices", s["customers_accepted"]), ("Open invoices", s["invoices"]),
                    (f"Open total {ccy}".strip(), money(s["open_total"])), (f"Overdue total {ccy}".strip(), money(s["overdue_total"])),
                    ("As of", s["as_of"]), ("Waiting for a company answer", s["customers_queued"]))

    def notes(self, r):
        out = notes_from(r.review)
        out += [f"{f.name}: replaced by the newer {f.superseded_by}" for f in r.files if f.superseded_by]
        out += [f"{f.name}: {f.reason}" for f in r.excluded()]
        return out

    def file_stats(self, r):
        return {f.path: (f.invoices, "invoices") for f in r.files if f.kind == "receivables"}

    def file_notes(self, r):
        return {f.path: ("superseded", f"An older snapshot: {f.superseded_by} is newer and is used instead.")
                for f in r.files if f.kind == "receivables" and f.superseded_by}

    def dry_run(self, client, r):
        p = receivables_margince.plan(client, r)
        a = p["accounts"]
        return rows(("New accounts", p["companies_new"]), ("Accounts already in the CRM", p["companies_present"]),
                    ("Accounts that get new AR figures", a["write"]), ("Accounts already up to date", a["same"] + a["newer"]),
                    ("Accounts whose AR goes back to zero (paid)", a["zero"]),
                    ("Customers waiting for a company answer", p["customers_waiting_for_a_company_answer"]),
                    ("How the figures are stored", "custom fields on the account" if p["mode"] == "fields" else p["mode"]))

    def commit(self, client, r, emit):
        c = receivables_margince.commit(client, r, on_progress=emit)
        d = c["companies"] or {}
        if c["mode"] == "fields":
            written, present = c["accounts_written"] + c["accounts_zeroed"], c["accounts_unchanged"] + c["accounts_kept_newer"]
            lines = [("Accounts given AR figures", c["accounts_written"]), ("Accounts already up to date", c["accounts_unchanged"]),
                     ("Accounts set back to zero", c["accounts_zeroed"])]
        else:
            written = c["notes_created"] + c["notes_updated"]
            present = c["notes_unchanged"] + c["notes_kept_newer"]
            lines = [("Account notes created", c["notes_created"]), ("Account notes updated", c["notes_updated"]),
                     ("Account notes up to date", present)]
        return report(written + d.get("created", 0), present, len(c["errors"]), lines=lines, errors=c["errors"])


class Quotes(Adapter):
    key = "quotes"
    unit = "quotes"

    def read(self, paths):
        return quotes.read_paths(paths)

    def writable(self, r):
        return len(r.accepted())

    def waiting(self, r):
        return len(r.quotes) - len(r.accepted())

    def counts(self, r):
        s = quotes.summarize(r)
        return rows(("Quotes (revisions)", s["quotes_accepted"]), ("Deals (enquiries)", s["deals"]),
                    ("Won order, kept open", s["won"]), ("Open", s["open"]), ("Lost or cancelled", s["lost"]),
                    ("Status not understood, kept open", s["unknown"]), ("Without a price", s["without_price"]),
                    ("Waiting for a company answer", s["quotes_queued"]))

    def notes(self, r):
        return notes_from(r.review) + [
            f"The log names no currency; offers and deals are written in {quotes_margince.DEFAULT_CURRENCY} "
            "(one quote linked to an ERP order is exactly 1/8 of its CNY value). No rate is applied."]

    def file_stats(self, r):
        return {f["path"]: (f["rows"], "quotes") for f in r.files}

    def dry_run(self, client, r):
        p = quotes_margince.plan(client, r)
        lines = [("Currency written", p["currency"]), ("New accounts", p["companies_new"]),
                 ("Accounts already in the CRM", p["companies_present"]), ("New deals", p["deals_new"]),
                 ("Deals already in the CRM", p["deals_present"]), ("New offers", p["offers_new"]),
                 ("New follow-up notes", p["notes_new"]),
                 ("Quotes waiting for a company answer", p["quotes_waiting_for_a_company_answer"])]
        if p["fields_missing"]:
            lines.append(("Deal custom fields that will be created", ", ".join(p["fields_missing"])))
        return rows(*lines)

    def commit(self, client, r, emit):
        c = quotes_margince.commit(client, r, on_progress=emit)
        d = c["companies"] or {}
        written = d.get("created", 0) + c["deals_created"] + c["offers_created"] + c["notes_created"]
        notes = [f"{c['offers_left_draft']} offers stay drafts: the CRM refused to send them (no exchange rate?)"] \
            if c["offers_left_draft"] else []
        return report(written, c["deals_present"], len(c["failed"]), lines=[
            ("Deals created", c["deals_created"]), ("Deals already in the CRM", c["deals_present"]),
            ("Deals closed as lost", c["deals_closed_lost"]), ("Offers created", c["offers_created"]),
            ("Offers sent", c["offers_sent"]), ("Offers rejected", c["offers_rejected"]),
            ("Follow-up notes created", c["notes_created"])], notes=notes,
            errors=[f"{f['id']}: {f['error']}" for f in c["failed"]])


ADAPTERS = {a.key: a for a in (Accounts(), Contacts(), Todos(), Orders(), Receivables(), Quotes())}
