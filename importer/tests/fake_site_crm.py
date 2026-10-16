"""One in-memory stand-in for common.client.MargenceClient that every importer family can run against
(visits, accounts, contacts, to-dos, orders, receivables): the site tests use it to commit a real file
and then commit it again. It keeps every write in `writes`, so "the second commit wrote nothing" is
a count. Everything it holds is made up by the importers' own code from the files under test."""
import copy
import csv
import io
import itertools

from tests.fake_crm import _Conflict


class SiteCRM:
    def __init__(self, seats=None):
        self.ids = itertools.count(1)
        self.companies, self.contacts, self.leads, self.activities = {}, {}, {}, {}
        self.employment, self.writes = [], []
        self.fields = {"company": [], "contact": [], "lead": []}
        self.seats = dict(seats or {"Richard Yu": "seat-richard", "Ethan Lv": "seat-ethan"})
        self.email_owner = {}  # email (lower) -> contact id that holds it, like the CRM's duplicate_email 409
        self.deals = []

    def _id(self, prefix):
        return f"{prefix}{next(self.ids)}"

    def _write(self, kind, body):
        self.writes.append((kind, copy.deepcopy(body)))

    def snapshot(self):
        return copy.deepcopy((self.companies, self.contacts, self.leads, self.activities, self.employment, self.fields))

    # ---- reads ----
    def get(self, path, **kw):
        if path == "/v1/me":
            return {"user": {"id": "me", "display_name": "Test Admin"}, "workspace_name": "Test workspace"}
        kind, ident = path.strip("/").split("/")[1:3]
        return dict({"companies": self.companies, "contacts": self.contacts}[kind][ident])

    def paginate(self, path, params=None):
        if path == "/v1/deals":
            yield from [d for d in self.deals if d.get("status") == "open"]
            return
        store = {"/v1/companies": self.companies, "/v1/leads": self.leads, "/v1/activities": self.activities,
                 "/v1/contacts": self.contacts}[path]
        yield from (copy.deepcopy(v) for v in list(store.values()))

    def get_company(self, company_id):
        return dict(self.companies[company_id])

    def get_contact(self, contact_id):
        return dict(self.contacts[contact_id])

    def company_id_by_name(self):
        return {c["display_name"]: i for i, c in self.companies.items()}

    def contacts_for_company(self, company_id):
        return [{"full_name": self.contacts[c]["full_name"], "contact_id": c} for co, c in self.employment if co == company_id]

    def seats_by_name(self):
        return dict(self.seats)

    # ---- vocabulary ----
    def custom_fields(self, object_):
        return list(self.fields[object_])

    def create_custom_field(self, body):
        f = {**body, "column_name": "cf_" + body["label"].lower().replace(" ", "_")}
        self.fields[body["object"]].append(f)
        self._write("field", body)
        return dict(f)

    # ---- companies: the CSV run the visits pipeline uses, and the plain POST the account importer uses ----
    def upload_csv(self, object_, csv_text, filename):
        return csv_text

    def create_import_run(self, object_, source_ref, mapping, source_key, on_duplicate="skip"):
        return source_ref

    def wait_for_report(self, run_id, timeout_s=60.0):
        rows = list(csv.DictReader(io.StringIO(run_id)))
        return {"disposition": {"would_create": len(rows)}, "rows_read": len(rows), "issues": []}

    def run_csv_import(self, object_, csv_text, filename, mapping=None, source_key=None, on_duplicate="skip"):
        rows = list(csv.DictReader(io.StringIO(csv_text)))
        have = {c["display_name"] for c in self.companies.values()}
        created = 0
        for row in rows:
            if row["name"] in have:
                continue
            cid = self._id("co")
            self.companies[cid] = {"id": cid, "display_name": row["name"], "lifecycle": "unknown", "relationship_types": []}
            self._write("company", {"display_name": row["name"]})
            have.add(row["name"])
            created += 1
        return {"disposition": {"created": created, "skipped_duplicate": len(rows) - created}, "rows_read": len(rows), "issues": []}

    def post(self, path, json=None, **kw):
        assert path == "/v1/companies", path
        cid = self._id("co")
        self.companies[cid] = {"lifecycle": "unknown", "relationship_types": [], **json, "id": cid}
        self._write("company", json)
        return dict(self.companies[cid])

    def patch(self, path, body, **kw):
        kind, ident = path.strip("/").split("/")[1:3]
        store = {"companies": self.companies, "contacts": self.contacts, "leads": self.leads,
                 "activities": self.activities}[kind]
        if kind == "contacts":
            for e in body.get("emails") or []:  # like the CRM: another contact holding the address is a 409
                holder = self.email_owner.get(e["email"].lower())
                if holder and holder != ident:
                    raise _Conflict(holder)
            for e in body.get("emails") or []:
                self.email_owner[e["email"].lower()] = ident
        store[ident].update(body)
        self._write(kind + "-patch", body)
        return dict(store[ident])

    def post_status(self, path, json=None, **kw):
        assert path == "/v1/leads", path
        for lead in self.leads.values():
            if (lead["source_system"], lead["source_id"]) == (json["source_system"], json["source_id"]):
                return 200, dict(lead)
        lid = self._id("ld")
        self.leads[lid] = {**json, "id": lid}
        self._write("lead", json)
        return 201, dict(self.leads[lid])

    # ---- contacts ----
    def create_contact(self, body):
        for e in body.get("emails") or []:
            if e["email"].lower() in self.email_owner:
                raise _Conflict(self.email_owner[e["email"].lower()])
        cid = self._id("ct")
        self.contacts[cid] = {**body, "id": cid}
        for e in body.get("emails") or []:
            self.email_owner[e["email"].lower()] = cid
        self._write("contact", body)
        return dict(self.contacts[cid])

    def link_employment(self, contact_id, company_id):
        self.employment.append((company_id, contact_id))
        self._write("employment", {"contact": contact_id, "company": company_id})
        return {}

    # ---- activities ----
    def log_activity(self, body):
        for a in self.activities.values():
            if (a["source_system"], a["source_id"]) == (body["source_system"], body["source_id"]):
                return False, dict(a)
        aid = self._id("ac")
        self.activities[aid] = {**body, "id": aid, "is_done": False}
        self._write("activity", body)
        return True, dict(self.activities[aid])

    def count(self, kind):
        return sum(1 for a in self.activities.values() if a["kind"] == kind)
