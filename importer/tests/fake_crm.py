"""An in-memory stand-in for common.client.MargenceClient, just enough for visits.pipeline.run.
It keeps what the pipeline writes so a test can count it. Everything here is made up."""
import csv
import io
import itertools


class FakeCRM:
    def __init__(self):
        self.ids = itertools.count(1)
        self.companies, self.contacts, self.activities = {}, {}, {}
        self.employment = []
        self.fields = []
        self.contact_conflict = {}  # email (lower) -> contact id that already holds it

    def _id(self, prefix):
        return f"{prefix}{next(self.ids)}"

    # generic
    def get(self, path, **kw):
        if path == "/v1/me":
            return {"user": {"id": "me"}}
        raise KeyError(path)

    def patch(self, path, body, **kw):
        kind, ident = path.strip("/").split("/")[1:3]
        store = {"companies": self.companies, "contacts": self.contacts, "activities": self.activities}[kind]
        if kind == "contacts":
            for e in body.get("emails") or []:  # another contact holding the address is a 409
                holder = self.contact_conflict.get(e["email"].lower())
                if holder and holder != ident:
                    raise _Conflict(holder)
            for e in body.get("emails") or []:
                self.contact_conflict[e["email"].lower()] = ident
        store[ident].update(body)
        return store[ident]

    def paginate(self, path, params=None):
        assert path == "/v1/activities"
        yield from list(self.activities.values())

    # vocabulary
    def seats_by_name(self):
        return {}

    def custom_fields(self, object_):
        return list(self.fields)

    def create_custom_field(self, body):
        f = {**body, "column_name": "cf_" + body["label"].lower().replace(" ", "_")}
        self.fields.append(f)
        return f

    # companies / contacts
    def company_id_by_name(self):
        return {c["display_name"]: i for i, c in self.companies.items()}

    def run_csv_import(self, object_, csv_text, filename, **kw):
        for row in csv.DictReader(io.StringIO(csv_text)):
            self.companies[self._id("co")] = {"id": None, "display_name": row["name"], "lifecycle": None}
        return {"disposition": {}, "rows_read": 0, "issues": []}

    def get_company(self, company_id):
        return self.companies[company_id]

    def contacts_for_company(self, company_id):
        return [{"full_name": self.contacts[c]["full_name"], "contact_id": c}
                for co, c in self.employment if co == company_id]

    def get_contact(self, contact_id):
        return self.contacts[contact_id]

    def create_contact(self, body):
        for e in body.get("emails") or []:
            if e["email"].lower() in self.contact_conflict:
                from common.client import ApiError
                raise _Conflict(self.contact_conflict[e["email"].lower()])
        cid = self._id("ct")
        self.contacts[cid] = {**body, "id": cid}
        for e in body.get("emails") or []:
            self.contact_conflict[e["email"].lower()] = cid
        return self.contacts[cid]

    def link_employment(self, contact_id, company_id):
        self.employment.append((company_id, contact_id))
        return {}

    # activities
    def log_activity(self, body):
        for a in self.activities.values():
            if (a["source_system"], a["source_id"]) == (body["source_system"], body["source_id"]):
                return False, a
        aid = self._id("ac")
        self.activities[aid] = {**body, "id": aid, "is_done": False}
        return True, self.activities[aid]

    def count(self, kind):
        return sum(1 for a in self.activities.values() if a["kind"] == kind)


def _Conflict(existing_id):
    """An ApiError shaped like the CRM's 409 duplicate_email."""
    import requests
    from common.client import ApiError
    resp = requests.Response()
    resp.status_code = 409
    resp._content = ('{"code": "duplicate_email", "details": {"existing_id": "%s"}}' % existing_id).encode()
    resp.request = requests.Request("POST", "http://fake/v1/contacts").prepare()
    return ApiError(resp)
