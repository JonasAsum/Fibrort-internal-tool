"""Thin REST client for the Margince CRM endpoints this pipeline needs:
login, the three-step CSV import (upload -> run -> approve), and
POST /activities. See margince/docs/how-to/import-a-company-spreadsheet.md
and backend/api/crm.yaml (/activities, /relationships) for the contracts
this wraps.
"""
import time

import requests


class ApiError(RuntimeError):
    def __init__(self, response: requests.Response):
        self.response = response
        try:
            detail = response.json()
        except ValueError:
            detail = response.text
        super().__init__(f"{response.status_code} {response.request.method} {response.request.url}: {detail}")


class MargenceClient:
    def __init__(self, base_url: str):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()

    @classmethod
    def as_agent(cls, base_url: str, token: str) -> "MargenceClient":
        """A client acting under an Agent Seat Passport (Authorization: Bearer mgp_…). Its writes are
        stamped agent:<passport>, and a gated change answers 403 approval_required (see passport.py)."""
        client = cls(base_url)
        client.session.headers["Authorization"] = f"Bearer {token}"
        return client

    def login(self, email: str, password: str) -> None:
        r = self.session.post(f"{self.base_url}/v1/auth/login", json={"email": email, "password": password})
        if r.status_code >= 400:
            raise ApiError(r)
        # crm_session is marked Secure, so requests' own cookie jar refuses
        # to resend it over plain http (same reason the getting-started
        # tutorial's curl example pulls it out by hand rather than using a
        # jar). Setting it as a literal header sidesteps that attribute.
        cookie = r.cookies.get("crm_session")
        if not cookie:
            raise RuntimeError("login succeeded but set no crm_session cookie")
        self.session.headers["Cookie"] = f"crm_session={cookie}"

    def _request(self, method: str, path: str, **kwargs) -> requests.Response:
        r = self.session.request(method, f"{self.base_url}{path}", **kwargs)
        if r.status_code >= 400:
            raise ApiError(r)
        return r

    def get(self, path: str, **kwargs) -> dict:
        return self._request("GET", path, **kwargs).json()

    def post(self, path: str, json: dict | None = None, **kwargs) -> dict:
        r = self._request("POST", path, json=json, **kwargs)
        return r.json() if r.content else {}

    def patch(self, path: str, json: dict, **kwargs) -> dict:
        r = self._request("PATCH", path, json=json, **kwargs)
        return r.json() if r.content else {}

    def post_status(self, path: str, json: dict | None = None, **kwargs) -> tuple[int, dict]:
        r = self._request("POST", path, json=json, **kwargs)
        return r.status_code, (r.json() if r.content else {})

    def paginate(self, path: str, params: dict | None = None):
        params = dict(params or {})
        while True:
            page = self.get(path, params=params)
            yield from page["data"]
            cursor = page.get("page", {}).get("next_cursor")
            if not cursor:
                return
            params["cursor"] = cursor

    # ---- CSV import (company / contact) ----

    def upload_csv(self, object_: str, csv_text: str, filename: str) -> str:
        files = {"file": (filename, csv_text.encode("utf-8"), "text/csv")}
        profile = self._request("POST", "/v1/imports/sources", files=files, data={"object": object_}).json()
        return profile["source_ref"]

    def create_import_run(self, object_: str, source_ref: str, mapping: dict, source_key: str,
                           on_duplicate: str = "skip") -> str:
        body = {"connector": "csv", "object": object_, "source_ref": source_ref, "mapping": mapping,
                "source_key": source_key, "on_duplicate": on_duplicate}
        run = self.post("/v1/imports", json=body)
        return run["id"]

    def wait_for_report(self, run_id: str, timeout_s: float = 60.0) -> dict:
        deadline = time.monotonic() + timeout_s
        while True:
            try:
                return self.get(f"/v1/imports/{run_id}/report")
            except ApiError as e:
                if e.response.status_code != 409 or time.monotonic() > deadline:
                    raise
                time.sleep(0.5)

    def approve_import_run(self, run_id: str) -> dict:
        return self.post(f"/v1/imports/{run_id}/approve")

    def wait_for_run_status(self, run_id: str, statuses: set[str], timeout_s: float = 120.0) -> dict:
        deadline = time.monotonic() + timeout_s
        while True:
            run = self.get(f"/v1/imports/{run_id}")
            if run["status"] in statuses:
                return run
            if time.monotonic() > deadline:
                raise TimeoutError(f"import run {run_id} still {run['status']!r} after {timeout_s}s")
            time.sleep(0.5)

    def run_csv_import(self, object_: str, csv_text: str, filename: str, mapping: dict, source_key: str,
                        on_duplicate: str = "skip") -> dict:
        """Upload, preview, and commit one CSV. Returns the final report."""
        source_ref = self.upload_csv(object_, csv_text, filename)
        run_id = self.create_import_run(object_, source_ref, mapping, source_key, on_duplicate)
        report = self.wait_for_report(run_id)
        self.approve_import_run(run_id)
        self.wait_for_run_status(run_id, {"complete", "failed"})
        return self.get(f"/v1/imports/{run_id}/report")

    # ---- companies / contacts lookup (post-import id resolution) ----

    def company_id_by_name(self) -> dict[str, str]:
        return {c["display_name"]: c["id"] for c in self.paginate("/v1/companies", {"limit": 200})}

    def contact_ids_for_company(self, company_id: str) -> dict[str, str]:
        return {c["full_name"]: c["contact_id"]
                for c in self.paginate(f"/v1/companies/{company_id}/contacts", {"limit": 200})}

    def contacts_for_company(self, company_id: str) -> list[dict]:
        return list(self.paginate(f"/v1/companies/{company_id}/contacts", {"limit": 200}))

    def get_contact(self, contact_id: str) -> dict:
        return self.get(f"/v1/contacts/{contact_id}")

    def create_contact(self, body: dict) -> dict:
        return self.post("/v1/contacts", json=body)

    def link_employment(self, contact_id: str, company_id: str) -> dict:
        # No started_at: a visit proves they worked there by then, never when they started.
        return self.post("/v1/relationships", json={"kind": "employment", "contact_id": contact_id,
                                                     "company_id": company_id, "employment_status": "current",
                                                     "source": "import"})

    def get_company(self, company_id: str) -> dict:
        return self.get(f"/v1/companies/{company_id}")

    # ---- workspace vocabulary ----

    def seats_by_name(self) -> dict[str, str]:
        return {u["display_name"]: u["id"] for u in self.paginate("/v1/users", {"limit": 200})
                if u.get("display_name")}

    def custom_fields(self, object_: str) -> list[dict]:
        return list(self.paginate("/v1/custom-fields", {"object": object_, "status": "active", "limit": 200}))

    def create_custom_field(self, body: dict) -> dict:
        return self.post("/v1/custom-fields", json={**body, "source": "import"})

    # ---- activities ----

    def log_activity(self, body: dict) -> tuple[bool, dict]:
        """Returns (created, activity). created is False for the idempotent
        200 the API answers when (source_system, source_id) already exists."""
        status, activity = self.post_status("/v1/activities", json=body)
        return status == 201, activity
