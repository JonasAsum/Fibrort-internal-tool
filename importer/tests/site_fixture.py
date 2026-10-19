"""The site under test: the real Flask app, its upload folder and company answers in a temp folder, and an
in-memory CRM (tests/fake_site_crm.py) in place of the login to the real one. Used by the endpoint tests and by
the every-file test, so both drive exactly the same code the page does."""
import io
import os
import time

import pytest

from excel import run as excel_run
from tests.fake_site_crm import SiteCRM
from webui import server


class Site:
    def __init__(self, client, crm):
        self.client, self.crm = client, crm

    def upload(self, *files, rep=None, expect=200):
        """files: (path on disk, the name the browser would report, e.g. 'folder/sub/file.xlsx')."""
        data = {"files": [(io.BytesIO(open(path, "rb").read()), name) for path, name in files]}
        if rep:
            data["rep"] = rep
        r = self.client.post("/api/upload", data=data, content_type="multipart/form-data")
        assert r.status_code == expect, r.get_json()
        return r.get_json()

    def batch(self):
        return self.client.get("/api/batch").get_json()

    def dry_run(self):
        r = self.client.post("/api/dry_run")
        assert r.status_code == 200, r.get_json()
        return r.get_json()

    def commit(self, timeout=600):
        """Starts the import, waits for the job and returns it (status 'done' unless the job itself broke)."""
        r = self.client.post("/api/commit")
        assert r.status_code == 200, r.get_json()
        job_id = r.get_json()["job_id"]
        deadline = time.time() + timeout
        while time.time() < deadline:
            job = self.client.get(f"/api/commit/{job_id}").get_json()
            if job["status"] != "running":
                return job
            time.sleep(0.05)
        raise TimeoutError("the import did not finish")

    def families(self, job):
        return {f["key"]: f for f in job["result"]["families"]}


def reset_master():
    excel_run._master.cache_clear()


@pytest.fixture
def site(tmp_path, monkeypatch):
    decisions = tmp_path / "company_decisions.json"
    decisions.write_text("{}", encoding="utf-8")  # no earlier answers: the same questions every run
    monkeypatch.setenv("COMPANY_DECISIONS_FILE", str(decisions))
    monkeypatch.setattr(server, "UPLOAD_DIR", str(tmp_path / "uploads"))
    crm = SiteCRM()
    monkeypatch.setattr(server, "_client", lambda: crm)
    reset_master()
    server._forget()
    yield Site(server.app.test_client(), crm)
    server._forget()
    reset_master()
