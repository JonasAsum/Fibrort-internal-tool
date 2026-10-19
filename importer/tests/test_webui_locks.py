import threading
import time
from types import SimpleNamespace

from webui import server


def test_a_second_commit_while_one_runs_is_refused(monkeypatch):
    started, release = threading.Event(), threading.Event()

    class Report:
        activities_skipped_no_date = []

    monkeypatch.setattr(server, "_client", lambda: object())
    monkeypatch.setattr(server.pipeline, "run", lambda *a, **k: (started.set(), release.wait(5), Report())[2])
    monkeypatch.setattr(server, "asdict", lambda r: {})
    monkeypatch.setitem(server.CURRENT, "result", object())
    monkeypatch.setitem(server.CURRENT, "parsed", SimpleNamespace(visits=[object()]))
    client = server.app.test_client()
    first = client.post("/api/commit")
    assert first.status_code == 200
    assert started.wait(5)
    second = client.post("/api/commit")
    assert second.status_code == 409
    assert "already running" in second.get_json()["error"]
    release.set()
    job = first.get_json()["job_id"]
    for _ in range(50):
        if client.get(f"/api/commit/{job}").get_json()["status"] != "running":
            break
        time.sleep(0.1)
    assert client.post("/api/commit").status_code == 200  # the lock was released
    release.set()
