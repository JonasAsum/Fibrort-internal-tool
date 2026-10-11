from webui import server


def test_version_reports_stale_when_files_are_newer_than_the_start(monkeypatch):
    client = server.app.test_client()
    v = client.get("/api/version").get_json()
    assert v["stale"] is False
    monkeypatch.setitem(server.STARTED, "newest_py_mtime", v["current"]["newest_py_mtime"] - 5)
    assert client.get("/api/version").get_json()["stale"] is True


def test_version_reports_stale_when_head_moved(monkeypatch):
    monkeypatch.setitem(server.STARTED, "head", "0" * 40)
    assert server.app.test_client().get("/api/version").get_json()["stale"] is True
