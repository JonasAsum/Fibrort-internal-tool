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


def test_version_ignores_the_tests_folder():
    import os
    path = os.path.join(server.ROOT, "tests", "_probe_newer.py")
    before = server._code_state()["newest_py_mtime"]
    try:
        with open(path, "w") as f:
            f.write("")
        os.utime(path, (before + 1000, before + 1000))
        assert server._code_state()["newest_py_mtime"] == before
    finally:
        os.remove(path)


def test_git_head_is_asked_at_most_once_per_30_seconds(monkeypatch):
    calls = []
    real = server.subprocess.run

    def counting(*a, **k):
        calls.append(a)
        return real(*a, **k)

    now = [1000.0]
    monkeypatch.setattr(server.subprocess, "run", counting)
    monkeypatch.setattr(server.time, "monotonic", lambda: now[0])
    server._HEAD_CACHE.clear()
    server._code_state()
    server._code_state()
    assert len(calls) == 1
    now[0] += 31
    server._code_state()
    assert len(calls) == 2
