"""The server records the checkout's git HEAD and newest .py mtime when it starts; the page warns
when the files have moved on, so nobody reads the output of older code."""
import os
import shutil
import subprocess
import time

import pytest

from .conftest import IMPORTER, start_server

SKIP = shutil.ignore_patterns(".venv", "__pycache__", "staging*", ".uploads", "tests", "docs", "*.xlsx", "out", "gold")


def git(cwd, *args):
    subprocess.run(["git", *args], cwd=cwd, check=True, capture_output=True,
                   env={**os.environ, "GIT_AUTHOR_NAME": "t", "GIT_AUTHOR_EMAIL": "t@t", "GIT_COMMITTER_NAME": "t",
                        "GIT_COMMITTER_EMAIL": "t@t"})


@pytest.fixture
def copy_server(tmp_path):
    root = tmp_path / "importer"
    shutil.copytree(IMPORTER, root, ignore=SKIP)
    git(root, "init", "-q")
    git(root, "add", "-A")
    git(root, "commit", "-q", "-m", "first")
    state = tmp_path / "state"
    state.mkdir()
    proc, base = start_server(str(root), str(state))
    yield root, base
    proc.terminate()
    proc.wait(timeout=10)


def banner(page, base):
    page.goto(base)
    page.wait_for_function("document.querySelector('#stale-banner')")  # the element exists from the start
    page.wait_for_timeout(500)
    return page.locator("#stale-banner")


def test_fresh_server_shows_no_warning(page, copy_server):
    root, base = copy_server
    assert not banner(page, base).is_visible()
    v = page.request.get(base + "/api/version").json()
    assert v["stale"] is False and v["started"] == v["current"]
    assert len(v["started"]["head"]) == 40 and v["started"]["newest_py_mtime"] > 0


def test_edited_python_file_triggers_the_warning(page, copy_server):
    root, base = copy_server
    future = time.time() + 30
    os.utime(root / "excel" / "run.py", (future, future))
    el = banner(page, base)
    el.wait_for(state="visible", timeout=5000)
    assert "older code" in el.inner_text() and "restart" in el.inner_text().lower()
    assert page.request.get(base + "/api/version").json()["stale"] is True


def test_new_commit_triggers_the_warning(page, copy_server):
    root, base = copy_server
    (root / "marker.txt").write_text("x")
    git(root, "add", "marker.txt")
    git(root, "commit", "-q", "-m", "second")
    el = banner(page, base)
    el.wait_for(state="visible", timeout=5000)
    v = page.request.get(base + "/api/version").json()
    assert v["stale"] and v["started"]["head"] != v["current"]["head"]
