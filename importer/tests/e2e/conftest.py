"""Browser tests start their OWN server from this checkout on a free port, with its own upload
and config files, so they never talk to a stale server and never touch the user's on 5055."""
import os
import socket
import subprocess
import sys
import time
import urllib.request
import uuid

import pytest

IMPORTER = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def start_server(root, state_dir, extra_env=None):
    """(process, base_url) for `python -m webui.server --port N` run inside root."""
    port = free_port()
    env = {**os.environ, "IMPORTER_UPLOAD_DIR": os.path.join(state_dir, "uploads"),
           "IMPORTER_CONFIG": os.path.join(state_dir, "config.json"), **(extra_env or {})}
    log = open(os.path.join(state_dir, "server.log"), "w")
    proc = subprocess.Popen([sys.executable, "-m", "webui.server", "--port", str(port)], cwd=root, env=env,
                            stdout=log, stderr=subprocess.STDOUT)
    base = f"http://127.0.0.1:{port}"
    deadline = time.time() + 30
    while time.time() < deadline:
        if proc.poll() is not None:
            break
        try:
            urllib.request.urlopen(base + "/api/config", timeout=1).read()
            return proc, base
        except OSError:
            time.sleep(0.2)
    proc.kill()
    raise RuntimeError(f"server did not start: {open(os.path.join(state_dir, 'server.log')).read()[-2000:]}")


@pytest.fixture(scope="session")
def run_id():
    return uuid.uuid4().hex[:8]


@pytest.fixture(scope="session")
def server(tmp_path_factory, run_id):
    state = str(tmp_path_factory.mktemp("server-state"))
    proc, base = start_server(IMPORTER, state, {"IMPORTER_SOURCE_SYSTEM": f"e2e-importer-{run_id}"})
    yield base
    proc.terminate()
    proc.wait(timeout=10)


def pytest_collection_modifyitems(items):
    for item in items:
        if "/tests/e2e/" in str(item.fspath):
            item.add_marker(pytest.mark.e2e)
