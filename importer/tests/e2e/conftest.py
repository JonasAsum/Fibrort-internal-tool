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


def start_server(root, state_dir, extra_env=None, module="webui.server", args=(), self_config=False):
    """(process, base_url) for `python -m <module> --port N` run inside root. self_config points the server's
    saved CRM address at the server itself (a test CRM that is "reachable" without any CRM)."""
    port = free_port()
    base = f"http://127.0.0.1:{port}"
    config = os.path.join(state_dir, "config.json")
    if self_config:
        with open(config, "w") as fh:
            fh.write(f'{{"base_url": "{base}", "email": "test@example.test", "password": "not-a-real-password"}}')
    env = {**os.environ, "IMPORTER_UPLOAD_DIR": os.path.join(state_dir, "uploads"),
           "IMPORTER_CONFIG": config, "COMPANY_DECISIONS_FILE": os.path.join(state_dir, "company_decisions.json"),
           **(extra_env or {})}
    if not os.path.exists(env["COMPANY_DECISIONS_FILE"]):
        with open(env["COMPANY_DECISIONS_FILE"], "w") as fh:
            fh.write("{}")
    log = open(os.path.join(state_dir, "server.log"), "w")
    proc = subprocess.Popen([sys.executable, "-m", module, "--port", str(port), *args], cwd=root, env=env,
                            stdout=log, stderr=subprocess.STDOUT)
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


@pytest.fixture(scope="session")
def site_server(tmp_path_factory):
    """The site with an in-memory CRM and one invented customer in its company master (tests/e2e/site_server.py)."""
    state = str(tmp_path_factory.mktemp("site-state"))
    proc, base = start_server(IMPORTER, state, module="tests.e2e.site_server", self_config=True)
    yield base
    proc.terminate()
    proc.wait(timeout=10)


@pytest.fixture
def fresh_site_server(tmp_path):
    """A site no earlier test has imported into: its CRM has no custom fields yet."""
    proc, base = start_server(IMPORTER, str(tmp_path), module="tests.e2e.site_server", self_config=True)
    yield base
    proc.terminate()
    proc.wait(timeout=10)


@pytest.fixture(scope="module")
def real_site_server(tmp_path_factory):
    """The same, with the real customer lists as the company master (for the real sample files)."""
    state = str(tmp_path_factory.mktemp("real-site-state"))
    proc, base = start_server(IMPORTER, state, module="tests.e2e.site_server", args=("--real-master",), self_config=True)
    yield base
    proc.terminate()
    proc.wait(timeout=10)


def pytest_collection_modifyitems(items):
    for item in items:
        if "/tests/e2e/" in str(item.fspath):
            item.add_marker(pytest.mark.e2e)
