"""Click Import for real against the LOCAL test CRM only, twice, and prove nothing is written twice.
Names carry a per-run id and the records carry a per-run source_system, so test data is identifiable."""
from collections import Counter
from datetime import datetime
from urllib.parse import urlparse

import pytest
import requests

from common.client import MargenceClient
from tests.excel_book import LUCAS_HEADER, write_book

from .helpers import page_text, upload

CRM = "http://localhost:8080"
EMAIL, PASSWORD = "admin@demo.test", "demo-password-123"
LUCAS_NAME = "Weekly Sales Planning and Reporting 2026_CW10-Lucas.xlsx"


def local_crm():
    assert urlparse(CRM).hostname in ("localhost", "127.0.0.1"), "e2e import tests only talk to a local CRM"
    try:
        requests.get(CRM + "/", timeout=2)
        client = MargenceClient(CRM)
        client.login(EMAIL, PASSWORD)
        return client
    except Exception as exc:  # not reachable or no such login: nothing to test against
        pytest.skip(f"local test CRM not available at {CRM}: {exc}")


@pytest.fixture(scope="module")
def crm():
    return local_crm()


def workbook(path, run_id):
    cells = dict(LUCAS_HEADER)
    cells.update({"C4": "Owner：", "D4": "Lucas Guo"})
    # two clearly different companies: names that differ by one letter are asked about ("same company?") and wait
    for row, tag, day, word in ((15, "a", 2, "Alpha Marine"), (19, "b", 3, "Zephyr Foundry")):
        cells.update({f"C{row}": datetime(2026, 3, day), f"D{row}": "Shenzhen",
                      f"E{row}": f"E2E{run_id} {word} Co., Ltd.", f"F{row}": "A", f"G{row}": "exist",
                      f"H{row}": f"Tester{run_id}{tag}", f"K{row}": f"e2e-{run_id}-{tag}@example.com",
                      f"L{row}": "Phone call", f"M{row}": f"Send e2e quote {run_id}{tag}",
                      f"O{row}": datetime(2026, 3, 20)})
    return write_book(str(path), {"W01": {"cells": cells}})


def import_through_page(page, base, path):
    """Clicks Import for real; returns (the visit card's status, its label/value pairs)."""
    assert upload(page, base, [path]) == "review", page_text(page)
    page.click("#btn-go-import")
    page.click("#confirm-ok")
    page.wait_for_selector("#import-done:not(.hidden), #import-error:not(.hidden)", timeout=180_000)
    assert page.locator("#import-done:not(.hidden)").count(), page.locator("#error-text").inner_text()
    card = page.locator('#result-report .fam-result[data-fam="visits"]')
    pairs = {}
    for dl in card.locator("dl").all():
        for dt, dd in zip(dl.locator("dt").all_inner_texts(), dl.locator("dd").all_inner_texts()):
            pairs.setdefault(dt, dd)
    return card.get_attribute("data-status"), pairs


def read_back(crm, run_id, source_system):
    companies = [c for c in crm.paginate("/v1/companies", {"limit": 200}) if f"E2E{run_id}" in c["display_name"]]
    contacts = [c for c in crm.paginate("/v1/contacts", {"limit": 200}) if f"Tester{run_id}" in c["full_name"]]
    activities = [a for a in crm.paginate("/v1/activities", {"limit": 200}) if a.get("source_system") == source_system]
    return companies, contacts, activities


def test_import_then_import_again_writes_nothing_twice(page, server, crm, run_id, tmp_path):
    source_system = f"e2e-importer-{run_id}"
    path = workbook(tmp_path / LUCAS_NAME, run_id)

    first_status, first = import_through_page(page, server, path)
    assert first_status == "done" and first["Visits logged"] == "2" and first["Visits already in the CRM"] == "0"
    assert first["Contacts created"] == "2" and first["Tasks created"] == "2" and first["Weekly reports created"] == "1"

    companies, contacts, activities = read_back(crm, run_id, source_system)
    assert len(companies) == 2 and len(contacts) == 2
    assert Counter(a["kind"] for a in activities) == {"meeting": 2, "task": 2, "note": 1}
    assert len({a["source_id"] for a in activities}) == len(activities)

    second_status, second = import_through_page(page, server, path)
    assert second_status == "nothing_new" and second["Visits logged"] == "0" and second["Visits already in the CRM"] == "2"
    assert second["Contacts created"] == "0" and second["Tasks created"] == "0" and second["Weekly reports created"] == "0"
    assert page.locator("#result-title").inner_text() == "Nothing new"

    again = read_back(crm, run_id, source_system)
    assert [len(x) for x in again] == [2, 2, 5]
    assert {c["id"] for c in again[0]} == {c["id"] for c in companies}
    assert {c["id"] for c in again[1]} == {c["id"] for c in contacts}
    assert {a["id"] for a in again[2]} == {a["id"] for a in activities}
