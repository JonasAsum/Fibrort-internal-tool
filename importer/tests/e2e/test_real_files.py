"""Real files stay where they are (never in git); each test skips when the folder is missing.
The page must show exactly what the command-line path computes for the same single file."""
import os
import re

import pytest

from excel.run import read_paths
from excel_margince import accepted
from visits import pipeline
from visits.match import match

from .helpers import entry_rows, page_text, real_files, real_id, shot, stats, upload

VISIT_FILES = real_files("01", "02", "03")
OTHER_FILES = real_files("04", "05", "06", "07", "08", "09")


@pytest.fixture(autouse=True, scope="module")
def no_earlier_answers(tmp_path_factory):
    """The test servers start with no saved company answers, so the command-line side must not use the owner's either."""
    from excel import run
    answers = tmp_path_factory.mktemp("answers") / "company_decisions.json"
    answers.write_text("{}", encoding="utf-8")
    patch = pytest.MonkeyPatch()
    patch.setenv("COMPANY_DECISIONS_FILE", str(answers))
    run._master.cache_clear()
    yield
    patch.undo()
    run._master.cache_clear()


def command_line(path):
    run = read_paths([path])
    parsed = accepted(run)
    matched = match(parsed.visits)
    counts = pipeline.plan_counts(parsed, matched)
    found = len(run.parsed.visits)
    return {"found": found, "visits": len(parsed.visits), "unsure": found - len(parsed.visits),
            "companies": len(matched.companies), "contacts": len(matched.contacts),
            "tasks": counts["tasks"], "weeks": counts["weeks"]}


@pytest.mark.skipif(not VISIT_FILES, reason="real Excel folder not available")
@pytest.mark.parametrize("path", VISIT_FILES, ids=real_id)
def test_page_shows_what_the_command_line_computes(page, server, path):
    expected = command_line(path)
    assert upload(page, server, [path]) == "review", page.locator("#upload-error").inner_text()
    (row,) = entry_rows(page)
    assert row["status"] in ("ready", "empty"), row
    assert "no visit list" not in page_text(page).lower()
    if expected["found"] == 0:  # a blank template: recognised, and it says it holds nothing
        assert row["status"] == "empty" and page.locator("#btn-go-import").is_disabled()
        return
    assert stats(page) == expected
    assert page.locator("#companies-body tr").count() == expected["companies"]
    assert page.locator("#contacts-body tr").count() == expected["contacts"]


@pytest.mark.skipif(not OTHER_FILES, reason="real Excel folder not available")
@pytest.mark.parametrize("path", OTHER_FILES, ids=real_id)
def test_every_other_list_is_placed_by_the_page_never_called_a_missing_visit_list(page, real_site_server, path):
    outcome = upload(page, real_site_server, [path])
    assert outcome == "review", page.locator("#upload-error").inner_text()
    text = page_text(page)
    assert not re.search(r"no visit list", text, re.I), text
    (row,) = entry_rows(page)
    assert row["status"] in ("ready", "kept_out", "empty"), row
    if row["status"] == "kept_out":
        assert "kept out of the CRM on purpose" in row["outcome"] and page.locator("#upload-error").is_hidden()
    else:
        assert row["target"] and row["target"] != "Not imported"
        assert page.locator("#family-tabs .ftab.active").count() == 1


# The seven kinds the owner tried by hand, each driven through the page with a screenshot (E2E_SHOTS=<folder>).
OWNER_TRIED = {
    "04-erp-north": ("04", "North Region.xls"), "05-customer-list": ("05", "Customer_List_20260104"),
    "06-monica-handover": ("06", "Handover_Contacts_List_Monica"), "06-finance-contacts": ("06", "客户供应商"),
    "07-todo": ("07", "to_do_list__JET-IT"), "08-orders": ("08", "Sales Order List_ List_0928"),
    "09-ar": ("09", "Financial_Status_Fibro_CN_20260410"),
}


@pytest.mark.skipif(not OTHER_FILES, reason="real Excel folder not available")
@pytest.mark.parametrize("key", list(OWNER_TRIED))
def test_the_lists_the_owner_tried_by_hand_open_on_a_review_with_numbers(page, real_site_server, key):
    folder, needle = OWNER_TRIED[key]
    path = next((p for p in real_files(folder) if needle in os.path.basename(p)), None)
    if path is None:
        pytest.skip(f"no {needle} in folder {folder}")
    assert upload(page, real_site_server, [path]) == "review"
    (row,) = entry_rows(page)
    assert row["status"] == "ready" and row["rows"] != "—", row
    assert page.locator("#family-tabs .ftab.active .badge.info").count() == 1
    assert page.locator(".fam-panel:not(.hidden) .stat").count() >= 1
    shot(page, f"real-{key}")
    page.click("#btn-dry-run")
    page.wait_for_selector("#confirm-modal:not(.hidden)", timeout=120_000)
    assert page.locator("#confirm-report .report-card dt").count() >= 1
    shot(page, f"real-{key}-dryrun")
