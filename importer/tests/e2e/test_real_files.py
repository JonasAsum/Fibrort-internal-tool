"""Real files stay where they are (never in git); each test skips when the folder is missing.
The page must show exactly what the command-line path computes for the same single file."""
import re

import pytest

from excel.run import read_paths
from excel_margince import accepted
from visits import pipeline
from visits.match import match

from .helpers import real_files, real_id, stats, upload, page_text

VISIT_FILES = real_files("01", "02", "03")
OTHER_FILES = real_files("04", "05", "06", "07", "08", "09")


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
    outcome = upload(page, server, [path])
    if expected["found"] == 0:
        assert outcome == "error"
        assert re.search(r"no visit list", page_text(page), re.I)
        return
    assert outcome == "review", page.locator("#upload-error").inner_text()
    assert stats(page) == expected
    assert page.locator("#companies-body tr").count() == expected["companies"]
    assert page.locator("#contacts-body tr").count() == expected["contacts"]


@pytest.mark.skipif(not OTHER_FILES, reason="real Excel folder not available")
@pytest.mark.parametrize("path", OTHER_FILES, ids=real_id)
def test_files_that_are_not_visit_lists_say_so(page, server, path):
    if command_line(path)["found"]:
        pytest.skip("this file does hold visits; covered by the numbers test")
    assert upload(page, server, [path]) == "error"
    text = page_text(page)
    assert re.search(r"no visit list (was )?recogni[sz]ed", text, re.I), text
    assert page.locator("#upload-error").is_visible()
