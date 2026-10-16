"""A company the list does not know well is asked as "Is that the same company?": the three closest companies by
fuzzy score and "New company", four one-click answers. A click answers the question for good."""
import re
from datetime import datetime

from excel import run as excel_run
from tests.excel_book import WEEKLY_HEADER, write_book

from .helpers import page_text, stats
from .test_rep_choice import upload_with_rep

NAME = "Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"


def a_name_that_is_asked():
    """A made-up spelling close to a real customer, one the matcher cannot decide alone (it must show three candidates)."""
    master = excel_run._master(excel_run.IMPORTER_DIR)
    for c in master.companies:
        if not c.en:
            continue
        core = re.sub(r"[\s,]*(Co\.?,?\s*Ltd\.?|Ltd\.?|Inc\.?|GmbH|Corporation)\s*$", "", c.en.strip())
        if len(core) < 14:
            continue
        name = core[:-2] + "qx Co., Ltd."
        m = master.match(name)
        if m.method == "review" and len(m.candidates) == 3:
            return name
    raise AssertionError("no spelling of a real company is asked about with three candidates")


def workbook(path, name):
    cells = dict(WEEKLY_HEADER)
    cells.update({"D9": datetime(2026, 5, 18), "F9": name, "P9": "Call", "H9": "exist"})
    return write_book(str(path), {"CW21": {"cells": cells}})


def test_the_question_offers_three_companies_and_new_company(page, server, tmp_path):
    name = a_name_that_is_asked()
    upload_with_rep(page, server, workbook(tmp_path / NAME, name), "Mark Wang")
    card = page.locator(".cq-card", has=page.locator(".cq-name", has_text=name))
    assert card.count() == 1, page_text(page)
    assert card.locator(".cq-question").inner_text() == "Is that the same company?"
    assert card.locator(".cq-cand").count() == 3 and card.locator(".cq-new").count() == 1
    scores = [int(t.split("%")[0].split()[-1]) for t in card.locator(".cq-cand-note").all_inner_texts()]
    assert scores == sorted(scores, reverse=True), scores  # best first
    assert stats(page)["unsure"] == 1  # the visit waits for the answer


def test_one_click_on_new_company_answers_it(page, server, tmp_path):
    name = a_name_that_is_asked()
    upload_with_rep(page, server, workbook(tmp_path / NAME, name), "Mark Wang")
    card = page.locator(".cq-card", has=page.locator(".cq-name", has_text=name))
    card.locator(".cq-new").click()
    page.wait_for_function("name => ![...document.querySelectorAll('.cq-name')].some(e => e.textContent.includes(name))", arg=name, timeout=30_000)
    assert stats(page)["unsure"] == 0, page_text(page)
