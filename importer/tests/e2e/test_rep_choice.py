"""A workbook that names no sales rep gets one from the page's "Sales rep" choice."""
from datetime import datetime

from tests.excel_book import WEEKLY_HEADER, write_book

from .helpers import page_text, stats

NAME = "Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"


def ownerless(path):
    cells = dict(WEEKLY_HEADER)
    cells.update({"D9": datetime(2026, 5, 18), "F9": "Beta Test Co., Ltd.", "P9": "Call", "H9": "exist"})
    return write_book(str(path), {"CW21": {"cells": cells}})


def upload_with_rep(page, base, path, rep):
    page.request.post(base + "/api/clear")
    page.goto(base)
    if rep:
        page.select_option("#rep-select", rep)
    with page.expect_response(lambda r: r.url.endswith("/api/upload"), timeout=60_000):
        page.set_input_files("#file-input", [str(path)])
    page.wait_for_selector("#step-2.active", timeout=60_000)


def test_without_a_rep_the_visit_waits_for_review(page, server, tmp_path):
    upload_with_rep(page, server, ownerless(tmp_path / NAME), "")
    assert stats(page)["unsure"] == 1, page_text(page)


def test_the_chosen_rep_owns_the_visit(page, server, tmp_path):
    upload_with_rep(page, server, ownerless(tmp_path / NAME), "Mark Wang")
    assert stats(page)["visits"] == 1 and stats(page)["unsure"] == 0, page_text(page)
