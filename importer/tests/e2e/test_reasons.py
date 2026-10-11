"""Every way a visit or file is held back must say so, in plain words, on the page.
One synthetic workbook per reason (tests/excel_book.py); nothing here is real data."""
import re
import shutil
from datetime import datetime

import pytest

from excel.run import _master, IMPORTER_DIR
from tests.excel_book import WEEKLY_HEADER, write_book
from tests.test_webui_upload import LUCAS_NAME, lucas_book

from .helpers import page_text, stats, upload

NAME21 = "Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"
NAME22 = "Template for Weekly Sales Reporting 2026_CW22_Test Rep.xlsx"


def weekly(path, date=datetime(2026, 5, 18), customer="Beta Test Co., Ltd.", sheet="CW21", extra=None):
    cells = dict(WEEKLY_HEADER)
    cells.update({"E4": "Owner：", "F4": "Jason Zhao", "B9": "Jason Zhao", "P9": "Call",
                  "H9": "exist", "I9": "Name Test"})
    if date is not None:
        cells["D9"] = date
    if customer is not None:
        cells["F9"] = customer
    sheets = {sheet: {"cells": cells}}
    sheets.update(extra or {})
    return write_book(str(path), sheets)


def shown(page, base, path):
    assert upload(page, base, [path]) == "review", page_text(page)
    return page_text(page)


def test_week_mismatch(page, server, tmp_path):
    p = weekly(tmp_path / NAME21, date=datetime(2026, 3, 2))
    text = shown(page, server, p)
    assert re.search(r"not in the week", text)
    assert stats(page)["unsure"] == 1 and stats(page)["visits"] == 0


def test_missing_date(page, server, tmp_path):
    text = shown(page, server, weekly(tmp_path / NAME21, date=None))
    assert re.search(r"no readable date", text)
    assert stats(page)["unsure"] == 1


def test_missing_customer(page, server, tmp_path):
    text = shown(page, server, weekly(tmp_path / NAME21, customer=None))
    assert re.search(r"no customer", text)
    assert stats(page)["unsure"] == 1


def test_ambiguous_day_month(page, server, tmp_path):
    text = shown(page, server, weekly(tmp_path / NAME21, date="03/04/2026"))
    assert re.search(r"day/month or month/day", text)
    assert stats(page)["unsure"] == 1


def test_company_close_to_several_master_companies(page, server, tmp_path):
    master = _master(IMPORTER_DIR)
    name = next((n for c in master.companies for n in (c.en, c.cn) if n and master.match(n).method == "ambiguous"), None)
    if name is None:
        pytest.skip("the company master has no ambiguous name")
    text = shown(page, server, weekly(tmp_path / NAME21, customer=name))
    assert re.search(r"close to more than one company", text)
    assert stats(page)["unsure"] == 1


def test_same_visit_in_two_files_says_which_file_won(page, server, tmp_path):
    older = weekly(tmp_path / NAME21)
    newer = weekly(tmp_path / NAME22, sheet="CW22")
    assert upload(page, server, [older, newer]) == "review"
    text = page_text(page)
    assert stats(page)["found"] == 1
    # The CW22 copy holds a visit dated in week 21: the CW21 file is where it belongs, so that one is used.
    assert re.search(r"same visit.*" + re.escape(NAME21), text, re.S), text
    assert re.search(re.escape(NAME22) + r".*(ignored|left out|not used)", text, re.S), text


def test_sheet_that_cannot_be_read_is_named(page, server, tmp_path):
    # Date in column A looks like the old visit list, but the sheet name carries no week number.
    bad = {"Notes": {"cells": {"A1": "Date", "A2": datetime(2026, 5, 18), "C2": "Gamma Test Co., Ltd."}}}
    p = weekly(tmp_path / NAME21, extra=bad)
    text = shown(page, server, p)
    assert stats(page)["found"] == 1  # the good sheet still counts
    assert re.search(r"sheet .?Notes.? could not be read", text, re.I), text
    assert re.search(r"no CW/Week number", text, re.I), text


def test_huge_number_in_date_cell_does_not_hide_the_sheet(page, server, tmp_path):
    # Used to kill the whole sheet with "Python int too large to convert to C int".
    text = shown(page, server, weekly(tmp_path / NAME21, date=10 ** 30))
    assert stats(page)["found"] == 1
    assert "too large" not in text
    assert re.search(r"no readable date", text)


def test_lucas_copy_of_another_week(page, server, tmp_path):
    p = tmp_path / "Weekly Sales Planning and Reporting 2026_CW13-Lucas2.xlsx"
    lucas_book(str(p))
    text = shown(page, server, p)
    assert re.search(r"not in the week", text)


def test_filled_sheet_without_a_visit_layout_is_named(page, server, tmp_path):
    misc = {"Misc": {"cells": {f"{c}{r}": f"x{r}{c}" for c in "ABC" for r in range(1, 6)}}}
    text = shown(page, server, weekly(tmp_path / NAME21, extra=misc))
    assert stats(page)["found"] == 1
    assert re.search(r"sheet .?Misc.? was not used: no visit list recognised", text), text
