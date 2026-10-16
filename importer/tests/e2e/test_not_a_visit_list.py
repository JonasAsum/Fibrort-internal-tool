"""A file the site cannot use must produce a visible sentence about that file, never a row of zeros and never the
old "no visit list recognised" for a list that is simply another kind."""
import re

from tests.excel_book import write_book

from .helpers import entry_rows, fresh_page, page_text, stats, upload
from .test_reasons import NAME21, weekly


def unknown_list(path):
    cells = {"A1": "Birthday", "B1": "Cake", "C1": "Guests", "A2": "x", "B2": "chocolate", "C2": 12}
    return write_book(str(path), {"List": {"cells": cells}})


def test_a_workbook_that_is_no_known_list_is_named_and_nothing_is_offered_to_import(page, server, tmp_path):
    assert upload(page, server, [unknown_list(tmp_path / "party.xlsx")]) == "review"
    (row,) = entry_rows(page)
    assert row["file"] == "party.xlsx" and row["badge"] == "Not recognised"
    assert re.search(r"no sheet has the header columns", row["outcome"]), row
    assert page.locator("#btn-go-import").is_disabled()
    assert not re.search(r"no visit list recogni[sz]ed", page_text(page), re.I)


def test_corrupt_file(page, server, tmp_path):
    bad = tmp_path / "broken.xlsx"
    bad.write_bytes(b"this is not a zip file at all")
    assert upload(page, server, [bad]) == "review"
    (row,) = entry_rows(page)
    assert row["file"] == "broken.xlsx" and row["badge"] == "Unreadable"
    assert re.search(r"could not be opened", row["outcome"], re.I), row


def test_not_an_excel_file(page, server, tmp_path):
    txt = tmp_path / "notes.txt"
    txt.write_text("hello")
    fresh_page(page, server)
    page.set_input_files("#file-input", str(txt))
    page.wait_for_selector("#upload-error:not(.hidden)")
    assert re.search(r"\.xlsx", page.locator("#upload-error").inner_text())
    assert page.locator("#step-2.active").count() == 0


def test_good_file_next_to_a_wrong_one_names_the_wrong_one(page, server, tmp_path):
    good = weekly(tmp_path / NAME21)
    other = unknown_list(tmp_path / "party.xlsx")
    assert upload(page, server, [good, other]) == "review"
    assert stats(page)["found"] == 1
    rows = {r["file"]: r for r in entry_rows(page)}
    assert rows["party.xlsx"]["badge"] == "Not recognised" and rows[NAME21]["badge"] == "Ready"
