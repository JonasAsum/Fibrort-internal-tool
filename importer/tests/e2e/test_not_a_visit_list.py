"""A file that is not a readable visit list must produce a visible sentence, never a row of zeros."""
import re

from tests.excel_book import write_book

from .helpers import fresh_page, page_text, stats, upload
from .test_reasons import NAME21, weekly


def customer_list(path):
    cells = {"A1": "Customer Code", "B1": "Customer Name", "C1": "Region", "A2": "X-1", "B2": "Made Up Co.", "C2": "North"}
    return write_book(str(path), {"List": {"cells": cells}})


def test_other_kind_of_workbook(page, server, tmp_path):
    assert upload(page, server, [customer_list(tmp_path / "customers.xlsx")]) == "error"
    text = page_text(page)
    assert re.search(r"no visit list recogni[sz]ed", text, re.I), text
    assert "customers.xlsx" in text


def test_corrupt_file(page, server, tmp_path):
    bad = tmp_path / "broken.xlsx"
    bad.write_bytes(b"this is not a zip file at all")
    assert upload(page, server, [bad]) == "error"
    text = page_text(page)
    assert re.search(r"broken\.xlsx.*could not be read", text, re.I | re.S), text


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
    other = customer_list(tmp_path / "customers.xlsx")
    assert upload(page, server, [good, other]) == "review"
    assert stats(page)["found"] == 1
    text = page_text(page)
    assert re.search(r"customers\.xlsx.*no visit list recogni[sz]ed", text, re.I | re.S), text
