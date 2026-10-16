"""The page as a person uses it: one front door for every kind of list, the review of each, a dry run, the import,
the second import, and the Checks. The site runs with an in-memory CRM (tests/e2e/site_server.py); every
workbook is a tiny invented one."""
import re

import openpyxl
import pytest

from webui import selftest

from .helpers import crm_counts, entry_rows, fresh_page, page_text, shot, upload, wait_import_result, watch_console

KINDS = {  # the key the selftest builds -> the tab label, the folder, what the table must say about it
    "accounts": ("Customer lists (04, 05)", "04"), "contacts": ("Contact lists (06)", "06"),
    "visits": ("Visit reports (01-03)", "01"), "todos": ("To-dos and sales topics (07)", "07"),
    "orders": ("Sales orders (08)", "08"), "receivables": ("Receivables (09)", "09"),
}


@pytest.fixture
def books(tmp_path):
    return {k: p for k, (p, _) in selftest.make_workbooks(str(tmp_path)).items()}


def invoice_list(tmp_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    wb.create_sheet("Invoice").append(["Item", "Qty", "Unit price", "Amount"])
    wb.create_sheet("Packing List Crate 1").append(["Crate", "Weight"])
    path = str(tmp_path / "2026-03-09_-_317_Fibro_Order_List.xlsx")
    wb.save(path)
    return path


def test_header_has_the_fibro_logo_title_and_no_console_errors(page, site_server):
    errors = watch_console(page)
    fresh_page(page, site_server)
    page.wait_for_load_state("networkidle")
    assert "FIBRO" in page.title() and "CRM Importer" in page.title()
    logo = page.locator("header .brand-logo img")
    assert "Orange Solutions" in logo.get_attribute("alt")
    assert logo.evaluate("img => img.complete && img.naturalWidth > 0")
    assert page.locator("header h1").inner_text() == "CRM Importer"
    assert page.locator("#btn-checks").is_visible()
    assert errors == []


def test_every_kind_of_list_is_recognised_in_one_upload_and_gets_its_own_panel(page, site_server, books):
    errors = watch_console(page)
    assert upload(page, site_server, list(books.values())) == "review", page_text(page)
    rows = entry_rows(page)
    assert len(rows) == 6 and {r["status"] for r in rows} == {"ready"}
    assert all(r["target"] and r["outcome"] and r["rows"] != "—" for r in rows)
    labels = [t.split("\n")[0].strip() for t in page.locator("#family-tabs .ftab").all_inner_texts()]
    assert [re.sub(r"\d+ files?.*", "", l).strip() for l in labels] == [  # the order they are committed in
        KINDS[k][0] for k in ("accounts", "contacts", "visits", "todos", "orders", "receivables")]
    assert "no visit list" not in page_text(page).lower()
    shot(page, "entry-review-mixed")
    for tab in page.locator("#family-tabs .ftab").all():
        tab.click()
        key = tab.get_attribute("data-fam")
        panel = page.locator(f"#fam-{key}")
        assert panel.is_visible(), key
        assert panel.locator(".stat").count() >= 1, key
    assert errors == []


def test_a_list_that_stays_out_of_the_crm_is_recognised_not_an_error(page, site_server, tmp_path):
    assert upload(page, site_server, [invoice_list(tmp_path)]) == "review"
    (row,) = entry_rows(page)
    assert row["status"] == "kept_out" and row["badge"] == "Kept out on purpose"
    assert "Recognised" in row["outcome"] and "kept out of the CRM on purpose" in row["outcome"]
    assert "invoice and packing lists do not belong in the CRM" in row["outcome"]
    assert page.locator("#upload-error").is_hidden()
    assert page.locator("#btn-go-import").is_disabled() and page.locator("#btn-dry-run").is_disabled()
    assert "Nothing in this upload can be imported" in page.locator("#nothing-banner").inner_text()
    shot(page, "entry-kept-out")


def test_a_file_nobody_knows_says_so_and_why(page, site_server, tmp_path):
    path = str(tmp_path / "mystery.xlsx")
    wb = openpyxl.Workbook()
    wb.active["A1"] = "Birthday plans"
    wb.save(path)
    assert upload(page, site_server, [path]) == "review"
    (row,) = entry_rows(page)
    assert row["badge"] == "Not recognised" and "no sheet has the header columns" in row["outcome"]
    assert page.locator("#upload-error").is_hidden()


def test_a_whole_folder_is_taken_and_only_the_workbooks_count(page, site_server, books, tmp_path):
    folder = tmp_path / "CRM lists" / "07 To-do and Sales Topics"
    folder.mkdir(parents=True)
    (folder / "todos.xlsx").write_bytes(open(books["todos"], "rb").read())
    (folder / ".DS_Store").write_bytes(b"\0")
    (folder / "notes.txt").write_text("not a workbook")
    fresh_page(page, site_server)
    with page.expect_response(lambda r: r.url.endswith("/api/upload")):
        page.set_input_files("#folder-input", str(tmp_path / "CRM lists"))
    page.wait_for_selector("#step-2.active")
    (row,) = entry_rows(page)
    assert row["status"] == "ready" and "To-do" in row["what"]
    assert "left out" in page.locator(".toast").first.inner_text()


def test_non_excel_files_picked_one_by_one_are_refused_by_name(page, site_server, tmp_path):
    txt = tmp_path / "notes.txt"
    txt.write_text("x")
    fresh_page(page, site_server)
    page.set_input_files("#file-input", str(txt))
    page.wait_for_selector("#upload-error:not(.hidden)")
    assert "notes.txt" in page.locator("#upload-error").inner_text()
    assert page.locator("#step-2.active").count() == 0


def test_a_flagged_group_can_be_left_out_with_one_click(page, site_server, tmp_path):
    from tests.excel_book import write_book
    rows = {f"{chr(65 + i)}3": h for i, h in enumerate(selftest.TODO_HEAD)}
    rows.update({"D4": "Server", "E4": "migrate the server", "F4": "Chenley", "H4": "WIP"})
    path = write_book(str(tmp_path / "to_do_list__JET-IT.xlsx"), {"Projekt-OPL": {"cells": rows}})
    assert upload(page, site_server, [path]) == "review"
    flag = page.locator("#fam-todos .flag input")
    assert flag.is_checked() and "JET-IT" in page.locator("#fam-todos .flag").inner_text()
    shot(page, "review-flag")
    flag.click()
    page.wait_for_selector("#fam-todos .flag.off")
    assert page.locator("#btn-go-import").is_disabled()
    page.locator("#fam-todos .flag input").click()
    page.wait_for_selector("#fam-todos .flag:not(.off)")
    assert page.locator("#btn-go-import").is_enabled()


def test_a_dry_run_shows_what_would_be_written_and_writes_nothing(page, site_server, books):
    assert upload(page, site_server, list(books.values())) == "review"
    before = crm_counts(page, site_server)
    page.click("#btn-dry-run")
    page.wait_for_selector("#confirm-modal:not(.hidden)")
    assert "Test import" in page.locator("#confirm-title").inner_text()
    assert page.locator("#confirm-ok").is_hidden() and "Nothing was written" in page.locator("#confirm-hint").inner_text()
    cards = page.locator("#confirm-report .report-card")
    assert cards.count() == 6
    text = page.locator("#confirm-report").inner_text()
    assert "New accounts" in text and "New tasks" in text
    shot(page, "dry-run")
    page.click("#confirm-cancel")
    assert crm_counts(page, site_server)["writes"] == before["writes"]


def test_import_everything_then_import_again_and_nothing_is_new(page, site_server, books):
    errors = watch_console(page)
    assert upload(page, site_server, list(books.values())) == "review"
    page.click("#btn-go-import")
    page.wait_for_selector("#confirm-modal:not(.hidden)")
    assert page.locator("#confirm-ok").is_visible()
    page.click("#confirm-ok")
    wait_import_result(page)
    cards = page.locator("#result-report .fam-result")
    assert cards.count() == 6
    assert [c.get_attribute("data-status") for c in cards.all()] == ["done"] * 6
    assert "Import complete" in page.locator("#result-title").inner_text()
    shot(page, "result-first")
    after_first = crm_counts(page, site_server)
    assert after_first["companies"] and after_first["contacts"] and after_first["activities"] and after_first["fields"]

    assert upload(page, site_server, list(books.values())) == "review"
    page.click("#btn-go-import")
    page.click("#confirm-ok")
    wait_import_result(page)
    assert [c.get_attribute("data-status") for c in page.locator("#result-report .fam-result").all()] == ["nothing_new"] * 6
    assert page.locator("#result-title").inner_text() == "Nothing new"
    assert "Nothing new:" in page.locator("#result-report").inner_text() and "already in the CRM" in page.locator("#result-report").inner_text()
    shot(page, "result-second")
    assert crm_counts(page, site_server) == after_first
    assert errors == []


def test_checks_page_lists_pass_and_warn_rows_with_a_fix_for_each_problem(page, fresh_site_server):
    site_server = fresh_site_server
    errors = watch_console(page)
    fresh_page(page, site_server)
    page.click("#btn-checks")
    page.wait_for_selector("#checks-env li")
    page.wait_for_function("!document.querySelector('#btn-checks-run').disabled")
    env = {li.get_attribute("data-check"): li for li in page.locator("#checks-env li").all()}
    assert {"crm", "login", "modules", "folders", "master", "fields"} <= set(env)
    assert env["crm"].get_attribute("data-status") == "pass" and env["modules"].get_attribute("data-status") == "pass"
    assert env["fields"].get_attribute("data-status") == "warn" and "How to fix" in env["fields"].inner_text()
    assert page.locator("#checks-self li[data-status=pass]").count() == 6
    assert "pass" in page.locator("#checks-summary").inner_text().lower()
    assert page.locator("#view-excel").is_hidden() and page.locator("#btn-checks.active").count() == 1
    shot(page, "checks")
    page.click("#btn-checks-run")
    page.wait_for_function("!document.querySelector('#btn-checks-run').disabled")
    assert errors == []


def test_phone_width_has_no_sideways_scroll_on_the_review(page, site_server, books):
    page.set_viewport_size({"width": 390, "height": 844})
    assert upload(page, site_server, list(books.values())) == "review"
    shot(page, "review-390")
    assert page.evaluate("document.documentElement.scrollWidth <= window.innerWidth + 1")
    page.click("#btn-checks")
    page.wait_for_selector("#checks-env li")
    shot(page, "checks-390")
    assert page.evaluate("document.documentElement.scrollWidth <= window.innerWidth + 1")
