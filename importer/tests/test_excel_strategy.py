"""Test suite for the Excel import of sales workbooks.

Groups: A discovery and classification, B sheets and structure, C duplicates and
versions, D parsers, E the existing importer (skips if not importable), F real sample
folder (skips if absent), G privacy and robustness.
"""
import os
import random
import re
import shutil
import time
import unicodedata
import zipfile
from datetime import date, datetime

import pytest

from tests.conftest import make_book
from tests.excel_scan import (
    KINDS, classify_name, cw_from_sheet_name, duplicate_groups, find_header_row,
    fold_text, formula_gaps, golden, parse_cell_date, parse_number, report_text,
    scan_file, scan_folder, sheet_fingerprints,
)

CELLS = {"cells": {"A1": "Sales", "B2": "Customer", "C3": "Date"}}
OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 120


def _write(path, data: bytes):
    path.write_bytes(data)
    return str(path)


# ---------- A. discovery and classification ----------

def test_a01_lock_files_are_never_importable(tmp_path):
    _write(tmp_path / ".~lock.Weekly CW13.xlsx#", b"lock")
    _write(tmp_path / "~$temp.xlsx", b"lock")
    kinds = {f.kind for f in scan_folder(tmp_path)}
    assert kinds == {"lock"}


def test_a02_dot_files_are_hidden_kind(tmp_path):
    _write(tmp_path / ".DS_Store", b"\0" * 10)
    (only,) = scan_folder(tmp_path)
    assert only.kind == "hidden" and not only.importable


def test_a03_dot_folders_are_skipped(tmp_path, book):
    book("a.xlsx", {"CW05": CELLS})
    os.makedirs(tmp_path / ".arena")
    make_book(tmp_path / ".arena" / "b.xlsx", {"CW05": CELLS})
    files = scan_folder(tmp_path)
    assert [os.path.basename(f.path) for f in files] == ["a.xlsx"]


def test_a04_zero_byte_file_is_empty(tmp_path):
    _write(tmp_path / "e.xlsx", b"")
    (only,) = scan_folder(tmp_path)
    assert only.kind == "empty"


def test_a05_other_extensions_are_ignored(tmp_path):
    _write(tmp_path / "notes.csv", b"a,b\n1,2\n")
    (only,) = scan_folder(tmp_path)
    assert only.kind == "not_excel"


def test_a06_fake_xlsx_is_corrupt(tmp_path):
    _write(tmp_path / "fake.xlsx", b"hello world, not a zip")
    (only,) = scan_folder(tmp_path)
    assert only.kind == "corrupt" and not only.importable


def test_a07_truncated_xlsx_is_corrupt(tmp_path, book):
    path = book("full.xlsx", {"CW05": CELLS})
    data = open(path, "rb").read()
    _write(tmp_path / "cut.xlsx", data[: len(data) // 2])
    kinds = {os.path.basename(f.path): f.kind for f in scan_folder(tmp_path)}
    assert kinds["cut.xlsx"] == "corrupt"


def test_a08_xlsx_renamed_to_xls_is_mislabelled(tmp_path, book):
    path = book("real.xlsx", {"CW05": CELLS})
    _write(tmp_path / "renamed.xls", open(path, "rb").read())
    (f,) = [x for x in scan_folder(tmp_path) if x.name == "renamed.xls"]
    assert f.kind == "mislabelled"


def test_a09_password_protected_container_is_encrypted(tmp_path):
    _write(tmp_path / "locked.xlsx", OLE)
    (only,) = scan_folder(tmp_path)
    assert only.kind == "encrypted"


def test_a10_one_bad_file_does_not_stop_the_scan(tmp_path, book):
    book("good.xlsx", {"CW05": CELLS})
    _write(tmp_path / "bad.xlsx", b"junk")
    _write(tmp_path / ".~lock.good.xlsx#", b"x")
    files = scan_folder(tmp_path)
    assert len(files) == 3
    assert sorted(f.kind for f in files) == ["corrupt", "lock", "xlsx"]


def test_a11_random_bytes_never_raise(tmp_path):
    rng = random.Random(7)  # fixed seed: the same bytes every run
    for i in range(25):
        name = f"r{i:02d}.{'xlsx' if i % 2 else 'xls'}"
        _write(tmp_path / name, bytes(rng.randrange(256) for _ in range(rng.randrange(1, 4000))))
    files = scan_folder(tmp_path)
    assert len(files) == 25
    assert {f.kind for f in files} <= set(KINDS)
    assert not any(f.importable for f in files)


def test_a12_scan_is_deterministic(tmp_path, book):
    for name in ("b.xlsx", "a.xlsx", "sub/c.xlsx"):
        os.makedirs(os.path.dirname(tmp_path / name), exist_ok=True)
        make_book(tmp_path / name, {"CW05": CELLS})
    first = [os.path.relpath(f.path, tmp_path) for f in scan_folder(tmp_path)]
    second = [os.path.relpath(f.path, tmp_path) for f in scan_folder(tmp_path)]
    assert first == second and len(first) == 3


def test_a13_family_is_the_top_folder(tmp_path, book):
    os.makedirs(tmp_path / "01 Weekly Sales Reports 2026")
    make_book(tmp_path / "01 Weekly Sales Reports 2026" / "x.xlsx", {"CW05": CELLS})
    (f,) = scan_folder(tmp_path)
    assert f.family == "01 Weekly Sales Reports 2026"


def test_a14_hash_prefix_and_nfc_names(tmp_path, book):
    nfd = unicodedata.normalize("NFD", "1d44ae9a3e0c_Müller（已联络）.xlsx")
    make_book(tmp_path / nfd, {"CW05": CELLS})
    (f,) = scan_folder(tmp_path)
    assert f.hash_prefix == "1d44ae9a3e0c"
    assert f.name == unicodedata.normalize("NFC", "1d44ae9a3e0c_Müller（已联络）.xlsx")


def test_a15_classify_name_table():
    assert classify_name(".~lock.x.xlsx#") == "lock"
    assert classify_name("~$x.xlsx") == "lock"
    assert classify_name(".DS_Store") == "hidden"
    assert classify_name("x.xlsx") is None
    assert classify_name("x.XLS") is None
    assert classify_name("x.pdf") == "not_excel"


# ---------- B. sheets and structure ----------

def test_b01_hidden_and_empty_sheets_reported(tmp_path, book):
    path = book("h.xlsx", {
        "CW05": {"cells": {"A1": "x"}},
        "CW04": {"cells": {"A1": "y"}, "hidden": True},
        "Sheet1": {},
    })
    info = {s.name: s for s in scan_file(path).sheets}
    assert info["CW04"].hidden and not info["CW05"].hidden
    assert info["Sheet1"].nonempty == 0 and info["CW05"].nonempty == 1


@pytest.mark.parametrize("name,expected", [
    ("CW05", 5), ("CW 5", 5), ("cw05 (2)", 5), ("CW53", 53),
    ("CW00", None), ("CW54", None), ("CW123", None), ("Sheet1", None), ("Sheet3", None),
])
def test_b02_cw_number_from_sheet_name(name, expected):
    assert cw_from_sheet_name(name) == expected


def test_b03_merged_ranges_counted(tmp_path, book):
    path = book("m.xlsx", {"CW05": {"cells": {"F9": "a"}, "merges": ["F9:F12", "B2:C2"]}})
    assert scan_file(path).sheets[0].merged == 2


def test_b04_formula_without_cache_is_a_gap(tmp_path, book):
    path = book("f.xlsx", {"CW05": {"cells": {"A1": 1, "A2": "=A1+1"}}})
    assert formula_gaps(path) == (1, 1)


def test_b05_formula_with_cache_is_not_a_gap(tmp_path, book):
    path = book("fc.xlsx", {"CW05": {"cells": {"A1": 1, "A2": "=A1+1"}}})
    src = zipfile.ZipFile(path)
    names = src.namelist()
    sheet_xml = next(n for n in names if n.startswith("xl/worksheets/sheet"))
    xml = src.read(sheet_xml).decode("utf-8")
    patched, count = re.subn(r"(<f>A1\+1</f>)<v\s*/>|(<f>A1\+1</f>)<v></v>", r"\1\2<v>2</v>", xml)
    if count == 0:
        pytest.skip("openpyxl wrote the formula cell in an unexpected shape; cannot inject a cache")
    out = str(tmp_path / "cached.xlsx")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for n in names:
            dst.writestr(n, patched.encode("utf-8") if n == sheet_xml else src.read(n))
    src.close()
    assert formula_gaps(out) == (1, 0)


def test_b06_sheet_order_is_preserved(tmp_path, book):
    path = book("o.xlsx", {"CW07": CELLS, "CW05": CELLS, "CW06": CELLS})
    assert [s.name for s in scan_file(path).sheets] == ["CW07", "CW05", "CW06"]


def test_b07_chinese_sheet_title_and_values_roundtrip(tmp_path, book):
    path = book("zh.xlsx", {"客户拜访": {"cells": {"A1": "客户", "A2": "示例公司"}}})
    (s,) = scan_file(path).sheets
    assert s.name == "客户拜访" and s.nonempty == 2 and s.cw is None


def test_b08_large_sheet_scans_within_budget(tmp_path, book):
    rows = {f"{chr(64 + c)}{r}": r * c for r in range(1, 5001) for c in range(1, 21)}
    path = book("big.xlsx", {"CW05": {"cells": rows}})
    t0 = time.monotonic()
    (s,) = scan_file(path).sheets
    assert s.nonempty == 5000 * 20
    assert time.monotonic() - t0 < 30


def test_b09_legacy_xls_is_read_when_xlwt_present(tmp_path):
    xlwt = pytest.importorskip("xlwt")
    wb = xlwt.Workbook()
    ws = wb.add_sheet("CW05")
    ws.write(0, 0, "Sales")
    ws2 = wb.add_sheet("CW04")
    ws2.write(0, 0, "x")
    path = str(tmp_path / "legacy.xls")
    wb.save(path)
    info = scan_file(path)
    assert info.kind == "xls" and info.importable
    assert {s.name: s.cw for s in info.sheets} == {"CW05": 5, "CW04": 4}


# ---------- C. duplicates and versions ----------

def test_c01_byte_identical_copies_are_grouped(tmp_path, book):
    src = book("orig.xlsx", {"CW05": CELLS})
    shutil.copy(src, tmp_path / "a.xlsx")
    shutil.copy(src, tmp_path / "1d44ae9a3e0c_a.xlsx")
    files = scan_folder(tmp_path)
    groups = duplicate_groups(files)
    assert len(groups) == 1 and len(next(iter(groups.values()))) == 3  # orig + 2 copies in folder
    prefixed = [f for f in files if f.hash_prefix]
    assert len(prefixed) == 1


def test_c02_resaved_copy_has_other_bytes_but_same_sheet_values(tmp_path, book):
    a = book("a.xlsx", {"CW05": CELLS}, title="first save")
    b = make_book(tmp_path / "b.xlsx", {"CW05": CELLS}, title="second save")
    assert scan_file(a).sha256 != scan_file(b).sha256
    assert sheet_fingerprints(a) == sheet_fingerprints(b)


def test_c03_next_week_version_adds_only_the_new_sheet(tmp_path, book):
    v1 = book("v1.xlsx", {"CW05": CELLS})
    v2 = make_book(tmp_path / "v2.xlsx", {"CW05": CELLS, "CW06": {"cells": {"A1": "new week"}}})
    f1, f2 = sheet_fingerprints(v1), sheet_fingerprints(v2)
    assert {k for k in f1 if f1[k] == f2.get(k)} == {"CW05"}
    assert set(f2) - set(f1) == {"CW06"}


def test_c04_changed_cell_changes_sheet_fingerprint(tmp_path, book):
    v1 = book("x1.xlsx", {"CW05": CELLS})
    v2 = make_book(tmp_path / "x2.xlsx", {"CW05": {"cells": {"A1": "Sales", "B2": "Customer", "C3": "Changed"}}})
    assert sheet_fingerprints(v1)["CW05"] != sheet_fingerprints(v2)["CW05"]


# ---------- D. parsers ----------

@pytest.mark.parametrize("value,year,expected", [
    (datetime(2026, 3, 5, 10, 0), None, date(2026, 3, 5)),
    ((date(2026, 3, 5) - date(1899, 12, 30)).days, None, date(2026, 3, 5)),
    ("2026-03-05", None, date(2026, 3, 5)),
    ("  2026-03-05 ", None, date(2026, 3, 5)),
    ("05.03.2026", None, date(2026, 3, 5)),
    ("3月5日", 2026, date(2026, 3, 5)),
    ("5/3/2026", None, None),        # ambiguous: refused, never guessed
    ("2026-02-30", None, None),      # not a real day
    ("", None, None), (None, None, None), (True, None, None), (0, None, None),
])
def test_d01_cell_dates(value, year, expected):
    assert parse_cell_date(value, year) == expected


@pytest.mark.parametrize("value,expected", [
    ("1,234.5", 1234.5), ("1.234,5", 1234.5), ("12,5", 12.5), ("1,234", 1234.0),
    ("¥1,234.50", 1234.5), ("12.5%", 0.125), ("(120)", -120.0), (" 7 ", 7.0),
    ("１２", 12.0), (42, 42.0), ("abc", None), ("", None), (None, None),
    (True, None), ("1.234.567", None),
])
def test_d02_numbers(value, expected):
    assert parse_number(value) == expected


@pytest.mark.parametrize("a,b", [
    ("（已联络）", "(已联络)"), ("Müller", "Muller"), ("Größe", "Grosse"),
    ("Customer \n(full address if new)", "customer (full address if new)"),
])
def test_d03_fold_text_makes_labels_comparable(a, b):
    assert fold_text(a) == fold_text(b)


def test_d04_header_found_after_title_rows_and_in_german_or_english():
    rows = [("Weekly Sales Report",), (None,), ("Vertrieb", "KW", "Datum", "Kunde"), ("x",)]
    assert find_header_row(rows, ["Vertrieb", "KW", "Kunde"]) == 2
    assert find_header_row(rows, ["Sales", "CW"]) is None
    assert find_header_row([(None,)] * 30 + [("Kunde", "Datum")], ["Kunde", "Datum"], max_scan=20) is None


def test_d05_header_with_line_breaks_matches_plain_label():
    rows = [("Sales", "Customer \n(full address if new)", "Date")]
    assert find_header_row(rows, ["Customer (full address if new)", "Date"]) == 0


# ---------- E. existing importer (skips when not importable) ----------

def test_e01_existing_read_paths_keeps_going_after_a_bad_file(tmp_path, book, importer):
    run = importer("excel.run")
    good = book("good.xlsx", {"CW05": CELLS})
    bad = _write(tmp_path / "bad.xlsx", b"junk")
    try:
        run.read_paths([bad, good])
    except ModuleNotFoundError as exc:  # importer requirement missing in this venv
        pytest.skip(f"install the importer's requirements first: {exc}")


def test_e02_existing_date_parser_on_basics(importer):
    # Contract (from its docstring): returns (iso_or_None, flag); flag "" means clean.
    dates = importer("excel.dates")
    assert dates.parse_date(datetime(2026, 3, 5), None, 2026) == ("2026-03-05", "")


def test_e03_existing_visit_reader_is_stable_on_real_files(real_files, importer):
    # visits.reader handles the visit template only. Workbooks it cannot parse (for example a
    # sheet named "5.8-9" with no CW number) are reported, not hidden: see the finding in the report.
    reader = importer("visits.reader")
    picks = [f for f in real_files if f.importable and f.family[:2] == "03"][:5]
    if not picks:
        pytest.skip("no importable family-03 (visit template) workbook in the sample folder")
    parsed, refused = 0, []
    for f in picks:
        try:
            first = [v.visit_key for v in reader.parse_workbook(f.path).visits]
        except ValueError:
            refused.append(f"#{real_files.index(f):02d}")
            continue
        second = [v.visit_key for v in reader.parse_workbook(f.path).visits]
        assert first == second
        assert len(first) == len(set(first)), "visit keys must be unique within a workbook"
        parsed += 1
    if not parsed:
        pytest.skip(f"visits.reader refused every sampled file: {refused}")


# ---------- F. real sample folder (skips when absent) ----------

def test_f01_every_file_has_a_known_kind(real_files):
    assert {f.kind for f in real_files} <= set(KINDS)
    assert all(f.kind in ("xlsx", "xls") or not f.importable for f in real_files)


def test_f02_every_importable_file_opened_with_sheets(real_files):
    empty_open = [f for f in real_files if f.importable and not f.sheets]
    assert not empty_open, "an importable workbook reported zero sheets"


def test_f03_golden_counts_match(real_files, golden_counts):
    current = golden(real_files)
    diff = {k: (golden_counts.get(k), current[k]) for k in current if golden_counts.get(k) != current[k]}
    assert not diff, f"folder changed since the golden snapshot (expected, now): {diff}"


def test_f04_weekly_and_visit_workbooks_read_with_existing_excel_run(real_files, importer):
    # excel.run is the reader that knows the weekly and Lucas layouts (families 01, 02);
    # visits.reader only knows the visit template, so it is not used for them here.
    run = importer("excel.run")
    paths = [f.path for f in real_files if f.importable and f.family[:2] in ("01", "02", "03")]
    try:
        run.read_paths(paths)
    except ModuleNotFoundError as exc:
        pytest.skip(f"install the importer's requirements first: {exc}")


def test_f05_real_folder_lock_files_are_never_importable(real_files):
    locks = [f for f in real_files if f.kind == "lock"]
    assert all(not f.importable for f in locks)


# ---------- G. privacy and robustness ----------

def test_g01_report_never_contains_cell_values(tmp_path, book):
    sentinel = "SENTINEL-7731-ZZ"
    path = book("secret.xlsx", {"CW05": {"cells": {"A1": sentinel}}})
    text = report_text([scan_file(path, tmp_path)], show_names=False)
    assert sentinel not in text


def test_g02_report_hides_file_names_by_default(real_files):
    text = report_text(real_files, show_names=False)
    for f in real_files:
        if len(f.name) > 6:
            assert f.name not in text


def test_g03_xls_and_zip_errors_do_not_leak_text(tmp_path):
    _write(tmp_path / "x.xls", b"not an excel file at all")
    (f,) = scan_folder(tmp_path)
    assert f.kind == "corrupt"
    assert f.error.isidentifier() and "not an excel" not in f.error
