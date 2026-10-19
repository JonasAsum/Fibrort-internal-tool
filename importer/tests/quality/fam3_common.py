"""Helpers shared by the todos / orders / receivables oracles. Built on the cell ledger only (openpyxl / xlrd);
no importer code is imported here.

`workbook_cells(path)` is the ledger's `cells()` with one fallback: a workbook whose styles.xml holds a
<cellStyle> without a name (seen in ERP exports) is refused by openpyxl, so it is read from a repaired copy
in a temp file (the file on disk is never touched). The cells keep the real file name.
"""
import dataclasses
import io
import os
import re
import tempfile
import zipfile
from decimal import ROUND_HALF_UP, Decimal

from tests.quality import ledger

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.environ.get("EXCEL_SAMPLE_DIR") or os.path.abspath(
    os.path.join(os.path.dirname(HERE), "..", "..", "..", "CRM Excel Lists 2026-10-11"))
_NAMELESS = re.compile(r"<cellStyle(?![^>]*\bname=)[^>]*/>")
CENT = Decimal("0.01")


def folder_files(prefix, sample_dir=SAMPLE_DIR):
    """Absolute paths of the workbooks in the numbered folder ("07"), sorted."""
    out = []
    for folder in sorted(os.listdir(sample_dir)):
        full = os.path.join(sample_dir, folder)
        if folder[:2] == prefix and os.path.isdir(full):
            out += [os.path.join(full, n) for n in sorted(os.listdir(full))
                    if not n.startswith((".", "~$")) and n.lower().endswith((".xlsx", ".xlsm", ".xls"))]
    return out


def _repaired_copy(path):
    with zipfile.ZipFile(path) as zin:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/styles.xml":
                    data = _NAMELESS.sub("", data.decode("utf8")).encode("utf8")
                zout.writestr(item, data)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    with os.fdopen(fd, "wb") as f:
        f.write(buf.getvalue())
    return tmp


_CACHE = {}


def workbook_cells(path, merged_out=None):
    """List of ledger Cells of the workbook; file name is always the real base name. The result of a path is
    kept for the session (the files do not change during a test run); do not modify the returned list."""
    if path not in _CACHE:
        base = os.path.basename(path)
        merged = {}
        try:
            got = list(ledger.cells(path, merged))
        except TypeError:  # nameless <cellStyle>: read a repaired copy
            tmp = _repaired_copy(path)
            try:
                merged = {}
                got = [dataclasses.replace(c, file=base) for c in ledger.cells(tmp, merged)]
            finally:
                os.unlink(tmp)
        _CACHE[path] = (got, merged)
    got, merged = _CACHE[path]
    if merged_out is not None:
        merged_out.update(merged)
    return got


def display_name(path):
    """File name without the 12-hex-digit prefix the folder export puts in front of some."""
    return re.sub(r"^[0-9a-f]{12}_", "", os.path.basename(path))


def norm(text) -> str:
    """Comparison form of a cell text: whitespace collapsed, stripped."""
    return re.sub(r"\s+", " ", str(text)).strip()


def dec(text):
    """Decimal of a ledger text, or None when it is not a plain number."""
    try:
        d = Decimal(str(text).strip())
    except Exception:
        return None
    return d if d.is_finite() else None


def cent(d: Decimal) -> Decimal:
    return d.quantize(CENT, ROUND_HALF_UP)


def rows_by_sheet(cells):
    """{sheet: {row: {col: Cell}}}"""
    out = {}
    for c in cells:
        out.setdefault(c.sheet, {}).setdefault(c.row, {})[c.col] = c
    return out
