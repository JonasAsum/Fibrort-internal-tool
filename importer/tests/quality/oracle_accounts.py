"""Independent oracle for the account lists (folders 04 and 05): for every row of every workbook, what kind
of row it is and, for every cell, the header that stands over it. Built from the cell ledger only.

A header row is found by its own words, whatever the reader thinks:
  erp           客户编码 + 客户名称                         (ERP customer export)
  master        Customer Code + CN name
  supporting    a column "FIBRO Existing Customer"
  effective     "English name" + 客户名称 / 厂家名称      (several lists on one sheet)
  registration  "Customer Name" + "Location"
Rows above the first header are titles. In an `effective` sheet a lone non-numeric cell in column A is the
heading of the next list, a lone number is an empty numbered line. Everything else under a header is data.
"""
import os
from collections import defaultdict

from tests.quality.ledger import cells as ledger_cells
from tests.quality.oracle_common import OCell, ORow, flat, is_seq


def header_layout(texts) -> str:
    low = [flat(t).casefold() for t in texts]
    has = lambda f: any(f(t) for t in low)
    if has(lambda t: t == "客户编码") and has(lambda t: t == "客户名称"):
        return "erp"
    if has(lambda t: t == "customer code") and has(lambda t: t == "cn name"):
        return "master"
    if has(lambda t: "fibro existing customer" in t):
        return "supporting"
    if has(lambda t: "english name" in t) and has(lambda t: "客户名称" in t or "厂家名称" in t):
        return "effective"
    if has(lambda t: t.startswith("customer name")) and has(lambda t: t.startswith("location")):
        return "registration"
    return ""


def rows_of(path) -> list:
    """Every non-empty row of one workbook as an ORow."""
    base = os.path.basename(path)
    by_row = defaultdict(list)
    for c in ledger_cells(path):
        by_row[(c.sheet, c.row)].append(c)
    out = []
    for sheet in sorted({s for s, _ in by_row}, key=lambda s: min(r for (sh, r) in by_row if sh == s)):
        layout, header_row, labels = "", 0, {}
        for r in sorted(r for (sh, r) in by_row if sh == sheet):
            row = sorted(by_row[(sheet, r)], key=lambda c: c.col)
            found = header_layout(c.text for c in row)
            if found:
                layout, header_row = found, r
                labels = {c.col: flat(c.text) for c in row}
                kind = "header"
            elif not layout:
                kind = "title"
            elif layout == "effective" and len(row) == 1 and row[0].col == 1 and not is_seq(row[0].text):
                kind = "heading"
            elif len(row) == 1 and row[0].col == 1 and is_seq(row[0].text):
                kind = "seq_only"
            else:
                kind = "data"
            out.append(ORow(base, sheet, r, kind, layout, header_row,
                            tuple(OCell(base, sheet, r, c.col, c.coord, c.text,
                                        "" if kind in ("header", "title", "heading") else labels.get(c.col, ""),
                                        c.hidden) for c in row)))
    return out


def oracle(paths) -> list:
    out = []
    for p in paths:
        out += rows_of(p)
    return out
