"""Independent oracle for the contact lists (folder 06): for every row of every workbook, what kind of row it
is and, for every cell, the header that stands over it. Built from the cell ledger only.

The header row is the first row with three or more cells that are known contact-list column names
(Contact Person, E-mail Addr., Phone No., Position, Remark, 序号, 公司名称, 地址, 联系人, 电话 ...).
  list      the header has a company column (公司名称): every row under it is a company/person row.
  sections  no company column: a row with only a column-A cell is a heading (a team or a supplier group),
            any other row is a person whose column-A label (when given) names the team or supplier.
Rows above the header are titles.
"""
import os
from collections import defaultdict

from tests.quality.ledger import cells as ledger_cells
from tests.quality.oracle_common import OCell, ORow, flat

VOCAB = {"contact person", "e-mail addr.", "phone no.", "position", "remark", "序号", "公司名称", "地址", "联系人", "电话",
         "email", "e-mail", "phone", "company", "contact", "name", "title", "mobile", "手机", "邮箱", "职位", "备注"}
COMPANY = {"公司名称", "company"}


def rows_of(path) -> list:
    base = os.path.basename(path)
    by_row = defaultdict(list)
    for c in ledger_cells(path):
        by_row[(c.sheet, c.row)].append(c)
    out = []
    for sheet in sorted({s for s, _ in by_row}):
        layout, header_row, labels = "", 0, {}
        for r in sorted(r for (sh, r) in by_row if sh == sheet):
            row = sorted(by_row[(sheet, r)], key=lambda c: c.col)
            known = [c for c in row if flat(c.text).casefold() in VOCAB]
            if not layout and len(known) >= 3:
                layout = "list" if any(flat(c.text).casefold() in COMPANY for c in known) else "sections"
                header_row, labels, kind = r, {c.col: flat(c.text) for c in row}, "header"
            elif not layout:
                kind = "title"
            elif layout == "sections" and len(row) == 1 and row[0].col == 1:
                kind = "heading"
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
