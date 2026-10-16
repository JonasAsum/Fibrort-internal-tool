"""Shared shapes of the family oracles. An oracle reads the workbooks with tests.quality.ledger only
(openpyxl / xlrd), never with an importer reader, and says for every row what it is and for every cell
which column header stands over it."""
from collections import Counter
from dataclasses import dataclass, field


@dataclass(frozen=True)
class OCell:
    file: str
    sheet: str
    row: int
    col: int
    coord: str
    text: str
    label: str  # the header text over this column ("" when the column has none)
    hidden: bool = False


@dataclass
class ORow:
    file: str
    sheet: str
    row: int
    kind: str  # header | title | heading | seq_only | data
    layout: str  # the layout of the header above ("" before the first header)
    header_row: int  # row number of the header that governs this row (0 = none)
    cells: tuple = field(default_factory=tuple)

    @property
    def key(self):
        return (self.file, self.sheet, self.row)


def flat(text) -> str:
    return " ".join(str(text or "").split())


def is_seq(text) -> bool:
    return flat(text).isdigit()


def summary(rows) -> dict:
    """Counts by kind: rows and cells, plus the files and sheets seen."""
    out = Counter()
    for r in rows:
        out[f"rows_{r.kind}"] += 1
        out[f"cells_{r.kind}"] += len(r.cells)
        out["cells_total"] += len(r.cells)
    out["files"] = len({r.file for r in rows})
    out["sheets"] = len({(r.file, r.sheet) for r in rows})
    return dict(out)
