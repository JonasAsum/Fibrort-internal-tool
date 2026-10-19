"""Record-level snapshot of what the receivables reader produces for folder 09: every invoice row of every
AR sheet, the per-customer totals of the chosen snapshot, file verdicts and review items.

    .venv/bin/python -m tests.quality.snapshot_receivables write
    .venv/bin/python -m tests.quality.snapshot_receivables diff
"""
import os
import sys

from excel import receivables
from tests.quality import snapshot

FOLDER = "09"
GOLDEN = os.path.join(os.path.dirname(snapshot.GOLDEN), "receivables.jsonl")


def files(sample_dir=snapshot.SAMPLE_DIR):
    out = []
    for folder in sorted(os.listdir(sample_dir)):
        full = os.path.join(sample_dir, folder)
        if folder[:2] == FOLDER and os.path.isdir(full):
            out += [os.path.join(full, n) for n in sorted(os.listdir(full))
                    if not n.startswith((".", "~$")) and n.lower().endswith((".xlsx", ".xlsm", ".xls"))]
    return out


def records_of(paths) -> dict:
    out = {}
    for p in paths:
        base = os.path.basename(p)
        info = receivables.FileInfo(path=p, name=receivables.file_display_name(p))
        for sh in receivables.read_workbook(p, info):
            out[f"sheet|{base}|{sh.sheet}"] = {
                "type": "sheet", "currency": sh.currency, "as_of": sh.as_of, "as_of_source": sh.as_of_source,
                "report_date": sh.report_date, "skipped_rows": sh.skipped_rows, "claimed_open": str(sh.claimed_open),
                "claimed_overdue": str(sh.claimed_overdue), "open_total": str(sh.open_total),
                "overdue_total": str(sh.overdue_total), "warnings": sh.warnings}
            for inv in sh.invoices:
                out[f"invoice|{base}|{sh.sheet}|{inv.row}"] = {"type": "invoice", **snapshot._plain(inv)}
        out[f"excluded|{base}"] = {"type": "excluded", "sheets": snapshot._plain(info.excluded_sheets),
                                    "read": info.sheets_read}
    book = receivables.build_book(paths)
    out["book"] = {"type": "book", "as_of": book.as_of, "source": book.source, "currency": book.currency}
    for k, c in book.customers.items():
        rec = snapshot._plain(c)
        rec["customer_key"] = rec.pop("key")  # `key` is the snapshot's own record key
        out[f"customer|{k}"] = {"type": "customer", **rec}
    for f in book.files:
        out[f"file|{os.path.basename(f.path)}"] = {"type": "file", **snapshot._plain(f)}
    seen = {}
    for i in book.review:
        base = f"review|{i.file}|{i.sheet}|{i.row}|{i.field}|{i.reason}"
        seen[base] = seen.get(base, 0) + 1
        out[f"{base}|{seen[base]}"] = {"type": "review", **snapshot._plain(i)}
    for s in book.skipped_files:
        out[f"skipped|{os.path.basename(s['file'])}"] = {"type": "skipped", **snapshot._plain(s)}
    return out


def current(sample_dir=snapshot.SAMPLE_DIR) -> dict:
    return records_of(files(sample_dir))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "diff"
    now = current()
    if cmd == "write":
        snapshot.write(now, GOLDEN)
        print(f"froze {len(now)} records to {GOLDEN}")
    else:
        d = snapshot.diff(snapshot.load(GOLDEN), now)
        print(snapshot.report(d))
        sys.exit(1 if d["added"] or d["removed"] or d["changed"] else 0)
