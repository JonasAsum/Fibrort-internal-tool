"""Record-level snapshot of what the order reader produces for folder 08: every raw line of every file
(`read_orders`), and the rolled-up orders, file verdicts and review items of `build_book`.

    .venv/bin/python -m tests.quality.snapshot_orders write
    .venv/bin/python -m tests.quality.snapshot_orders diff
"""
import os
import sys

from excel import orders
from tests.quality import snapshot

FOLDER = "08"
GOLDEN = os.path.join(os.path.dirname(snapshot.GOLDEN), "orders.jsonl")


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
        res = orders.read_orders(p)
        for ln in res.lines:
            out[f"line|{base}|{ln.sheet}|{ln.row}"] = {"type": "line", **snapshot._plain(ln)}
        for no, d in res.documents.items():
            out[f"document|{base}|{no}"] = {"type": "document", **snapshot._plain(d)}
        out[f"read|{base}"] = {"type": "read", "kind": res.kind, "reason": res.reason, "layout": res.layout,
                               "headers": snapshot._plain(res.headers_found), "warnings": res.warnings,
                               "skipped": snapshot._plain(res.skipped), "skipped_sheets": snapshot._plain(res.skipped_sheets),
                               "excluded_sheets": snapshot._plain(res.excluded_sheets), "export_time": str(res.export_time)}
    book = orders.build_book(paths)
    for o in book.orders:
        out[f"order|{o.source_id}"] = {"type": "order", **snapshot._plain(o)}
    for f in book.files:
        out[f"file|{f.name}|{os.path.basename(f.path)}"] = {"type": "file", **snapshot._plain(f)}
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
