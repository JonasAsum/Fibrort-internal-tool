"""Record-level snapshot of what the to-do reader produces for folder 07 (see snapshot.py for the idea).

    .venv/bin/python -m tests.quality.snapshot_todos write      # freeze the current output
    .venv/bin/python -m tests.quality.snapshot_todos diff       # what changed since the freeze
"""
import os
import sys

from excel import todos
from tests.quality import snapshot

FOLDER = "07"
GOLDEN = os.path.join(os.path.dirname(snapshot.GOLDEN), "todos.jsonl")


def files(sample_dir=snapshot.SAMPLE_DIR):
    out = []
    for folder in sorted(os.listdir(sample_dir)):
        full = os.path.join(sample_dir, folder)
        if folder[:2] == FOLDER and os.path.isdir(full):
            out += [os.path.join(full, n) for n in sorted(os.listdir(full))
                    if not n.startswith((".", "~$")) and n.lower().endswith((".xlsx", ".xlsm", ".xls"))]
    return out


def records_of(result) -> dict:
    out = {}
    for t in result.todos:
        out[f"todo|{t.source_id}"] = {"type": "todo", **snapshot._plain(t)}
    seen = {}
    for i in result.review:
        base = f"review|{i.file}|{i.sheet}|{i.row}|{i.field}|{i.reason}"
        seen[base] = seen.get(base, 0) + 1
        out[f"{base}|{seen[base]}"] = {"type": "review", **snapshot._plain(i)}
    for s in result.per_file_stats:
        out[f"file|{os.path.basename(s.path)}"] = {"type": "file", **snapshot._plain(s)}
    for s in result.skipped_files:
        out[f"skipped|{os.path.basename(s['file'])}"] = {"type": "skipped", **snapshot._plain(s)}
    return out


def current(sample_dir=snapshot.SAMPLE_DIR) -> dict:
    return records_of(todos.read_paths(files(sample_dir)))


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
