"""Record-level snapshot of what the visit readers produce, for every visit workbook (folders 01-03).

One JSON line per record, sorted by a stable key, every field kept. `write` freezes today's output,
`diff` says exactly which records and fields changed after a code change, grouped by field.
It detects change only: a snapshot is never proof that the output is correct.

    .venv/bin/python -m tests.quality.snapshot write            # freeze the current output
    .venv/bin/python -m tests.quality.snapshot diff             # what changed since the freeze
"""
import dataclasses
import json
import os
import sys
from collections import defaultdict

from excel.run import read_paths

HERE = os.path.dirname(os.path.abspath(__file__))
GOLDEN = os.path.join(os.path.dirname(HERE), "golden_records", "visits.jsonl")
SAMPLE_DIR = os.environ.get("EXCEL_SAMPLE_DIR") or os.path.abspath(
    os.path.join(os.path.dirname(HERE), "..", "..", "..", "CRM Excel Lists 2026-10-11"))
VISIT_FOLDERS = ("01", "02", "03")


def visit_files(sample_dir=SAMPLE_DIR):
    out = []
    for folder in sorted(os.listdir(sample_dir)):
        if folder[:2] not in VISIT_FOLDERS or not os.path.isdir(os.path.join(sample_dir, folder)):
            continue
        for name in sorted(os.listdir(os.path.join(sample_dir, folder))):
            if name.startswith((".", "~$")) or not name.lower().endswith((".xlsx", ".xlsm", ".xls")):
                continue
            out.append(os.path.join(sample_dir, folder, name))
    return out


def _plain(obj):
    return json.loads(json.dumps(dataclasses.asdict(obj) if dataclasses.is_dataclass(obj) else obj,
                                 ensure_ascii=False, default=str))


def records_of(path) -> dict:
    """{stable key: record} for one workbook, read on its own so copies in other files do not interfere."""
    base = os.path.basename(path)
    result = read_paths([path])
    out = {}
    for v in result.parsed.visits:
        out[f"{base}|{v.sheet}|visit|{v.excel_row_start}"] = {"type": "visit", **_plain(v)}
    for w in result.parsed.weeks:
        out[f"{base}|{w.sheet}|week"] = {"type": "week", **_plain(w)}
    for i, s in enumerate(result.superseded):
        out[f"{base}|superseded|{i}"] = {"type": "superseded", **_plain(s)}
    return out


def current(sample_dir=SAMPLE_DIR) -> dict:
    out = {}
    for path in visit_files(sample_dir):
        out.update(records_of(path))
    return out


def write(records: dict, path=GOLDEN):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        for key in sorted(records):
            f.write(json.dumps({"key": key, **records[key]}, ensure_ascii=False, sort_keys=True) + "\n")


def load(path=GOLDEN) -> dict:
    out = {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            rec = json.loads(line)
            out[rec.pop("key")] = rec
    return out


def diff(old: dict, new: dict) -> dict:
    """{"added": [keys], "removed": [keys], "changed": {field: [(key, old, new)]}}"""
    changed = defaultdict(list)
    for key in sorted(old.keys() & new.keys()):
        a, b = old[key], new[key]
        for fld in sorted(a.keys() | b.keys()):
            if fld == "extra":
                for sub in sorted((a.get(fld) or {}).keys() | (b.get(fld) or {}).keys()):
                    x, y = (a.get(fld) or {}).get(sub), (b.get(fld) or {}).get(sub)
                    if x != y:
                        changed[f"extra.{sub}"].append((key, x, y))
            elif a.get(fld) != b.get(fld):
                changed[fld].append((key, a.get(fld), b.get(fld)))
    return {"added": sorted(new.keys() - old.keys()), "removed": sorted(old.keys() - new.keys()), "changed": dict(changed)}


def report(d: dict, examples=2) -> str:
    lines = [f"records added: {len(d['added'])}, removed: {len(d['removed'])}, fields changed: {len(d['changed'])}"]
    for fld, items in sorted(d["changed"].items(), key=lambda t: -len(t[1])):
        lines.append(f"  {fld}: {len(items)} record(s)")
        for key, a, b in items[:examples]:
            lines.append(f"      {key}\n        was: {str(a)[:90]!r}\n        now: {str(b)[:90]!r}")
    return "\n".join(lines)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "diff"
    now = current()
    if cmd == "write":
        write(now)
        print(f"froze {len(now)} records to {GOLDEN}")
    else:
        d = diff(load(), now)
        print(report(d))
        sys.exit(1 if d["added"] or d["removed"] or d["changed"] else 0)
