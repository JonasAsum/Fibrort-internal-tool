"""Record-level snapshot of what the contacts reader produces for folder 06.

Kinds of record: `entry` (one list row as read, before merging), `contact` (merged people), `org`
(companies), `review` (every skipped or questioned item). Detects change only; never proof of correctness.

    .venv/bin/python -m tests.quality.snapshot_contacts write     # freeze the current output
    .venv/bin/python -m tests.quality.snapshot_contacts diff      # what changed since the freeze
"""
import os
import sys
from collections import Counter

from excel import contacts
from excel.files import unique_contents
from excel.sheet import load
from tests.quality.family_common import GOLDEN_DIR, SAMPLE_DIR, family_files, main, plain

FOLDERS = ("06",)
GOLDEN = os.path.join(GOLDEN_DIR, "contacts.jsonl")


def files(sample_dir=SAMPLE_DIR):
    return family_files(FOLDERS, sample_dir)


def entries_of(paths, sample_dir=SAMPLE_DIR) -> dict:
    out = {}
    for path in unique_contents(paths, []):
        base = os.path.basename(path)
        for sheet in load(path):
            layout, got, _, _ = contacts.read_contacts_sheet(sheet, path)
            for e in got:
                out[f"entry|{base}|{sheet.name}|{e.row:05d}"] = {"type": "entry", **plain(e, sample_dir)}
    return out


def current(sample_dir=SAMPLE_DIR) -> dict:
    paths = files(sample_dir)
    out = entries_of(paths, sample_dir)
    result = contacts.read_paths(paths)
    for c in result.contacts:
        out[f"contact|{c.source_id}"] = {"type": "contact", **plain(c, sample_dir)}
    for k, o in result.orgs.items():
        out[f"org|{k}"] = {"type": "org", **plain(o, sample_dir)}
    seen = Counter()
    for i in result.review:
        base = f"review|{os.path.basename(i.file)}|{i.sheet}|{i.row:05d}|{i.field}|{i.reason}"
        seen[base] += 1
        out[f"{base}|{seen[base]}"] = {"type": "review", **plain(i, sample_dir)}
    return out


if __name__ == "__main__":
    sys.exit(main("contacts", current, GOLDEN))
