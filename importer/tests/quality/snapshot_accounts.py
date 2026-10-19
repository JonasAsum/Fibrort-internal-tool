"""Record-level snapshot of what the accounts reader produces for folders 04 and 05.

Four kinds of record, one JSON line each, sorted by a stable key: `row` (one list row as read, before
merging), `account` and `lead` (after the merge across files), `review` (every skipped or questioned item).
It detects change only; a snapshot is never proof that the output is correct.

    .venv/bin/python -m tests.quality.snapshot_accounts write     # freeze the current output
    .venv/bin/python -m tests.quality.snapshot_accounts diff      # what changed since the freeze
"""
import os
import sys
from collections import Counter

from excel import accounts
from excel.files import unique_contents
from excel.sheet import load
from tests.quality.family_common import GOLDEN_DIR, SAMPLE_DIR, family_files, main, plain

FOLDERS = ("04", "05")
GOLDEN = os.path.join(GOLDEN_DIR, "accounts.jsonl")


def files(sample_dir=SAMPLE_DIR):
    return family_files(FOLDERS, sample_dir)


def rows_of(paths, sample_dir=SAMPLE_DIR) -> dict:
    """{key: row} of every list row as the sheet readers see it, before rows are merged into accounts."""
    out = {}
    for path in unique_contents(paths, []):
        base = os.path.basename(path)
        for sheet in load(path):
            layout, got, _, _ = accounts.read_accounts_sheet(sheet, path)
            for r in got:
                out[f"row|{base}|{sheet.name}|{r.row:05d}|{r.layout}"] = {"type": "row", **plain(r, sample_dir)}
    return out


def current(sample_dir=SAMPLE_DIR) -> dict:
    paths = files(sample_dir)
    out = rows_of(paths, sample_dir)
    result = accounts.read_paths(paths)
    for a in result.accounts:
        out[f"account|{a.key}"] = {"type": "account", **plain(a, sample_dir)}
    for l in result.leads:
        out[f"lead|{l.key}"] = {"type": "lead", **plain(l, sample_dir)}
    seen = Counter()
    for i in result.review:
        base = f"review|{os.path.basename(i.file)}|{i.sheet}|{i.row:05d}|{i.field}|{i.reason}"
        seen[base] += 1
        out[f"{base}|{seen[base]}"] = {"type": "review", **plain(i, sample_dir)}
    return out


if __name__ == "__main__":
    sys.exit(main("accounts", current, GOLDEN))
