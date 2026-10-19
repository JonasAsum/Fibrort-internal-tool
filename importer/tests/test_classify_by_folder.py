"""Every real workbook is classified into the list kind its folder says it is.

The folder is the expectation (01-03 visits, 04-05 accounts, 06 contacts, 07 todos, 08 orders,
09 receivables). The one exception: 08 and 09 may hold files the folder review keeps out of the
CRM on purpose (kept_out), which is also a correct answer.
"""
import os

import pytest

from webui.entry import classify

EXPECTED = {
    "01": "visits", "02": "visits", "03": "visits",
    "04": "accounts", "05": "accounts",
    "06": "contacts", "07": "todos", "08": "orders", "09": "receivables",
}
KEPT_OUT_OK = {"08", "09"}


def _files(sample_dir):
    for folder in sorted(os.listdir(sample_dir)):
        if folder[:2] not in EXPECTED or not os.path.isdir(os.path.join(sample_dir, folder)):
            continue
        for name in sorted(os.listdir(os.path.join(sample_dir, folder))):
            if name.startswith(("~$", ".~lock", ".")) or not name.lower().endswith((".xlsx", ".xlsm", ".xls")):
                continue
            yield folder[:2], os.path.join(sample_dir, folder, name)


def test_every_file_is_classified_as_its_folder_says(sample_dir):
    wrong = []
    count = 0
    for folder, path in _files(sample_dir):
        count += 1
        got = classify(path, folder)
        want = EXPECTED[folder]
        ok = (got["kind"] == "import" and got["family"] == want) or (folder in KEPT_OUT_OK and got["kind"] == "kept_out")
        if not ok:
            wrong.append(f"{folder} {os.path.basename(path)}: expected {want}, got {got['kind']}/{got['family']} ({got['what']})")
    assert count, "no workbooks found in folders 01-09"
    assert not wrong, f"{len(wrong)} of {count} files classified against their folder:\n  " + "\n  ".join(wrong)
