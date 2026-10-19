"""Shared helpers for the accounts and contacts snapshots: the files of a family, plain JSON of a record,
and the write / diff command line. The generic parts (write, load, diff, report) are those of snapshot.py."""
import dataclasses
import json
import os
import sys

from tests.quality import snapshot

HERE = os.path.dirname(os.path.abspath(__file__))
GOLDEN_DIR = os.path.join(os.path.dirname(HERE), "golden_records")
SAMPLE_DIR = snapshot.SAMPLE_DIR


def family_files(folders, sample_dir=SAMPLE_DIR):
    """Every workbook in the numbered folders (e.g. ("04", "05")), sorted, as absolute paths."""
    out = []
    for folder in sorted(os.listdir(sample_dir)):
        if folder[:2] not in folders or not os.path.isdir(os.path.join(sample_dir, folder)):
            continue
        for name in sorted(os.listdir(os.path.join(sample_dir, folder))):
            if name.startswith((".", "~$")) or not name.lower().endswith((".xlsx", ".xlsm", ".xls")):
                continue
            out.append(os.path.join(sample_dir, folder, name))
    return out


def plain(obj, sample_dir=SAMPLE_DIR):
    """A dataclass / dict as JSON-safe data; absolute paths inside the sample folder become file base names."""
    data = json.loads(json.dumps(dataclasses.asdict(obj) if dataclasses.is_dataclass(obj) else obj,
                                 ensure_ascii=False, default=str))

    def fix(x):
        if isinstance(x, str):
            return os.path.basename(x) if x.startswith(sample_dir) else x
        if isinstance(x, list):
            return [fix(i) for i in x]
        if isinstance(x, dict):
            return {k: fix(v) for k, v in x.items()}
        return x
    data = fix(data)
    if isinstance(data, dict) and "key" in data:  # snapshot lines use `key` for the record's own key
        data["record_key"] = data.pop("key")
    return data


def main(family, current, golden):
    cmd = sys.argv[1] if len(sys.argv) > 1 else "diff"
    now = current()
    if cmd == "write":
        snapshot.write(now, golden)
        print(f"froze {len(now)} {family} records to {golden}")
        return 0
    d = snapshot.diff(snapshot.load(golden), now)
    print(snapshot.report(d, examples=int(os.environ.get("EXAMPLES", "3"))))
    if os.environ.get("VERBOSE"):
        for fld, items in sorted(d["changed"].items()):
            for key, a, b in items:
                print(f"{fld}\t{key}\n   was {a!r}\n   now {b!r}")
        for k in d["added"]:
            print("added", k)
        for k in d["removed"]:
            print("removed", k)
    return 1 if d["added"] or d["removed"] or d["changed"] else 0
