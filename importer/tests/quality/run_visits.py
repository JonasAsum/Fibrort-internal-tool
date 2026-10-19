"""Run the visit invariants over every record of the frozen snapshot (or the current output with --live).

    .venv/bin/python -m tests.quality.run_visits            # check the frozen snapshot
    .venv/bin/python -m tests.quality.run_visits --live     # read the workbooks again and check that
"""
import collections
import json
import os
import sys

from tests.quality import snapshot
from tests.quality.invariants import check_visit
from tests.quality.ledger import Ledger


def run(records: dict) -> list:
    paths = {os.path.basename(p): p for p in snapshot.visit_files()}
    ledgers, out = {}, []
    for key, rec in records.items():
        if rec.get("type") != "visit":
            continue
        base = key.split("|")[0]
        if base not in ledgers:
            ledgers[base] = Ledger(paths[base])
        out += check_visit({**rec, "_file": base}, ledgers[base])
    return out


def summary(violations) -> str:
    by_rule = collections.Counter(v["rule"] for v in violations)
    rows = {(v["file"], v["sheet"], v["row"]) for v in violations}
    lines = [f"{n:5}  {rule}" for rule, n in by_rule.most_common()]
    return "\n".join(lines + [f"records with at least one violation: {len(rows)}"])


if __name__ == "__main__":
    recs = snapshot.current() if "--live" in sys.argv else snapshot.load()
    viol = run(recs)
    print(f"visit records checked: {sum(1 for r in recs.values() if r.get('type') == 'visit')}")
    print(summary(viol))
    if "--json" in sys.argv:
        print(json.dumps(viol, ensure_ascii=False))
    sys.exit(1 if viol else 0)
