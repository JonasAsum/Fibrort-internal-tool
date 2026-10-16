"""Re-read the importer's answer for every distinct customer text (after a reader change), keeping the same
cell ids, so the two AI labellers' answers stay valid and only the importer side is recomputed.

    .venv/bin/python -m tests.quality.l2_refresh
"""
import json
import os

from tests.quality import snapshot
from tests.quality.ledger import Ledger

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "l2")
TRUTH = os.path.join(HERE, "customer_truth_hidden.json")


def refresh(records=None):
    records = records if records is not None else snapshot.current()
    paths = {os.path.basename(p): p for p in snapshot.visit_files()}
    ledgers, now = {}, {}
    for key, r in records.items():
        if r.get("type") != "visit":
            continue
        coord = (r.get("extra") or {}).get("source_cells", {}).get("customer")
        if not coord:
            continue
        base = key.split("|")[0]
        if base not in ledgers:
            ledgers[base] = Ledger(paths[base])
        raw = ledgers[base].text(r["sheet"], coord)
        if raw.strip():
            now[raw] = {"name": r.get("customer_raw") or "", "name_en": (r["extra"].get("name_en") or ""),
                        "address": (r["extra"].get("address") or ""), "code": (r["extra"].get("customer_code") or "")}
    truth = json.load(open(TRUTH, encoding="utf-8"))
    changed = 0
    for cid, t in truth.items():
        new = now.get(t["text"])
        if new is not None and new != t["reader"]:
            t["reader"] = new
            changed += 1
    missing = [t["text"] for t in truth.values() if t["text"] not in now]
    json.dump(truth, open(TRUTH, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    return changed, missing, len(now) - len(truth)


if __name__ == "__main__":
    changed, missing, new_texts = refresh()
    print(f"importer answers changed for {changed} cell texts; texts no longer in the files: {len(missing)}; new texts not yet labelled: {new_texts}")
