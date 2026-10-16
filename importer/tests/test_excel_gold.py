"""Accuracy against the hand-checked golden set (excel/gold/gold.json, real data,
gitignored). Skips when the golden set is not on this machine. Run with -s to see
the per-family, per-field table."""
import json
import os
import re
import unittest
from collections import defaultdict

from excel.run import read_paths

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GOLD = os.path.join(HERE, "excel", "gold", "gold.json")
TARGET = 0.95


def _norm(v) -> str:
    return re.sub(r"\s+", " ", str(v if v is not None else "")).strip().casefold()


def actual_fields(visit, family="") -> dict:
    # The mock layout marks actions by fill colour; the golden set counts every note line there.
    steps = visit.notes if family == "mock" else [n for n in visit.notes if n.highlight == "green"]
    first_due = next((n.deadline for n in steps if re.fullmatch(r"\d{4}-\d{2}-\d{2}", n.deadline or "")), None)
    return {
        "date": visit.date_iso or "",
        "owner": visit.owner,
        "customer_name": visit.customer_raw,
        "city": visit.city,
        "contact": visit.contact_raw,
        "phone": re.sub(r"\D", "", visit.phone_raw or ""),
        "email": visit.email,
        "potential": visit.potential,
        "customer_type": visit.customer_type_raw,
        "purpose_start": re.sub(r"\s+", " ", visit.purpose or "").strip()[:20].rstrip(),
        "next_step_count": len(steps),
        "first_due": first_due,
    }


def same(field, want, got) -> bool:
    if field == "phone":
        return re.sub(r"\D", "", str(want or "")) == got
    if field == "next_step_count":
        return int(want or 0) == got
    if field == "first_due":
        return (want or None) == (got or None)
    if field == "purpose_start":
        return _norm(want)[:20].rstrip() == _norm(got)[:20].rstrip()
    return _norm(want) == _norm(got)


OVERRIDES = os.path.join(os.path.dirname(GOLD), "overrides.json")


def _overrides():
    """Decisions that replace a hand-checked value (excel/gold/overrides.json): each names the decision and why.
    They are counted and printed, never silent (tests/quality/decisions/visits.md V-11)."""
    if not os.path.exists(OVERRIDES):
        return {}
    return {(os.path.basename(o["file"]), o["sheet"].strip(), o["row"], o["field"]): o for o in json.load(open(OVERRIDES))}


def score():
    gold = json.load(open(GOLD))
    overrides, used = _overrides(), []
    paths = sorted({os.path.normpath(os.path.join(HERE, os.path.expanduser(g["file"]))) for g in gold})
    # One file at a time: the golden set scores the reader, not the cross-file dedupe.
    by_key = {}
    for path in paths:
        for v in read_paths([path]).parsed.visits:
            by_key[(v.workbook, v.sheet.strip(), v.excel_row_start)] = v
    tally = defaultdict(lambda: defaultdict(lambda: [0, 0]))
    misses = []
    for g in gold:
        v = by_key.get((os.path.basename(g["file"]), g["sheet"].strip(), g["row"]))
        got = actual_fields(v, g["family"]) if v else {}
        for f, want in g["expected"].items():
            o = overrides.get((os.path.basename(g["file"]), g["sheet"].strip(), g["row"], f))
            if o is not None:
                want = o["value"]
                used.append((o["decision"], os.path.basename(g["file"]), g["sheet"], g["row"], f))
            ok = bool(v) and same(f, want, got.get(f))
            t = tally[g["family"]][f]
            t[0] += ok
            t[1] += 1
            if not ok:
                misses.append((g["family"], os.path.basename(g["file"]), g["sheet"], g["row"], f,
                               "missing visit" if not v else "differs"))
    score.overrides_used = used
    return tally, misses


@unittest.skipUnless(os.path.exists(GOLD), "golden set not on this machine")
class GoldAccuracy(unittest.TestCase):
    def test_per_family_accuracy(self):
        tally, misses = score()
        report = {}
        for fam, fields in sorted(tally.items()):
            ok = sum(a for a, _ in fields.values())
            tot = sum(b for _, b in fields.values())
            report[fam] = ok / tot
            print(f"\n{fam}: {ok}/{tot} = {ok / tot:.1%}  " +
                  "  ".join(f"{f}={a}/{b}" for f, (a, b) in fields.items() if a < b))
        for m in misses:
            print("  miss", *m)
        for u in getattr(score, "overrides_used", []):
            print("  scored against an owner-decision override:", *u)
        for fam, acc in report.items():
            self.assertGreaterEqual(acc, TARGET, f"{fam} at {acc:.1%}")


if __name__ == "__main__":
    unittest.main()
