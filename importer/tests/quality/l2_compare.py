"""L2: two independent AI labellers against the importer, on every distinct customer-cell text.

The labellers (two different prompt styles) saw only the raw cell text, never the importer's answer. Each field
they return must be copied from the cell: a field that is not literally in the cell is rejected as an invented
value (a labeller error, never counted for or against the importer). The importer's answer is then compared
with both labellers, field by field (letters and digits only, so brackets and spacing do not matter).

Verdicts per cell text:
  agree        both labellers and the importer give the same name text, address and code
  one-agrees   the importer matches exactly one labeller: the labellers disagree, a person may look
  FAIL         both labellers give the same answer and the importer differs: a bug until proven otherwise
  review       all three differ

Decisions the owner took on a cell are in l2/decisions.json ({sha1 of the text: "accept reader" | note}).

    .venv/bin/python -m tests.quality.l2_compare
"""
import hashlib
import json
import os
import re
import sys
import unicodedata

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "l2")
FIELDS = ("name", "name_en", "address", "code")
LABELLERS = ("A", "B")


def key(text) -> str:
    return re.sub(r"[\W_]+", "", unicodedata.normalize("NFKC", str(text or "")).casefold())


def text_id(text) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:12]


def _valid(item, raw) -> list:
    """Problems with one labeller answer; [] when every field is copied from the cell and the cell is covered."""
    problems = []
    flat = key(raw)
    covered = ""
    for f in FIELDS + ("other",):
        v = item.get(f) or ""
        if v and key(v) not in flat:
            problems.append(f"{f} not in the cell: {v[:40]!r}")
        covered += key(v)
    label = re.sub(r"(?i)addr(?:ess)?", "", raw)  # the 'Addr:' label may be dropped
    if sorted(covered) != sorted(key(label)) and sorted(covered) != sorted(flat):
        problems.append("fields do not cover the cell exactly")
    return problems


def load():
    truth = json.load(open(os.path.join(HERE, "customer_truth_hidden.json"), encoding="utf-8"))
    out = {}
    for who in LABELLERS:
        answers = {}
        for n in range(1, 7):
            p = os.path.join(HERE, f"out_{who}_batch_{n}.json")
            if not os.path.exists(p):
                continue
            batch = json.load(open(os.path.join(HERE, f"customer_batch_{n}.json"), encoding="utf-8"))
            got = json.load(open(p, encoding="utf-8"))
            assert len(got) == len(batch), f"labeller {who} batch {n}: {len(got)} answers for {len(batch)} cells"
            for item, asked in zip(got, batch):  # same order as the input; an id the labeller wrote must agree
                assert item.get("id", asked["id"]) == asked["id"], f"labeller {who} batch {n}: answers out of order"
                answers[asked["id"]] = item
        out[who] = answers
    return truth, out


def compare():
    truth, out = load()
    decisions = {}
    dp = os.path.join(HERE, "decisions.json")
    if os.path.exists(dp):
        decisions = json.load(open(dp, encoding="utf-8"))
    rows, counts = [], {"agree": 0, "one-agrees": 0, "FAIL": 0, "review": 0, "no-answer": 0, "invalid-label": 0}
    for cid, t in truth.items():
        raw, reader = t["text"], t["reader"]
        a, b = out["A"].get(cid), out["B"].get(cid)
        if a is None or b is None:
            counts["no-answer"] += 1
            rows.append({"id": cid, "verdict": "no-answer", "text": raw})
            continue
        bad = _valid(a, raw) + _valid(b, raw)
        if bad:
            counts["invalid-label"] += 1
            rows.append({"id": cid, "verdict": "invalid-label", "text": raw, "problems": bad})
            continue
        # POLICY (hand-checked gold set): a bilingual name stays one name, so name and name_en are compared as
        # one text, in any order. Address and code are compared on their own.
        sig = lambda d: ("".join(sorted(key(d.get("name")) + key(d.get("name_en")))), key(d.get("address")), key(d.get("code")))  # noqa: E731
        sr, sa, sb = sig(reader), sig(a), sig(b)
        if sr == sa == sb:
            verdict = "agree"
        elif sa == sb:
            verdict = "FAIL"
        elif sr in (sa, sb):
            verdict = "one-agrees"
        else:
            verdict = "review"
        if verdict in ("FAIL", "review", "one-agrees") and text_id(raw) in decisions:
            verdict = "agree"  # the owner decided this cell
        counts[verdict] += 1
        if verdict != "agree":
            rows.append({"id": cid, "verdict": verdict, "text": raw, "reader": {f: reader.get(f) for f in FIELDS},
                         "A": {f: a.get(f) for f in FIELDS}, "B": {f: b.get(f) for f in FIELDS}, "places": t["places"]})
    return counts, rows


if __name__ == "__main__":
    counts, rows = compare()
    print(json.dumps(counts))
    json.dump({"counts": counts, "rows": rows}, open(os.path.join(HERE, "report.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    for r in rows[:int(sys.argv[1]) if len(sys.argv) > 1 else 12]:
        print(f"\n[{r['verdict']}] {r['text']!r}")
        for who in ("reader", "A", "B"):
            if who in r:
                print(f"   {who:6}", {k: v for k, v in r[who].items() if v})
        if r.get("problems"):
            print("   ", r["problems"][:3])
    sys.exit(1 if counts["FAIL"] else 0)
