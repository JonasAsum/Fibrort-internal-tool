"""L2: two independent AI labellers agree with the importer on every distinct customer cell, except the cases
listed in l2/pending.json, which wait for the owner's decision (the recommended default is already applied).

The labellers' answers are in l2/out_*_batch_*.json (made once; run `python -m tests.quality.l2_refresh` after a
reader change to re-read only the importer's side). A new cell with no labelled answer is reported, never ignored.
"""
import json
import os

import pytest

from tests.quality import l2_compare, l2_refresh, snapshot

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "l2")


@pytest.fixture(scope="module")
def result(sample_dir):
    if not os.path.exists(os.path.join(HERE, "customer_truth_hidden.json")):
        pytest.skip("labelled cells not on this machine")
    l2_refresh.refresh(snapshot.current(sample_dir))  # the importer's answers as of now
    return l2_compare.compare()


def test_every_labelled_cell_still_exists_and_has_answers(result):
    counts, _ = result
    assert counts["no-answer"] == 0 and counts["invalid-label"] == 0, counts


def test_importer_agrees_with_both_labellers_except_pending_cases(result):
    counts, rows = result
    pending = json.load(open(os.path.join(HERE, "pending.json"), encoding="utf-8"))
    failing = {l2_compare.text_id(r["text"]) for r in rows if r["verdict"] in ("FAIL", "review")}
    unexpected = failing - set(pending)
    assert not unexpected, "new disagreements:\n" + "\n".join(r["text"][:80] for r in rows if l2_compare.text_id(r["text"]) in unexpected)
    assert counts["agree"] >= 395, counts
