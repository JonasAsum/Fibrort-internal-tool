"""Quality checks for the visit readers (folders 01-03), run on every record of every workbook.

Each rule below is a whole-population check. A rule listed in KNOWN_BUGS is a bug we know about and
have not fixed yet: it is marked strict-xfail, so it keeps failing loudly in the report, and the day it
is fixed the marker must be removed (the test then passes, and a surprise pass fails the run).
"""
import collections
import os

import pytest

from tests.quality import run_visits, snapshot

KNOWN_BUGS = {  # empty: every known visit bug is fixed; a new known bug goes here as strict-xfail with its reason
}
ALL_RULES = sorted(set(KNOWN_BUGS) | {"not-preserved", "customer-code-in-name", "english-name-in-address", "stray-bracket-in-address", "invented-value", "email-field-holds-non-email", "customer-not-preserved",
                                      "address-text-in-name", "customer-code-not-extracted", "phone-field-not-a-phone"})


@pytest.fixture(scope="session")
def violations(sample_dir):
    if not os.path.isdir(sample_dir):
        pytest.skip("sample folder not found")
    return run_visits.run(snapshot.current(sample_dir))


def _param(rule):
    if rule in KNOWN_BUGS:
        return pytest.param(rule, marks=pytest.mark.xfail(strict=True, reason=KNOWN_BUGS[rule]))
    return rule


@pytest.mark.parametrize("rule", [_param(r) for r in ALL_RULES])
def test_rule_has_no_violations(rule, violations):
    bad = [v for v in violations if v["rule"] == rule]
    assert not bad, f"{len(bad)} record(s) break '{rule}':\n" + "\n".join(
        f"  {v['file']} {v['sheet'].strip()} row {v['row']}: source {v['source'][:60]!r} -> {str(v['got'])[:60]!r}"
        for v in bad[:5])


def test_every_visit_record_is_checked(violations, sample_dir):
    n = sum(1 for r in snapshot.current(sample_dir).values() if r.get("type") == "visit")
    assert n >= 1200  # the 31 visit workbooks hold 1,267 visits; a drop means a file stopped being read


def test_snapshot_is_unchanged(sample_dir):
    """The output equals the accepted baseline. After a deliberate reader change: read the diff
    (`python -m tests.quality.snapshot diff`), then accept it (`... write`)."""
    if not os.path.exists(snapshot.GOLDEN):
        pytest.skip("no accepted snapshot on this machine (real customer data is not in git)")
    d = snapshot.diff(snapshot.load(), snapshot.current(sample_dir))
    assert not (d["added"] or d["removed"] or d["changed"]), snapshot.report(d)


def test_the_checks_catch_the_old_bugs(sample_dir, monkeypatch):
    """Test the test: put the old behaviour back (no two-line split, no code split, office text as a phone)
    and the invariants must fail on the Sean Wang and Lucas files. If they do not, the checks are blind."""
    from excel import readers, run
    monkeypatch.setattr(run, "_FILE_CACHE", {})  # excel.run keeps parsed files: without this the fixed result is returned
    monkeypatch.setattr(readers, "split_two_lines", lambda raw: None)
    monkeypatch.setattr(readers, "split_customer_code", lambda raw: ("", raw))
    monkeypatch.setattr(readers, "looks_like_phone", lambda text: True)
    files = [p for p in snapshot.visit_files(sample_dir) if "Sean Wang" in p or "CW04-Lucas" in p]
    assert len(files) == 2
    recs = {}
    for p in files:
        recs.update(snapshot.records_of(p))
    rules = collections.Counter(v["rule"] for v in run_visits.run(recs))
    for expected in ("stray-bracket-in-address", "english-name-in-address", "customer-code-in-name", "phone-field-not-a-phone"):
        assert rules[expected] > 0, f"the checks did not notice the old bug '{expected}': {dict(rules)}"
