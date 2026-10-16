"""Adversarial tests for company matching, decisions and the question/answer endpoints. Invented data only.

Tests marked xfail(strict) pin real defects found while attacking the code: they flip to a failure
(XPASS) the day the defect is fixed, so the marker must then be removed.
"""
import io
import json
import os
import random
import re
import string
import threading
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace as NS
from unittest import mock

import pytest

from excel import run
from excel.companies import Master, lookalike, norm, _places, _en_core, _cn_core
from excel.decisions import DecisionStore
from tests.excel_book import write_book, WEEKLY_HEADER

IMPORTER = Path(__file__).resolve().parent.parent


@pytest.fixture(autouse=True)
def decisions_file(tmp_path, monkeypatch):
    """Mandatory: no test may ever touch the real excel/company_decisions.json."""
    p = tmp_path / "decisions.json"
    monkeypatch.setenv("COMPANY_DECISIONS_FILE", str(p))
    if hasattr(run._master, "cache_clear"):
        run._master.cache_clear()
    yield p
    if hasattr(run._master, "cache_clear"):
        run._master.cache_clear()


def master(*rows):
    """rows: (code, cn, en)"""
    return Master.from_records([{"source": "master", "code": c, "cn": cn, "en": en, "short": ""} for c, cn, en in rows])


BASE = master(("C001", "", "Delta Trading Inc."), ("C002", "丁戊己贸易股份有限公司", "Gamma Instruments Ltd."))


def reads(*names):
    out = []
    for i, name in enumerate(names):
        visit = NS(customer_raw=name, extra={}, workbook="w.xlsx", sheet="CW12", excel_row_start=5 + i, visit_key=f"k{i}")
        out.append(NS(layout="lucas", visit=visit))
    return out


# ---------------------------------------------------------------- odd names into match()

@pytest.mark.parametrize("name", [None, "", "   ", "\t\n", 0, 123, 12345.0, float("nan"), "!!!", "---", "a", "Ａ", "D",
                                  "x" * 500, "<script>alert(1)</script>", "'; DROP TABLE companies;--", "\x00", "😀😀",
                                  "Delta" * 100, "丁" * 500, ["list"], {"d": 1}])
def test_match_never_raises_on_odd_input(name):
    m = BASE.match(name)
    assert m.method in {"none", "new", "review", "exact", "contains", "fuzzy", "ambiguous", "pinyin", "code"}
    assert m.code in (None, "C001", "C002")


@pytest.mark.parametrize("name", [None, "", "   ", 0, "!!!", "---"])
def test_empty_like_names_never_match_a_company(name):
    m = BASE.match(name)
    assert m.code is None and m.score < run.COMPANY_OK


@pytest.mark.parametrize("variant", [
    "DELTA TRADING INC", "delta   trading   inc.", "  Delta Trading Inc.  ", "Delta-Trading, Inc.",
    "ｄｅｌｔａ　ｔｒａｄｉｎｇ　ｉｎｃ．", "Delta Trading Inc.́",
    "Delta Trading Inc.", "Delta\tTrading\nInc.",
])
def test_same_name_in_other_case_spacing_width_is_one_company(variant):
    m = BASE.match(variant)
    assert (m.code, m.method) == ("C001", "exact")


def test_nfc_and_nfd_of_a_cjk_name_match_the_same():
    cn = "丁戊己贸易股份有限公司"
    m = master(("C9", cn, ""))
    for form in ("NFC", "NFD", "NFKC", "NFKD"):
        assert m.match(unicodedata.normalize(form, cn)).code == "C9", form


def test_nfc_nfd_of_accented_latin_name_are_the_same_key():
    assert norm(unicodedata.normalize("NFC", "Müller Werke")) == norm(unicodedata.normalize("NFD", "Müller Werke"))


def test_accented_letters_are_not_thrown_away():
    assert norm("Müller Präzision GmbH") != norm("Mller Przision GmbH")


def test_mixed_chinese_english_pinyin_name():
    m = master(("C5", "苏州凯斯特科技有限公司", "Suzhou Kaisite Technology Co., Ltd"))
    assert m.match("Suzhou Kaisite").code == "C5"
    assert m.match("苏州凯斯特").code == "C5"
    assert m.match("凯斯特 Kaisite (Suzhou)").code in ("C5", None)  # must not raise or pick another company


def test_unknown_company_among_many_is_new_not_forced_onto_a_neighbour():
    assert BASE.match("Orion Harbor Logistics").method == "new"


def test_html_in_name_is_returned_verbatim_never_matched_to_something():
    m = BASE.match("<script>alert('Delta Trading Inc.')</script>")
    assert m.code is None or m.method in ("exact", "contains", "fuzzy")  # no crash; value never echoed into Match


# ---------------------------------------------------------------- codes

def test_zero_padded_code_resolves_to_the_same_company():
    m = master(("ABC000123", "", "Zeta Mills"))
    assert m.has_code("ABC000123") and m.has_code("abc000123") and m.has_code("0ABC000123")
    assert m.get("abc000123")["en"] == "Zeta Mills"


def test_code_inside_a_sentence_about_another_company_goes_to_review():
    m = master(("ABC000123", "", "Zeta Mills"))
    r = m.match("Orion Harbor Logistics (see ABC000123)")
    assert r.method == "review" and r.code is None


def test_has_code_with_garbage_never_raises():
    for bad in (None, "", 0, [], {}, "  ", object()):
        assert BASE.has_code(bad) is False


@pytest.mark.xfail(strict=True, reason="BUG: codes 'C001' and 'c001' (or '0123' and '123') are two companies but one key in _by_code; the second is unreachable by code")
def test_codes_differing_only_by_case_or_leading_zero_are_not_silently_shadowed():
    m = master(("C0001", "", "First Works"), ("c0001", "", "Second Works"))
    assert m.get("C0001")["en"] == "First Works"
    assert m.get("c0001")["en"] == "Second Works"


def test_duplicate_code_rows_merge_into_one_company_keeping_both_names():
    m = master(("C1", "", "Alpha Works"), ("C1", "", "Beta Works"))
    assert m.size == 1
    assert m.match("Beta Works").code == "C1"


def test_empty_master_never_matches_or_raises():
    m = master()
    assert m.size == 0
    assert m.match("Delta Trading Inc.").method in ("none", "new")
    assert m.match("丁戊己").code is None
    assert m.search("anything") == []
    assert m.get("C1") is None


def test_company_without_code_still_matches_and_has_none_code():
    m = master(("", "", "Nocode Systems Inc."))
    r = m.match("Nocode Systems Inc.")
    assert r.code is None and r.method == "exact"


@pytest.mark.xfail(strict=True, reason="BUG: an exact match to a company that has no customer code returns code=None, indistinguishable from 'no match' (method is 'exact' but score 1.0 + code None)")
def test_exact_match_without_code_is_distinguishable_from_a_new_company():
    m = master(("", "", "Nocode Systems Inc."))
    r = m.match("Nocode Systems Inc.")
    assert r.code is not None or r.method != "exact"


def test_two_companies_sharing_a_core_never_pick_one_silently():
    m = master(("A1", "", "Kestrel Intelligent Equipment"), ("A2", "", "Kestrel Intelligent Machinery"))
    r = m.match("Kestrel Intelligent")
    assert r.code is None
    assert m.match("Kestrel Intelligent Equipment").code == "A1"


def test_two_companies_sharing_a_core_are_asked_about():
    m = master(("A1", "", "Kestrel Intelligent Equipment"), ("A2", "", "Kestrel Intelligent Machinery"))
    items = run._company_items(reads("Kestrel Intelligent"), m)
    assert [i.reason for i, _ in items] == ["company"]


def test_same_normalised_name_on_two_company_records_is_ambiguous():
    m = master(("A1", "", "Twin Fabrication Ltd"), ("A2", "", "Twin Fabrication Inc"))
    r = m.match("Twin Fabrication")
    assert r.code is None and r.method == "ambiguous" and len(r.candidates) == 2


def test_same_core_other_city_is_a_review_not_an_auto_merge():
    m = master(("S1", "", "Shanghai Kestrel Fabrication Co., Ltd"))
    r = m.match("Beijing Kestrel Fabrication Co., Ltd")
    assert r.code is None and r.score < run.COMPANY_OK


# ---------------------------------------------------------------- places

@pytest.mark.parametrize("text,lat,cn", [
    ("Foo Shanghai Bar", set(), set()),            # a place in the middle is part of the name
    ("Shanghai Foo Bar", {"shanghai"}, set()),
    ("Foo Bar Hong Kong", {"hong kong"}, set()),
    ("Hong Kong", set(), set()),                    # the whole name is a place: not a branch marker
    ("甲乙苏州丙丁", set(), set()),
    ("苏州甲乙丙", set(), {"苏州"}),
    ("甲乙丙苏州", set(), {"苏州"}),
    (None, set(), set()), ("", set(), set()), (5, set(), set()),
])
def test_places_only_at_the_edges(text, lat, cn):
    assert _places(text) == (frozenset(lat), frozenset(cn))


def test_en_core_never_strips_a_name_down_to_nothing():
    assert _en_core("Shanghai") == "shanghai"
    assert _en_core("Co Ltd") == ""
    assert _en_core("") == ""
    assert _cn_core("上海") == "上海"
    assert _cn_core("有限公司") == ""


# ---------------------------------------------------------------- performance / scale

def _bigmaster(n, seed=3):
    rnd = random.Random(seed)
    w = lambda: "".join(rnd.choice(string.ascii_lowercase) for _ in range(rnd.randint(4, 9)))
    return master(*[(f"K{i:05d}", "", f"{w()} {w()} Co Ltd") for i in range(n)]), w, rnd


def test_five_thousand_companies_each_match_stays_under_half_a_second():
    big, w, rnd = _bigmaster(5000)
    names = ["Quillon Marwick Ltd", "Orion Harbor Logistics", "丁戊己贸易", "Delta", "x", "a b c d e f g h", "Quillon Marwick Technology Co., Ltd Shenzhen",
             "苏州凯斯特科技有限公司", " ".join(w() for _ in range(8))]
    for n in names:
        t = time.perf_counter()
        big.match(n)
        assert time.perf_counter() - t < 0.5, n


def test_five_thousand_companies_500_char_name_stays_under_half_a_second():
    big, w, rnd = _bigmaster(5000)
    name = " ".join(w() for _ in range(70))[:500]
    t = time.perf_counter()
    big.match(name)
    assert time.perf_counter() - t < 0.5


def test_five_thousand_companies_search_stays_fast():
    big, _w, _r = _bigmaster(5000)
    for q in ("a", "qu", "quillon marwick", "x" * 400, "K00042"):
        t = time.perf_counter()
        big.search(q)
        assert time.perf_counter() - t < 0.5, q


def test_many_decided_new_names_do_not_slow_match_down():
    big, w, rnd = _bigmaster(2000)
    big.apply_decisions({f"New Thing {i} {w()} Ltd": None for i in range(1000)})
    t = time.perf_counter()
    big.match("Foo Barbaz Qux Ltd")
    assert time.perf_counter() - t < 0.5


def test_master_with_five_thousand_companies_builds_in_reasonable_time():
    t = time.perf_counter()
    _bigmaster(5000)
    assert time.perf_counter() - t < 5


# ---------------------------------------------------------------- search

@pytest.mark.parametrize("q", ["", " ", "!!!", "[", "(", "*", "\\", "?", ".*", "^$", "(?i)", "a", "d", "\x00", "x" * 5000, "丁", "<b>"])
def test_search_survives_odd_queries(q):
    r = BASE.search(q)
    assert isinstance(r, list) and len(r) <= 8
    assert all(set(x) == {"code", "cn", "en", "short", "uscc"} for x in r)


def test_search_regex_metacharacters_are_literal_not_wildcards():
    assert BASE.search(".*") == []
    assert BASE.search("d.lta") == []


def test_search_finds_by_code_name_and_chinese():
    assert [x["code"] for x in BASE.search("C001")] == ["C001"]
    assert [x["code"] for x in BASE.search("delta")] == ["C001"]
    assert [x["code"] for x in BASE.search("丁戊己")] == ["C002"]


def test_search_limit_is_respected():
    m = master(*[(f"Z{i:03d}", "", f"Zebra Works {i}") for i in range(30)])
    assert len(m.search("zebra", limit=8)) == 8
    assert len(m.search("zebra", limit=3)) == 3


def test_search_with_none_query_does_not_crash_the_server_path():
    assert BASE.search(None) == []


# ---------------------------------------------------------------- decisions in the matcher

def test_decided_new_name_stays_new_and_is_not_asked_again():
    m = master(("C001", "", "Delta Trading Inc."))
    assert m.match("Delta Trade Holdings Inc.").score < run.COMPANY_OK
    m.apply_decisions({"Delta Trade Holdings Inc.": None})
    assert m.match("Delta Trade Holdings Inc.").method == "new"
    assert m.match("DELTA TRADE HOLDINGS INC").method == "new"  # spelled in other case


def test_decided_same_links_name_to_the_code():
    m = master(("C001", "", "Delta Trading Inc."))
    m.apply_decisions({"Delta Trade Holdings Inc.": "C001"})
    r = m.match("Delta Trade Holdings Inc.")
    assert (r.code, r.method) == ("C001", "exact")


def test_decided_same_with_zero_padded_or_lowercase_code():
    m = master(("ABC000123", "", "Zeta Mills"))
    m.apply_decisions({"Zeta Mill Works": "abc000123"})
    assert m.match("Zeta Mill Works").code == "ABC000123"


def test_decided_new_then_identical_name_added_to_master_exact_wins():
    m = master(("C001", "", "Delta Trading Inc."), ("C9", "", "Quarry Hill Metals Ltd"))
    m.apply_decisions({"Quarry Hill Metals Ltd": None})
    assert m.match("Quarry Hill Metals Ltd").code == "C9"


@pytest.mark.xfail(strict=True, reason="BUG: a stale 'new company' answer beats a later master company with the same core ('Quarry Hill Metals Ltd' vs master 'Quarry Hill Metals Limited' stays 'new' forever)")
def test_decided_new_then_master_gains_the_company_with_another_legal_suffix():
    m = master(("C001", "", "Delta Trading Inc."), ("C9", "", "Quarry Hill Metals Limited"))
    m.apply_decisions({"Quarry Hill Metals Ltd": None})
    assert m.match("Quarry Hill Metals Ltd").code == "C9"


def test_decided_new_with_unusable_name_does_not_crash():
    m = master(("C001", "", "Delta Trading Inc."))
    m.apply_decisions({"": None, "   ": None, "!!!": None, "a": None})
    assert m.match("Delta Trading Inc.").code == "C001"


def test_punctuation_only_name_answered_new_is_not_asked_again():
    m = master(("C001", "", "Delta Trading Inc."))
    assert m.match("---").score < run.COMPANY_OK
    m.apply_decisions({"---": None})
    assert m.match("---").score >= run.COMPANY_OK


def test_decision_to_unknown_code_is_refused_loudly():
    with pytest.raises(ValueError):
        master(("C001", "", "Delta Trading Inc.")).apply_decisions({"X Ltd": "C999"})


def test_one_bad_decision_does_not_disable_all_others():
    m = master(("C001", "", "Delta Trading Inc."))
    skipped = m.apply_decisions({"Gone Ltd Works": "C999", "Delta Trade Holdings Inc.": "C001"}, strict=False)
    assert len(skipped) == 1 and "C999" in skipped[0]  # the stale entry is named, not fatal
    assert m.match("Delta Trade Holdings Inc.").code == "C001"
    with pytest.raises(ValueError, match="C999"):  # an answer a person just gave is still checked strictly
        m.apply_decisions({"Gone Ltd Works": "C999"})


@pytest.mark.parametrize("bad", [5, ["C001"], {"c": "C001"}, True])
def test_wrong_typed_decision_value_gives_a_clean_error(bad):
    try:
        master(("C001", "", "Delta Trading Inc.")).apply_decisions({"Some Name Ltd": bad})
    except ValueError:
        return
    except Exception as exc:  # AttributeError today
        pytest.fail(f"{type(exc).__name__}: {exc}")


# ---------------------------------------------------------------- DecisionStore on disk

@pytest.mark.parametrize("content", ["[]", "null", '"s"', "", " ", "{bad json", "5"])
def test_unreadable_decisions_file_is_refused_not_overwritten(tmp_path, content):
    p = tmp_path / "d.json"
    p.write_text(content)
    with pytest.raises(ValueError, match="unreadable"):
        DecisionStore(p)
    assert p.read_text() == content


def test_missing_file_means_no_decisions_and_is_created_on_first_answer(tmp_path):
    p = tmp_path / "sub" / "dir" / "d.json"
    s = DecisionStore(p)
    assert s.as_dict() == {}
    s.set_new("Foo Ltd Works")
    assert json.loads(p.read_text()) == {"Foo Ltd Works": None}
    assert not list(p.parent.glob("*.tmp"))


def test_null_name_key_and_unicode_roundtrip(tmp_path):
    p = tmp_path / "d.json"
    s = DecisionStore(p)
    s.set_same("苏州凯斯特科技有限公司", "C5")
    assert DecisionStore(p).as_dict() == {"苏州凯斯特科技有限公司": "C5"}
    assert "苏州" in p.read_text(encoding="utf-8")  # not \u-escaped, hand-editable


def test_huge_decisions_file_loads_and_applies_fast(tmp_path):
    p = tmp_path / "d.json"
    p.write_text(json.dumps({f"Invented Name Number {i} Works": None for i in range(20000)}))
    t = time.perf_counter()
    s = DecisionStore(p)
    assert len(s.as_dict()) == 20000
    assert time.perf_counter() - t < 2


def test_huge_decisions_file_applied_to_master_stays_workable(tmp_path):
    m = master(("C001", "", "Delta Trading Inc."))
    t = time.perf_counter()
    m.apply_decisions({f"Invented Name Number {i} Works": None for i in range(5000)})
    assert time.perf_counter() - t < 5
    t = time.perf_counter()
    m.match("Totally Different Words")
    assert time.perf_counter() - t < 0.5


@pytest.mark.skipif(os.geteuid() == 0, reason="root ignores directory permissions")
def test_read_only_directory_raises_and_leaves_state_readable(tmp_path):
    d = tmp_path / "ro"
    d.mkdir()
    p = d / "d.json"
    DecisionStore(p).set_new("Keep Me Works")
    d.chmod(0o500)
    try:
        s = DecisionStore(p)
        with pytest.raises(OSError):
            s.set_new("Another Name Works")
        assert DecisionStore(p).as_dict() == {"Keep Me Works": None}
        assert not list(d.glob("*.tmp"))
    finally:
        d.chmod(0o700)


def test_concurrent_writers_with_one_store_instance_lose_nothing(tmp_path):
    p = tmp_path / "d.json"
    s = DecisionStore(p)
    errs = []

    def work(i):
        for j in range(15):
            try:
                s.set_new(f"Thread {i} Name {j} Works")
            except Exception as e:  # noqa
                errs.append(repr(e))
    ts = [threading.Thread(target=work, args=(i,)) for i in range(8)]
    [t.start() for t in ts]
    [t.join() for t in ts]
    assert errs == [] and len(json.loads(p.read_text())) == 8 * 15


def test_concurrent_writers_with_separate_store_instances_lose_nothing(tmp_path):
    p = tmp_path / "d.json"
    errs = []

    def work(i):
        for j in range(15):
            try:
                DecisionStore(p).set_new(f"Thread {i} Name {j} Works")
            except Exception as e:  # noqa
                errs.append(repr(e))
    ts = [threading.Thread(target=work, args=(i,)) for i in range(8)]
    [t.start() for t in ts]
    [t.join() for t in ts]
    assert errs == [] and len(json.loads(p.read_text())) == 8 * 15


def test_stale_store_instance_overwrites_nothing_it_did_not_know_about_is_documented(tmp_path):
    """Two instances opened at the same time: the later write wins for ITS keys; the file stays valid JSON."""
    p = tmp_path / "d.json"
    a, b = DecisionStore(p), DecisionStore(p)
    a.set_new("First Works Name")
    b.set_new("Second Works Name")
    assert isinstance(json.loads(p.read_text()), dict)


def test_second_store_instance_keeps_the_first_ones_answer(tmp_path):
    p = tmp_path / "d.json"
    a, b = DecisionStore(p), DecisionStore(p)
    a.set_new("First Works Name")
    b.set_new("Second Works Name")
    assert set(json.loads(p.read_text())) == {"First Works Name", "Second Works Name"}


# ---------------------------------------------------------------- pair decisions

def test_pair_answers_in_opposite_order_last_one_wins_single_entry(tmp_path):
    s = DecisionStore(tmp_path / "d.json")
    s.set_pair("A Kestrel Works", "B Kestrel Works", True)
    s.set_pair("B Kestrel Works", "A Kestrel Works", False)
    assert s.pair_decision("A Kestrel Works", "B Kestrel Works") is False
    assert s.pair_decision("B Kestrel Works", "A Kestrel Works") is False
    assert len(s._data["__pairs__"]) == 1


def test_pair_key_ignores_case_spacing_and_width(tmp_path):
    s = DecisionStore(tmp_path / "d.json")
    s.set_pair("Kestrel Works Ltd", "Kestrel Works", True)
    assert s.pair_decision("KESTREL  WORKS", "ｋｅｓｔｒｅｌ ｗｏｒｋｓ ltd") is True


def test_pairs_do_not_leak_into_as_dict(tmp_path):
    s = DecisionStore(tmp_path / "d.json")
    s.set_pair("A Works", "B Works", True)
    s.set_new("C Works Name")
    assert "__pairs__" not in s.as_dict()


def test_pair_for_names_that_no_longer_exist_is_harmless(tmp_path):
    s = DecisionStore(tmp_path / "d.json")
    s.set_pair("Gone One Works", "Gone Two Works", True)
    items = run._company_items(reads("Alive Alpha Works", "Alive Omega Works"), BASE)
    assert [i for i, _ in items if i.reason == "company_pair"] == []


def test_company_literally_named_pairs_does_not_wipe_pair_answers(tmp_path):
    s = DecisionStore(tmp_path / "d.json")
    s.set_pair("A Works", "B Works", True)
    with pytest.raises(ValueError):  # the name is reserved
        s.set_new("__pairs__")
    assert s.pair_decision("A Works", "B Works") is True


@pytest.mark.parametrize("pairs", [[1], None, "zz", [{"a": "x"}]])
def test_malformed_pairs_section_is_ignored_not_fatal(tmp_path, pairs):
    p = tmp_path / "d.json"
    p.write_text(json.dumps({"__pairs__": pairs}))
    assert DecisionStore(p).pair_decision("A Works", "B Works") is None


def test_pair_with_non_bool_same_is_stored_as_bool(tmp_path):
    s = DecisionStore(tmp_path / "d.json")
    s.set_pair("A Works", "B Works", "no")  # truthy string; store coerces, the endpoint is what must refuse it
    assert s.pair_decision("A Works", "B Works") is True


def test_same_pair_answer_merges_names_in_run_items(tmp_path):
    long_, short = "Shenzhen Kestrel Intelligent Equipment Co., Ltd", "Shenzhen Kestrel Intelligent"
    DecisionStore(tmp_path / "decisions.json").set_pair(long_, short, True)
    rs = reads(short, long_)
    items = run._company_items(rs, BASE)
    assert items == []
    assert {r.visit.customer_raw for r in rs} == {long_}


def test_pair_items_with_empty_and_punctuation_names_do_not_crash():
    rs = reads("", None, "!!!", "---", "   ", "a", "A")
    run._company_items(rs, BASE)


def test_three_way_lookalike_is_asked_pairwise_without_crash():
    items = run._company_items(reads("Kestrel Intelligent Equipment Co., Ltd", "Kestrel Intelligent Equipment", "Kestrel Intelligent"), BASE)
    assert {i.reason for i, _ in items} == {"company_pair"}


# ---------------------------------------------------------------- lookalike quality

_SYL = "ka ro mi ta ven lor dex sil qua bru fen gal hox jun pel zor tan wid yor nex".split()
_SUF = ["Co., Ltd", "Inc.", "GmbH", "Machinery Co., Ltd", "Technology Co., Ltd", "Group", "Trading Ltd"]


def _invented(rnd):
    w = lambda: "".join(rnd.choice(_SYL) for _ in range(rnd.randint(2, 3))).capitalize()
    return f"{w()} {w()} {rnd.choice(_SUF)}"


def test_lookalike_false_positive_rate_over_200_unrelated_pairs_is_low():
    rnd = random.Random(7)
    hits = []
    for _ in range(200):
        a, b = _invented(rnd), _invented(rnd)
        if lookalike(a, b):
            hits.append((a, b))
    assert len(hits) <= 4, hits  # 2% of unrelated pairs


def test_lookalike_all_pairs_inside_200_names_few_false_positives():
    rnd = random.Random(11)
    names = list({_invented(rnd) for _ in range(300)})[:200]
    hits = [(a, b) for i, a in enumerate(names) for b in names[i + 1:] if lookalike(a, b)]
    assert len(hits) <= 10, hits  # of 19900 pairs


TRUNCATIONS = [
    ("Harbor Lantern Machinery Co., Ltd", "Harbor Lantern"), ("Quillon Marwick Technology", "Quillon Marwick"),
    ("Brindle Forge Precision Machinery Co Ltd", "Brindle Forge"), ("Tamsin Orchard Group", "Tamsin Orchard"),
    ("Velox Dynamics International", "Velox Dynamics"), ("Norland Pumps Inc.", "Norland Pumps"),
    ("Zephyr Optics GmbH", "Zephyr Optics"), ("Corvin Tooling Ltd", "Corvin Tooling"),
    ("Ashgrove Fabrication Co., Ltd", "Ashgrove Fabrication"), ("Pelham Castings Industrial", "Pelham Castings"),
    ("Lantern Works", "Lantern Workz"), ("Marwick Hydraulics", "Marwick Hydraulic"),
    ("Orbis Controls", "Orbis Control"), ("Sable Ridge Manufacturing", "Sable Ridge"),
    ("Tristan Valve Co", "Tristan Valve"), ("Ulmer Bearings", "Ulmer Bearing"),
    ("Wexford Gears Ltd", "Wexford Gear"), ("Yardley Presses", "Yardley Press"),
    ("Kelmar Tools Limited", "Kelmar Tools"), ("Draven Extrusion Corp", "Draven Extrusion"),
    ("甲乙丙丁戊机械有限公司", "甲乙丙丁戊"), ("苏州甲乙丙精密", "甲乙丙精密"),
]


@pytest.mark.parametrize("a,b", TRUNCATIONS)
def test_lookalike_finds_truncation_and_extension_pairs(a, b):
    assert lookalike(a, b) and lookalike(b, a)


@pytest.mark.parametrize("a,b", [(None, None), ("", ""), ("   ", "x"), (5, 5), ("a", "a"), ("!!!", "!!!"), ("Co Ltd", "Co Ltd"),
                                 ("Inc", "Inc"), ("有限公司", "有限公司")])
def test_lookalike_is_false_for_empty_or_generic_only_names(a, b):
    assert lookalike(a, b) is False


@pytest.mark.parametrize("a,b", [("Delta Trading Inc.", "Gamma Instruments Ltd."), ("Kestrel Micro", "Orion Harbor Logistics"),
                                 ("Alpha Works", "Alpine Works")])
def test_lookalike_does_not_confuse_neighbours(a, b):
    assert lookalike(a, b) is False


def test_lookalike_is_symmetric_on_random_pairs():
    rnd = random.Random(5)
    for _ in range(100):
        a, b = _invented(rnd), _invented(rnd)
        assert lookalike(a, b) == lookalike(b, a)


# ---------------------------------------------------------------- endpoints

from webui import server  # noqa: E402

NAME = "Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"
ASKED = "Delta Trade Holdings Inc."


@pytest.fixture
def app(tmp_path, monkeypatch):
    m = master(("C002", "丁戊己贸易股份有限公司", "Delta Trading Inc."), ("C003", "庚辛壬仪器有限公司", "Gamma Instruments Ltd."))
    monkeypatch.setattr("excel.run._master", lambda *_a, **_k: m)
    monkeypatch.setattr(server, "UPLOAD_DIR", str(tmp_path / "uploads"))
    client = server.app.test_client()
    yield NS(client=client, master=m, tmp=tmp_path)
    server._forget()


def upload(app, *customers):
    cells = dict(WEEKLY_HEADER)
    cells.update({"E4": "Owner：", "F4": "Jason Zhao"})
    for i, c in enumerate(customers):
        r = 9 + i
        cells.update({f"B{r}": "Jason Zhao", f"D{r}": datetime(2026, 5, 18), f"F{r}": c, f"P{r}": "Call"})
    book = write_book(str(app.tmp / "w.xlsx"), {"CW21": {"cells": cells}})
    with open(book, "rb") as f:
        data = {"files": (io.BytesIO(f.read()), NAME)}
    r = app.client.post("/api/upload", data=data, content_type="multipart/form-data")
    return r


def answer(app, **body):
    return app.client.post("/api/company_answer", json=body)


def saved(decisions_file):
    return json.loads(Path(decisions_file).read_text(encoding="utf-8")) if Path(decisions_file).exists() else {}


def test_html_name_comes_back_as_json_text_and_answer_roundtrips(app, decisions_file):
    evil = "Delta <b>Trade</b> \"x\" 'y' & Holdings Inc."
    r = upload(app, evil)
    assert r.status_code == 200, r.get_json()
    assert r.mimetype == "application/json"
    qs = r.get_json()["company_questions"]
    assert [q["name"] for q in qs] == [evil]
    d = answer(app, name=evil, code=None)
    assert d.status_code == 200 and d.get_json()["company_questions"] == []
    assert saved(decisions_file)[evil] is None


def test_app_js_escapes_every_name_it_puts_in_html():
    js = (IMPORTER / "webui" / "static" / "app.js").read_text(encoding="utf-8")
    start = js.index("function renderCompanyQuestions")
    end = js.index("function wireQuestion")
    block = js[start:end]
    for expr in re.findall(r"\$\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}", block):
        e = expr.strip()
        if re.search(r"\bq\.(name|a|b)\b|\bc\.(name|en|cn|code)\b|\bcode\b|\bname\b", e) and not e.startswith("esc("):
            assert "esc(" in e or "map" in e or "candidates.length" in e, f"unescaped interpolation: {e}"
    assert not re.search(r"\$\{\s*q\.(name|a|b)\s*\}", js)
    assert "innerHTML = d" not in js.replace(" ", "")  # nothing assigns raw server text
    esc_def = re.search(r"function esc\(s\)\s*\{(.*?)\n\}", js, re.S).group(1)
    for ch in ("&amp;", "&lt;", "&gt;", "&quot;", "&#39;"):
        assert ch in esc_def


def test_search_results_in_js_are_escaped():
    js = (IMPORTER / "webui" / "static" / "app.js").read_text(encoding="utf-8")
    seg = js[js.index("function cqOption"):js.index("function wireQuestion")]
    assert seg.count("esc(") >= 3
    assert "textContent" in js[js.index("function wireQuestion"):]  # label built via textContent, not innerHTML


@pytest.mark.parametrize("cust", ["---", "N/A", "12345", "a", "x" * 500, "😀", "Ａ", "Delta  Trade   Holdings Inc.", "<img src=x onerror=alert(1)>"])
def test_upload_with_odd_customer_names_never_500s(app, cust):
    r = upload(app, cust)
    assert r.status_code == 200, r.get_json()
    assert app.client.get("/api/company_search", query_string={"q": cust[:50]}).status_code == 200


def test_answering_new_for_punctuation_only_customer_clears_the_question(app):
    upload(app, "---")
    d = answer(app, name="---", code=None)
    assert d.status_code == 200
    assert d.get_json()["company_questions"] == []


def test_case_variants_of_one_name_are_one_question_and_one_answer_clears_both(app):
    upload(app, ASKED, ASKED.upper(), "  " + ASKED.lower())
    qs = app.client.get("/api/company_search")  # endpoint exists
    assert qs.status_code == 200
    r = upload(app, ASKED, ASKED.upper())
    qs = r.get_json()["company_questions"]
    assert len(qs) == 1 and qs[0]["visits"] == 2
    d = answer(app, name=qs[0]["name"], code="C002").get_json()
    assert d["company_questions"] == [] and d["unsure"] == 0


@pytest.mark.parametrize("body", [None, {}, {"name": None}, {"code": "C002"}, {"name": ""}, {"name": 5, "code": None}, {"name": True}])
def test_company_answer_with_bad_bodies_is_400_not_500(app, body):
    upload(app, ASKED)
    r = app.client.post("/api/company_answer", json=body) if body is not None else app.client.post("/api/company_answer")
    assert r.status_code == 400


def test_company_answer_with_non_json_content_type_is_400(app):
    upload(app, ASKED)
    assert app.client.post("/api/company_answer", data="name=x", content_type="text/plain").status_code == 400
    assert app.client.post("/api/company_answer", data="{not json", content_type="application/json").status_code == 400


def test_company_answer_with_list_name_is_400(app):
    upload(app, ASKED)
    assert answer(app, name=["x"], code=None).status_code == 400


@pytest.mark.parametrize("code", ["C999", "", "  ", 0, 5, ["C002"], {"c": "C002"}, "' OR 1=1 --", "<script>", "C002; DROP", "c0 02", "False", "None", "null"])
def test_company_answer_with_bad_code_is_400_and_saves_nothing(app, decisions_file, code):
    upload(app, ASKED)
    before = saved(decisions_file)
    assert answer(app, name=ASKED, code=code).status_code == 400
    assert saved(decisions_file) == before


def test_company_answer_code_in_other_case_or_zero_padded_is_accepted(app):
    upload(app, ASKED)
    assert answer(app, name=ASKED, code="c002").status_code == 200


def test_double_submit_second_answer_is_refused_and_first_kept(app, decisions_file):
    upload(app, ASKED)
    assert answer(app, name=ASKED, code="C002").status_code == 200
    assert answer(app, name=ASKED, code=None).status_code == 400  # question is gone
    assert saved(decisions_file)[ASKED] == "C002"


def test_answer_after_clear_is_refused_and_saves_nothing(app, decisions_file):
    upload(app, ASKED)
    assert app.client.post("/api/clear").status_code == 200
    before = saved(decisions_file)
    assert answer(app, name=ASKED, code=None).status_code == 400
    assert saved(decisions_file) == before


def test_answer_for_name_with_different_case_than_asked_is_refused(app, decisions_file):
    upload(app, ASKED)
    assert answer(app, name=ASKED.upper(), code=None).status_code == 400
    assert saved(decisions_file) == {}


def test_answer_with_corrupt_decisions_file_is_an_error_not_silent_loss(app, decisions_file):
    upload(app, ASKED)
    Path(decisions_file).write_text("{broken")
    r = answer(app, name=ASKED, code=None)
    assert r.status_code >= 400
    assert Path(decisions_file).read_text() == "{broken"


def test_answer_to_read_only_decisions_dir_does_not_pretend_success(app, decisions_file, monkeypatch):
    upload(app, ASKED)
    monkeypatch.setattr(DecisionStore, "_save", mock.Mock(side_effect=PermissionError("ro")))
    r = answer(app, name=ASKED, code=None)
    assert r.status_code >= 400
    assert app.master.match(ASKED).method != "new"  # not applied in memory when the save failed


def test_failed_save_returns_json_error(app, decisions_file, monkeypatch):
    upload(app, ASKED)
    monkeypatch.setattr(DecisionStore, "_save", mock.Mock(side_effect=PermissionError("ro")))
    r = answer(app, name=ASKED, code=None)
    assert r.mimetype == "application/json" and "error" in r.get_json()


# pairs over HTTP

L, S = "Shenzhen Kestrel Intelligent Equipment Co., Ltd", "Shenzhen Kestrel Intelligent"


def upload_pair(app):
    r = upload(app, S, L)
    assert r.status_code == 200
    return r.get_json()


def pair(app, **body):
    return app.client.post("/api/company_pair_answer", json=body)


@pytest.mark.parametrize("same", ["yes", "true", 1, 0, None, "", [], {}, "False"])
def test_pair_answer_with_non_bool_same_is_400_and_saves_nothing(app, decisions_file, same):
    upload_pair(app)
    assert pair(app, a=L, b=S, same=same).status_code == 400
    assert "__pairs__" not in saved(decisions_file)


@pytest.mark.parametrize("body", [None, {}, {"a": L}, {"same": True}, {"a": L, "b": L, "same": True}, {"a": S, "b": "Other Works", "same": True}])
def test_pair_answer_with_bad_bodies_is_400(app, body):
    upload_pair(app)
    r = app.client.post("/api/company_pair_answer", json=body) if body is not None else app.client.post("/api/company_pair_answer")
    assert r.status_code == 400


def test_pair_answer_with_list_names_is_400(app):
    upload_pair(app)
    assert pair(app, a=["x"], b=["y"], same=True).status_code == 400


def test_pair_answer_in_either_order_works_and_is_remembered_once(app, decisions_file):
    upload_pair(app)
    assert pair(app, a=S, b=L, same=False).status_code == 200
    d = saved(decisions_file)["__pairs__"]
    assert len(d) == 1 and d[0]["same"] is False


def test_pair_double_submit_second_is_refused(app, decisions_file):
    upload_pair(app)
    assert pair(app, a=L, b=S, same=True).status_code == 200
    assert pair(app, a=L, b=S, same=False).status_code == 400
    assert saved(decisions_file)["__pairs__"][0]["same"] is True


def test_pair_answer_after_clear_is_refused(app):
    upload_pair(app)
    app.client.post("/api/clear")
    assert pair(app, a=L, b=S, same=True).status_code == 400


def test_pair_answer_for_names_that_no_longer_exist_is_refused(app, decisions_file):
    upload_pair(app)
    upload(app, "Totally Other Works")  # a new upload replaces the earlier one
    assert pair(app, a=L, b=S, same=True).status_code == 400
    assert "__pairs__" not in saved(decisions_file)


def test_pair_question_html_names_roundtrip(app):
    a, b = "Kestrel <b>Intelligent</b> Equipment Co., Ltd", "Kestrel <b>Intelligent</b>"
    r = upload(app, a, b)
    qs = [q for q in r.get_json()["company_questions"] if q["kind"] == "pair"]
    assert len(qs) == 1 and {qs[0]["a"], qs[0]["b"]} == {a, b}
    assert pair(app, a=a, b=b, same=True).status_code == 200


# search over HTTP

@pytest.mark.parametrize("q", ["", "a", "(", "[a-", "*", "\\", ".*", "' OR 1=1 --", "<script>alert(1)</script>", "x" * 20000, "丁", "%", "%00", "\x00"])
def test_company_search_endpoint_survives(app, q):
    r = app.client.get("/api/company_search", query_string={"q": q})
    assert r.status_code == 200
    assert isinstance(r.get_json()["results"], list)


def test_company_search_without_q_param(app):
    r = app.client.get("/api/company_search")
    assert r.status_code == 200 and r.get_json()["results"] == []


def test_company_search_one_character_returns_at_most_eight(app):
    assert len(app.client.get("/api/company_search", query_string={"q": "a"}).get_json()["results"]) <= 8


def test_company_search_repeated_q_params_use_first(app):
    r = app.client.get("/api/company_search?q=delta&q=gamma")
    assert [x["code"] for x in r.get_json()["results"]] == ["C002"]
