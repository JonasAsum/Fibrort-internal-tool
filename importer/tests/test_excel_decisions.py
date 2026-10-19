"""'Is this the same company?' answers: remembered, shared by file, and enforced in the run."""
import json

import pytest

from excel.companies import Master
from excel.decisions import DecisionStore
from tests.test_excel_companies import _master_records


@pytest.fixture
def master():
    return Master.from_records(_master_records())


def test_unanswered_near_name_is_a_review_not_a_match(master):
    m = master.match("Delta Trade Holdings Inc.")
    assert m.method == "review" and m.code is None and m.candidates


def test_answer_same_links_the_name_to_that_company(master):
    master.apply_decisions({"Delta Trade Holdings Inc.": "C002"})
    m = master.match("Delta Trade Holdings Inc.")
    assert (m.code, m.method, m.score) == ("C002", "exact", 1.0)


def test_answer_same_ignores_case_and_punctuation_of_the_name(master):
    master.apply_decisions({"Delta Trade Holdings Inc.": "C002"})
    assert master.match("DELTA  trade holdings, inc").code == "C002"


def test_answer_new_stops_the_question_and_makes_a_new_company(master):
    assert master.match("Kestrel Micro").method == "review"
    master.apply_decisions({"Kestrel Micro": None})
    m = master.match("Kestrel Micro")
    assert (m.code, m.method) == (None, "new")


def test_answer_new_also_covers_a_longer_spelling_of_the_same_name(master):
    master.apply_decisions({"Kestrel Micro": None})
    assert master.match("Kestrel Micro Co., Ltd.").method == "new"


def test_answer_new_does_not_swallow_a_different_company(master):
    master.apply_decisions({"Kestrel Micro": None})
    assert master.match("Delta Trade Holdings Inc.").method == "review"


def test_unknown_company_code_in_an_answer_is_rejected(master):
    with pytest.raises(ValueError, match="C999"):
        master.apply_decisions({"Delta Trade Holdings Inc.": "C999"})


def test_store_round_trip_and_atomic_file(tmp_path):
    path = tmp_path / "decisions.json"
    store = DecisionStore(path)
    store.set_same("Delta Trade Holdings Inc.", "C002")
    store.set_new("Kestrel Micro")
    again = DecisionStore(path)
    assert again.as_dict() == {"Delta Trade Holdings Inc.": "C002", "Kestrel Micro": None}
    assert not list(tmp_path.glob("*.tmp"))


def test_store_missing_file_is_empty(tmp_path):
    assert DecisionStore(tmp_path / "none.json").as_dict() == {}


def test_store_corrupt_file_fails_loudly_instead_of_dropping_answers(tmp_path):
    path = tmp_path / "decisions.json"
    path.write_text("{not json", encoding="utf-8")
    with pytest.raises(ValueError, match="decisions.json"):
        DecisionStore(path)


def test_store_later_answer_replaces_earlier(tmp_path):
    store = DecisionStore(tmp_path / "d.json")
    store.set_new("Kestrel Micro")
    store.set_same("Kestrel Micro", "C004")
    assert json.loads((tmp_path / "d.json").read_text(encoding="utf-8")) == {"Kestrel Micro": "C004"}


def test_unanswered_near_name_blocks_auto_accept_and_answered_one_does_not(master):
    from types import SimpleNamespace as NS
    from excel import run

    def kept(name):
        visit = NS(customer_raw=name, extra={}, workbook="w.xlsx", sheet="CW12",
                   excel_row_start=5, visit_key="k1")
        return [NS(layout="weekly", visit=visit)]

    items = run._company_items(kept("Delta Trade Holdings Inc."), master)
    assert [i.reason for i, _ in items] == ["company"]
    assert "company" not in run.AUTO_OK and "company_maybe_new" not in run.AUTO_OK

    master.apply_decisions({"Delta Trade Holdings Inc.": "C002"})
    assert run._company_items(kept("Delta Trade Holdings Inc."), master) == []
