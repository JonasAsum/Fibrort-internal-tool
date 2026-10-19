"""Two names in one upload that may be one company: ask, remember, then merge or keep apart. Invented names."""
import json
from types import SimpleNamespace as NS

import pytest

from excel import run
from excel.companies import Master, lookalike
from excel.decisions import DecisionStore

LONG = "Shenzhen Kestrel Intelligent Equipment Co., Ltd"
SHORT = "Shenzhen Kestrel Intelligent"


@pytest.fixture
def master():
    return Master.from_records([
        {"source": "master", "code": "C002", "cn": "丁戊己贸易股份有限公司", "en": "Delta Trading Inc.", "short": ""},
    ])


@pytest.fixture(autouse=True)
def decisions_file(tmp_path, monkeypatch):
    monkeypatch.setenv("COMPANY_DECISIONS_FILE", str(tmp_path / "decisions.json"))
    return tmp_path / "decisions.json"


def reads(*names):
    out = []
    for i, name in enumerate(names):
        visit = NS(customer_raw=name, extra={}, workbook="w.xlsx", sheet="CW12", excel_row_start=5 + i, visit_key=f"k{i}")
        out.append(NS(layout="lucas", visit=visit))
    return out


@pytest.mark.parametrize("a,b,expected", [
    (LONG, SHORT, True),                                  # one is the other with extra words
    ("Lixin Machinery", "Lixing Machinery", True),        # one letter apart
    ("上海甲乙丙科技有限公司", "甲乙丙", True),
    ("Kestrel Micro", "Orion Harbor Logistics", False),
    ("Delta Trading Inc.", "Gamma Instruments Ltd.", False),
])
def test_lookalike(a, b, expected):
    assert lookalike(a, b) is expected


def test_two_lookalike_new_names_are_asked_about_as_a_pair(master):
    items = run._company_items(reads(LONG, SHORT), master)
    assert sorted(i.reason for i, _ in items) == ["company_pair", "company_pair"]
    assert {i.value for i, _ in items} == {LONG, SHORT}
    other = {i.value: i.candidates[0][0] for i, _ in items}
    assert other[LONG] == SHORT and other[SHORT] == LONG
    assert "company_pair" not in run.AUTO_OK


def test_unrelated_new_names_are_not_asked_about(master):
    assert run._company_items(reads("Kestrel Micro", "Orion Harbor Logistics"), master) == []


def test_same_company_answer_merges_the_names_and_stops_the_question(master, decisions_file):
    DecisionStore(decisions_file).set_pair(LONG, SHORT, same=True)
    rs = reads(LONG, SHORT)
    assert run._company_items(rs, master) == []
    assert rs[0].visit.customer_raw == rs[1].visit.customer_raw == LONG  # the fuller spelling wins
    assert rs[1].visit.extra["customer_as_written"] == SHORT


def test_different_company_answer_keeps_them_apart_and_stops_the_question(master, decisions_file):
    DecisionStore(decisions_file).set_pair(LONG, SHORT, same=False)
    rs = reads(LONG, SHORT)
    assert run._company_items(rs, master) == []
    assert [r.visit.customer_raw for r in rs] == [LONG, SHORT]


def test_pair_answers_are_remembered_in_either_order_and_spelling(decisions_file):
    store = DecisionStore(decisions_file)
    store.set_pair(LONG, SHORT, same=True)
    again = DecisionStore(decisions_file)
    assert again.pair_decision(SHORT.upper(), LONG) is True
    assert again.pair_decision(LONG, "Something Else Ltd") is None


def test_pair_answers_do_not_leak_into_company_answers(decisions_file):
    store = DecisionStore(decisions_file)
    store.set_pair(LONG, SHORT, same=True)
    store.set_new("Kestrel Micro")
    assert store.as_dict() == {"Kestrel Micro": None}
    assert "__pairs__" in json.loads(decisions_file.read_text(encoding="utf-8"))
