"""A "same company?" question always carries the three closest companies by fuzzy score (best first), so the page can
offer them as one-click answers next to "New company"."""
from excel.companies import Master


def master(*rows):
    return Master.from_records([{"source": "master", "code": c, "cn": cn, "en": en, "short": ""} for c, cn, en in rows])


ROWS = [("C001", "", "Hengda Precision Machinery Co., Ltd."),
        ("C002", "", "Hengdi Precision Machinery Co., Ltd."),
        ("C003", "", "Hengda Precision Machine Tools Co., Ltd."),
        ("C004", "", "Hengda Precision Machinery Group"),
        ("C005", "", "Totally Different Logistics Ltd.")]


def test_a_question_carries_the_three_closest_companies_best_first():
    m = master(*ROWS).match("Henda Precision Machinery Co., Ltd")
    assert m.method == "review" and m.score < 0.9
    assert len(m.candidates) == 3
    scores = [c[2] for c in m.candidates]
    assert scores == sorted(scores, reverse=True)
    assert "C005" not in {c[0] for c in m.candidates}  # the unrelated company is never one of the three


def test_fewer_than_three_similar_companies_still_gives_the_closest_ones_that_exist():
    m = master(("C001", "", "Hengda Precision Machinery Co., Ltd."), ("C002", "", "Totally Different Logistics Ltd.")) \
        .match("Henda Precision Machinery Co., Ltd")
    assert m.method == "review" and 1 <= len(m.candidates) <= 3 and m.candidates[0][0] == "C001"


def test_an_exact_company_is_not_a_question():
    m = master(*ROWS).match("Hengda Precision Machinery Co., Ltd.")
    assert m.method == "exact" and m.score == 1.0
