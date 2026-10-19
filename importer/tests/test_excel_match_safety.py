"""Wrong-merge guards found by the adversarial review. All names here are invented."""
import pytest

from excel.companies import Master


def rec(code, cn="", en="", short=""):
    return {"source": "master", "code": code, "cn": cn, "en": en, "short": short}


@pytest.fixture
def master():
    return Master.from_records([
        rec("CR0100001", en="Lixin Machinery Co., Ltd."),
        rec("CR0100002", en="Ningbo Haitian Group"),
        rec("CR0100003", en="Johnson Electric (Shenzhen) Co., Ltd.", cn="德昌电机（深圳）有限公司"),
        rec("PO12345", en="Other Corp"),
        rec("CR0100004", en="Shenzhen Shida Shengfeng Intelligent Equipment Co., Ltd."),
        rec("CR0100005", en="Kunshan Telali Precision Machinery Co., Ltd."),
        rec("CR0100006", en="Shenyang Keying Automation Equipment Co., Ltd."),
    ])


def test_one_letter_longer_short_name_is_not_auto_merged(master):
    m = master.match("Lixing Machinery")
    assert m.code is None and m.method == "review"
    assert m.candidates and m.candidates[0][0] == "CR0100001"


def test_exact_short_name_still_matches(master):
    assert master.match("Lixin Machinery Co").code == "CR0100001"


def test_different_city_blocks_auto_merge_of_a_group_branch(master):
    m = master.match("Tianjin Haitian Group")
    assert m.code is None and m.method == "review"
    assert m.candidates[0][0] == "CR0100002"


def test_same_city_or_no_city_still_matches(master):
    assert master.match("Ningbo Haitian Group").code == "CR0100002"
    assert master.match("Haitian Group").code == "CR0100002"


def test_trailing_city_in_master_name_does_not_hide_the_company(master):
    m = master.match("Johnson Electric Group")
    assert (m.code, m.method) == ("CR0100003", "contains")


def test_code_inside_unrelated_text_is_not_trusted(master):
    m = master.match("Follow-up, order PO12345 from Foo Co")
    assert m.method != "code" and m.code != "PO12345"


def test_code_with_matching_or_no_name_is_trusted(master):
    assert master.match("PO12345").code == "PO12345"
    assert master.match("Other Corp PO12345").code == "PO12345"


def test_one_edit_on_a_long_name_still_auto_matches(master):
    m = master.match("Shenzhen Sidashengfeng Intelligent Equipment Co., Ltd.")
    assert (m.code, m.method) == ("CR0100004", "fuzzy")


@pytest.mark.parametrize("name", [
    "Shenzhen Defalai Intelligent Technology Co., Ltd.",   # shares 2 letters in a row with 'telali'
    "Shenzhen S-KING Intelligent Equipment Co., Ltd",      # shares 3 letters with 'keying'
])
def test_chance_resemblance_is_a_new_company_not_a_question(master, name):
    m = master.match(name)
    assert (m.code, m.method) == (None, "new")
