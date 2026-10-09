from pathlib import Path

import openpyxl
import pytest

from excel.companies import Master, norm, parse_rows


def _master_records():
    return [
        {"source": "master", "code": "C001", "cn": "甲乙丙机电有限公司", "en": "Alpha Beta Gamma Co., Ltd.", "short": ""},
        {"source": "master", "code": "C002", "cn": "丁戊己贸易股份有限公司", "en": "Delta Trading Inc.", "short": ""},
        {"source": "erp", "code": "C002", "cn": "丁戊己贸易股份有限公司", "en": "", "short": "丁戊己"},
        {"source": "erp", "code": "C003", "cn": "庚辛壬仪器有限公司", "en": "", "short": "庚辛壬"},
        {"source": "erp", "code": "C004", "cn": "", "en": "Kestrel Microsystems Corporation", "short": ""},
        {"source": "erp", "code": "C005", "cn": "子丑寅卯科技有限公司", "en": "Orion Harbor Logistics", "short": "子丑寅卯"},
    ]


@pytest.fixture
def master():
    return Master.from_records(_master_records())


def test_norm_ignores_case_width_punctuation_suffix_and_leading_city():
    assert norm("（上海）甲乙丙机电有限公司") == norm("甲乙丙机电")
    assert norm("Alpha Beta Gamma Co., Ltd.") == norm("alpha-beta gamma")
    assert norm("ACME Inc.") == "acme"
    assert norm("") == ""
    assert norm(None) == ""


def test_exact_cn_and_en_match_score_one(master):
    m = master.match("甲乙丙机电有限公司")
    assert (m.code, m.method, m.score) == ("C001", "exact", 1.0)
    m = master.match("alpha beta gamma")
    assert (m.code, m.method, m.score) == ("C001", "exact", 1.0)
    assert m.candidates == []


def test_short_name_and_merged_code_match(master):
    m = master.match("丁戊己")
    assert (m.code, m.method) == ("C002", "exact")
    assert m.cn == "丁戊己贸易股份有限公司"


def test_unique_substring_is_contains_auto_score(master):
    m = master.match("Kestrel Microsystems")
    assert (m.code, m.method, m.score) == ("C004", "contains", 0.9)
    assert m.candidates == []


def test_shared_exact_name_is_ambiguous_not_auto():
    master = Master.from_records([
        {"source": "master", "code": "A1", "cn": "同名公司", "en": "", "short": ""},
        {"source": "master", "code": "A2", "cn": "同名公司", "en": "", "short": ""},
    ])
    m = master.match("同名公司")
    assert m.method == "ambiguous"
    assert m.score < 0.9
    assert {c[0] for c in m.candidates} == {"A1", "A2"}


def test_fuzzy_typo_with_clear_gap_is_auto_accepted(master):
    m = master.match("Kestrel Microsystms Corporation")
    assert (m.code, m.method) == ("C004", "fuzzy")
    assert 0.9 <= m.score <= 0.95


def test_unrelated_name_is_confident_new_company(master):
    m = master.match("Quux Zzyzx Unrelated")
    assert (m.method, m.score, m.code) == ("new", 1.0, None)
    assert m.candidates == []


def test_empty_query_is_none(master):
    m = master.match("")
    assert m.method == "none" and m.code is None


def test_empty_master_is_none_not_new():
    m = Master.from_records([]).match("Quux Zzyzx Unrelated")
    assert m.method == "none" and m.code is None


def _tricky_master():
    return Master.from_records([
        {"source": "master", "code": "CR0701256", "cn": "甲乙丙机电有限公司", "en": "", "short": ""},
        {"source": "master", "code": "CR992066", "cn": "丁戊己贸易有限公司", "en": "", "short": ""},
        {"source": "master", "code": "P1", "cn": "宁夏巨冈机械有限公司", "en": "", "short": ""},
        {"source": "master", "code": "P2", "cn": "壬癸子丑", "en": "", "short": ""},
        {"source": "master", "code": "F1", "cn": "甲乙丙丁戊己庚辛机电有限公司", "en": "", "short": ""},
        {"source": "master", "code": "E1", "cn": "", "en": "Delta Machinery Technology Co., Ltd.", "short": ""},
        {"source": "master", "code": "E2", "cn": "", "en": "Abacus Trading", "short": ""},
        {"source": "master", "code": "E3", "cn": "", "en": "Kestrel Microsystems Corporation", "short": ""},
    ])


def test_customer_code_matches_exactly_and_ignores_leading_zeros():
    master = _tricky_master()
    m = master.match("00CR992066 随便什么")
    assert (m.code, m.method, m.score) == ("CR992066", "code", 1.0)
    m = master.match("CR0701256")
    assert (m.code, m.method, m.score) == ("CR0701256", "code", 1.0)


def test_chinese_query_never_matches_english_only_company():
    m = _tricky_master().match("石英港")
    assert (m.code, m.method) == (None, "new")


def test_latin_query_never_matches_chinese_name_without_pinyin_equality():
    m = _tricky_master().match("Quartz Harbor")
    assert (m.code, m.method) == (None, "new")


def test_generic_industry_words_are_not_scored():
    master = _tricky_master()
    m = master.match("Quux Machinery Technology Co., Ltd.")
    assert (m.code, m.method) == (None, "new")
    m = master.match("乙丙机械科技有限公司")
    assert (m.code, m.method) == (None, "new")
    m = master.match("Machinery Technology Co., Ltd.")
    assert m.method == "none" and m.code is None


def test_province_prefix_is_ignored_in_both_scripts():
    master = _tricky_master()
    m = master.match("上海甲乙丙机电")
    assert (m.code, m.method, m.score) == ("CR0701256", "contains", 0.9)
    m = master.match("Shanghai Abacus Trading Co.")
    assert (m.code, m.method) == ("E2", "contains")


def test_mixed_script_query_compares_each_part_with_its_own_script():
    m = _tricky_master().match("宁夏巨冈 JUGANG")
    assert (m.code, m.method) == ("P1", "contains")


def test_latin_core_equal_to_pinyin_of_chinese_core_scores_095():
    m = _tricky_master().match("JUGANG")
    assert (m.code, m.method, m.score) == ("P1", "pinyin", 0.95)


def test_pinyin_prefix_of_at_least_six_letters_scores_09_when_unique():
    m = _tricky_master().match("renguizicho")
    assert (m.code, m.method, m.score) == ("P2", "pinyin", 0.9)


def test_short_latin_core_cannot_auto_match():
    m = _tricky_master().match("AB Co., Ltd.")
    assert m.code is None
    assert m.method == "review"
    assert m.score < 0.9


def test_middling_similarity_goes_to_review_with_top_candidates():
    m = _tricky_master().match("Kestrel Micro Labs")
    assert m.method == "review" and m.code is None
    assert m.candidates[0][0] == "E3"
    assert len(m.candidates) <= 3


def test_chinese_fuzzy_typo_with_clear_gap_is_fuzzy():
    m = _tricky_master().match("甲乙丙丁戊己庚辛机屯有限公司")
    assert (m.code, m.method) == ("F1", "fuzzy")


def test_size_counts_merged_companies(master):
    assert master.size == 5


def test_summary_counts_per_source_and_fields(master):
    s = master.summary()
    assert s["records"] == {"master": 2, "erp": 4}
    assert s["companies"] == 5
    assert s["with_code"] == 5
    assert s["with_cn"] == 4
    assert s["with_en"] == 4


def test_parse_rows_finds_header_on_row_one_with_float_codes():
    rows = [
        ["序号", "客户编码", "客户名称", "客户简称", "地区名称"],
        [9.0, "CR0001001", "测试公司有限公司", "测试公司", "北京"],
        [10.0, 1002.0, "", "只有简称", "上海"],
        [None, None, None, None, None],
    ]
    recs = parse_rows(rows, source="erp")
    assert recs == [
        {"source": "erp", "code": "CR0001001", "cn": "测试公司有限公司", "en": "", "short": "测试公司"},
        {"source": "erp", "code": "1002", "cn": "", "en": "", "short": "只有简称"},
    ]


def test_parse_rows_finds_header_on_row_two_after_title():
    rows = [
        ["Customer List", None, None],
        ["Customer Code", "CN name", "EN name"],
        ["CR0002070", "宜科有限公司", "Eke Ltd"],
    ]
    recs = parse_rows(rows, source="master")
    assert recs == [{"source": "master", "code": "CR0002070", "cn": "宜科有限公司", "en": "Eke Ltd", "short": ""}]


def test_parse_rows_without_header_yields_nothing():
    assert parse_rows([["a", "b"], ["c", "d"]], source="erp") == []


def test_load_missing_files_gives_empty_master(tmp_path):
    m = Master.load(tmp_path)
    assert m.size == 0
    assert m.match("anything").method == "none"


def test_load_reads_master_xlsx_by_glob(tmp_path):
    folder = tmp_path / "mail" / "staging" / "files" / "0a"
    folder.mkdir(parents=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "客户"
    ws.append(["Customer List"])
    ws.append(["Customer Code", "CN name", "EN name", "Province"])
    ws.append(["CR9000001", "甲乙丙机电有限公司", "Alpha Beta Gamma Co., Ltd.", "上海市"])
    ws.append([None, None, None, None])
    wb.save(folder / "ab_Customer_List_20260104.xlsx")

    m = Master.load(tmp_path)
    assert m.summary()["records"] == {"master": 1, "erp": 0}
    assert m.match("Alpha Beta Gamma").code == "CR9000001"


def test_load_tolerates_unreadable_file(tmp_path):
    folder = tmp_path / "mail" / "staging" / "files" / "0b"
    folder.mkdir(parents=True)
    (folder / "bad_Customer_List_20260104.xlsx").write_bytes(b"not a workbook")
    m = Master.load(tmp_path)
    assert m.size == 0
    assert m.summary()["skipped"] == ["bad_Customer_List_20260104.xlsx"]
