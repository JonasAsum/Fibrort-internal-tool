"""Company matching cases from real visit sheets (16 Oct 2026): what must and must not be suggested."""
from excel.companies import Master


def master(*rows):
    return Master.from_records([{"source": "erp", "code": c, "cn": cn, "en": en, "short": ""} for c, cn, en in rows])


def test_two_companies_that_share_only_industry_words_are_not_alike():
    m = master(("C1", "西安欧朗机电技术有限公司", "Xian Oulang Mechanical and Electrical Technology Co., Ltd."),
               ("C2", "上海熠欧机电设备有限公司", "Shanghai YIOU Mechanical and Electrical Co., Ltd."))
    got = m.match("Wuhan Jiashengyuan Mechanical and Electrical Equipment Co., Ltd.")
    assert got.method == "new" and not got.candidates


def test_a_name_that_is_only_a_city_and_an_industry_offers_no_candidates():
    m = master(("C1", "北京机床所精密机电有限公司", "Beijing Machine Tool Institute PRECISION MECHATRONICS Co., Ltd."),
               ("C2", "Electronic", "Electronic"),
               ("C3", "大连豪森瑞德设备制造有限公司", "AFRY Engineering(Shanghai)Co.Ltd"))
    got = m.match("Wuhan Machinery & Electrical Industry 武汉机电")
    assert got.method == "new" and got.code is None and not got.candidates


def test_erp_rows_with_the_code_new_stay_separate_companies():
    m = Master.from_records([
        {"source": "erp", "code": "New", "cn": "宁波海天精工机械有限公司", "en": "", "short": ""},
        {"source": "erp", "code": "New", "cn": "舒能(苏州)工业技术有限公司", "en": "", "short": ""},
        {"source": "erp", "code": "", "cn": "复旦大学", "en": "", "short": ""}])
    assert m.size == 3
    assert all(not c.aliases for c in m.companies)
    assert m.match("宁波海天精工机械有限公司").method == "exact"
    assert m.match("舒能（苏州）工业技术有限公司").method == "exact"  # not "ambiguous" with an unrelated company


def test_a_code_with_a_name_glued_on_is_read_as_the_code():
    m = Master.from_records([{"source": "erp", "code": "广西玉林市南丰物资有CR00011672", "cn": "广西玉林南丰物资有限公司", "en": "", "short": ""}])
    assert m.has_code("CR00011672")


def test_a_branch_in_another_city_is_not_the_first_suggestion():
    m = master(("C1", "", "Hengda Precision Co., Ltd. Shenyang"), ("C2", "", "Hengda Precision Co., Ltd. Wuhan"))
    got = m.match("Hengdo Precision Co., Ltd. Wuhan")
    assert got.method == "review" and got.candidates[0][0] == "C2"


def test_a_trailing_city_in_the_chinese_name_is_not_part_of_the_name():
    m = master(("C1", "珩述自动化科技(上海)有限公司", "Hengshu Automation Technology (Shanghai) Co., Ltd."))
    got = m.match("Hengshu Automation Technology Co., Ltd.")
    assert got.code == "C1"


def test_an_english_name_is_compared_with_the_pinyin_of_a_chinese_one_and_a_wrong_city_is_never_accepted():
    m = master(("C1", "宁波信泰机械有限公司", ""), ("C2", "苏州沃非特智能科技有限公司", ""))
    wrong_city = m.match("Liaoning Xintai Machinery Co., Ltd.")
    assert wrong_city.code is None  # Liaoning is not Ningbo, however alike the names are
    typo = m.match("Suzhou Wofeite Intelligent Technology Co., Ltd.")
    assert typo.candidates and typo.candidates[0][0] == "C2" or typo.code == "C2"


def test_a_pinyin_that_only_shares_a_syllable_is_not_offered():
    m = master(("C1", "上海大侨誉远精密机械股份有限公司", ""))
    got = m.match("Landaoqiyuan 岚道启源")
    assert not got.candidates


def test_a_visit_the_excel_reader_called_an_exhibition_is_not_a_customer_in_the_visit_pipeline():
    from visits.reader import Visit
    v = Visit(**{**{f: None for f in ()}, **_visit_args()})
    v.extra["kind"] = "exhibition"
    assert v.kind == "exhibition"


def _visit_args():
    import dataclasses
    from visits.reader import Visit
    out = {}
    for f in dataclasses.fields(Visit):
        if f.default is dataclasses.MISSING and f.default_factory is dataclasses.MISSING:
            out[f.name] = [] if "notes" in f.name else ""
    return out
