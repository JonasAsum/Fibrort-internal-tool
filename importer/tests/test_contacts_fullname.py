"""Contacts group by FULL name within a company, never by surname alone."""
from visits.match import match
from visits.reader import Visit


def _v(i, customer, contact, owner="Owner A", phone="", email=""):
    return Visit(workbook="w.xlsx", sheet="s", owner=owner, year=2026, cw=1, excel_row_start=i,
                 excel_row_end=i, date_raw="", date_iso="2026-01-05", city="Ningbo",
                 customer_raw=customer, potential="", customer_type_raw="customer",
                 contact_raw=contact, phone_raw=phone, office_phone="", email=email, purpose="")


def _run(*rows):
    visits = [_v(i + 1, *r) if isinstance(r, tuple) else _v(i + 1, "Alpha Motors", r)
              for i, r in enumerate(rows)]
    res = match(visits)
    return res, [res.links[v.visit_key][1] for v in visits]


def test_same_name_variants_merge():
    res, keys = _run("Deyi.Zhao", "Zhao Deyi", "DEYI ZHAO", "Mr. Deyi Zhao", "deyi-zhao")
    assert len(set(keys)) == 1 and len(res.contacts) == 1


def test_same_surname_different_people_stay_apart():
    res, keys = _run("Wei.Li", "Na.Li", "Li Wei")
    assert keys[0] == keys[2] and keys[0] != keys[1] and len(res.contacts) == 2


def test_bare_surname_not_merged_into_full_name():
    res, keys = _run("Mr. Li", "Wei.Li", "Li", "Ms.Ouyang", "Ouyang Mei")
    assert keys[0] == keys[2]          # two bare "Li" rows may fold together
    assert keys[0] != keys[1]
    assert keys[3] != keys[4]
    assert len(res.contacts) == 4


def test_same_name_two_companies_stay_apart():
    res, keys = _run(("Alpha Motors", "Wei.Li"), ("Beta Hydraulics", "Wei.Li"))
    assert keys[0] != keys[1] and len(res.contacts) == 2
    assert {c.company_key for c in res.contacts.values()} == set(res.companies)


def test_chinese_names_compared_as_written():
    res, keys = _run("张伟", "张伟先生", "张娜", "李伟")
    assert keys[0] == keys[1]
    assert len({keys[0], keys[2], keys[3]}) == 3


def test_word_order_irrelevant():
    _, keys = _run("Peng Xiuming", "Xiuming.Peng", "XIUMING PENG")
    assert len(set(keys)) == 1


def test_empty_and_none_contact_has_no_contact():
    res, keys = _run("", None, "   ")
    assert keys == [None, None, None]  # empty, missing and whitespace-only cells are no contact (and no crash)
    assert not res.contacts


def test_salutation_only_is_its_own_contact():
    res, keys = _run("Mr.", "Ms.", "Mr", "Dr. Wei.Li")
    assert keys[:3] == [None, None, None]  # a bare "Mr."/"Ms." is no person (looks_like_person_name): correct over complete
    assert keys[3] is not None
    assert len(res.contacts) == 1


def test_unicode_width_variants_merge():
    _, keys = _run("Ｄｅｙｉ．Ｚｈａｏ", "Deyi.Zhao", "Deyi　Zhao")
    assert len(set(keys)) == 1


def test_run_together_merges_only_when_safe():
    res, keys = _run("Zhangwei", "Zhang Wei", "Wei Zhang")
    assert len(set(keys)) == 1 and len(res.contacts) == 1
    # Ambiguous: "Zhangwei" fits two different split names, so keep apart.
    res, keys = _run("Lixin", "Li Xin", "Xin Li", "Lix In")
    assert keys[0] != keys[1]
    # Not related at all.
    res, keys = _run("Zhangwei", "Zhang Wen")
    assert keys[0] != keys[1]


def test_merged_contact_keeps_data_and_fullest_name():
    res, keys = _run(("Alpha Motors", "Mr.Zhao", "A", "111", ""),
                     ("Alpha Motors", "Deyi.Zhao", "B", "222", "d@x.cn"),
                     ("Alpha Motors", "Zhao Deyi", "A", "111", ""))
    assert len(res.contacts) == 2 and keys[1] == keys[2] != keys[0]
    c = res.contacts[keys[1]]
    assert c.full_name in ("Deyi.Zhao", "Zhao Deyi")
    assert c.emails == ["d@x.cn"] and c.phones == ["222", "111"]
    assert c.owners == ["B", "A"] and c.spellings == ["Deyi.Zhao", "Zhao Deyi"]


def test_run_together_merge_updates_links_and_data():
    res, keys = _run(("Alpha Motors", "Zhangwei", "A", "1", ""), ("Alpha Motors", "Zhang Wei", "B", "2", "z@x.cn"))
    assert keys[0] == keys[1] and keys[0] in res.contacts
    c = res.contacts[keys[0]]
    assert set(c.spellings) == {"Zhangwei", "Zhang Wei"} and set(c.phones) == {"1", "2"}
    assert set(c.owners) == {"A", "B"} and c.emails == ["z@x.cn"]


def test_very_long_name():
    long = ".".join(["Given"] * 400 + ["Family"])
    _, keys = _run(long, long.upper(), "Given.Family")
    assert keys[0] == keys[1] != keys[2]


def test_many_rows():
    rows = [f"Given{i % 50}.Li" for i in range(2000)]
    res, keys = _run(*rows)
    assert len(res.contacts) == 50 and all(k for k in keys)
