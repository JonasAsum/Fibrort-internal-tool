"""Contact-list importer (folder 06): both layouts, copies, company matching, the Margince adapter
against tests/fake_crm.py. Every value here is invented; the real-folder test asserts counts only."""
import glob
import os
import shutil
from unittest import mock

import pytest

import excel_margince
from excel import contacts
from excel.companies import Master
from tests.conftest import SAMPLE_DIR
from tests.excel_book import write_book
from tests.fake_crm import FakeCRM, _Conflict

# a company list whose one entry looks like none of the invented names below
MASTER_RECORDS = [{"source": "master", "code": "KD00001", "cn": "宇宙星辰科技有限公司", "en": "Cosmos Stellar Co., Ltd.",
                   "short": ""}]


@pytest.fixture
def master():
    return Master.from_records(MASTER_RECORDS)


SECTION_HEAD = {"B3": "Contact Person", "C3": "E-mail Addr.", "D3": "Phone No.", "E3": "Position", "F3": "Remark"}


def sections_book(tmp_path, name="handover.xlsx", extra=None):
    cells = {"A1": "Handover list (invented)", **SECTION_HEAD,
             "A4": "Order Desk _ FIBRO Testtische GmbH",
             "A5": "Tables", "B5": "Anna Müller ", "C5": "a.mueller@fibro-test.example",
             "D5": "T  +49 (7134) 73-275", "E5": "Contract management",
             "B6": "Bernd Schäfer", "C6": "b.schaefer@fibro-test.example",
             "A8": "Order Confirmation_FIBRO Testtische GmbH",
             "A9": "Tables PO", "B9": "Clara Beck", "D9": "+49 7134 73-213",
             "A11": "Key Parts For Assembly",
             "A12": "Motors (Volt Test)", "B12": "Alan Test 测试", "C12": "alan@volt-test.example",
             "D12": "15800000001", "E12": "Sales", "F12": "China",
             "A13": "Sensor Cable (Sensa)", "B13": "Wu Test", "D13": "38938741",
             "A14": "Fast Post (Urgent Parts)", "B14": "无", "F14": "Shanghai",
             "A15": "Gearbox (Drive Test GmbH)", "F15": "Germany",
             "A17": "Shipping",
             "A18": "苏州测试货运有限公司 (Goods)", "B18": "Wang Test", "D18": "13900000002",
             "A19": "Is there another forwarder ?",
             **(extra or {})}
    return write_book(tmp_path / name, {"Contacts List": {"cells": cells}})


ZH_HEAD = {"A1": "序号", "D1": "公司名称", "E1": "地址", "F1": "联系人", "G1": "电话"}


def zh_rows(filled):
    rows = [("应收", "客户", "常州测试智能有限公司", "北京市测试路1号", "关测试", "13600000001"),
            ("预付", "供应商", "上海测试软件有限公司", "上海市测试路2号", "俞测试", "136 2165 9714"),
            ("其他应付", "供应商", "艾测试国际货运（上海）有限公司", "上海市静安区", "Test Yao", "021-60000000")]
    cells = dict(ZH_HEAD)
    for i, (ledger, kind, company, address, person, phone) in enumerate(rows, start=2):
        cells.update({f"A{i}": str(i - 1), f"B{i}": ledger, f"C{i}": kind, f"D{i}": company})
        if filled:
            cells.update({f"E{i}": address, f"F{i}": person, f"G{i}": phone})
    return cells


def zh_book(tmp_path, name="zh.xlsx", filled=True):
    return write_book(tmp_path / name, {"Sheet1": {"cells": zh_rows(filled)}})


def by_name(result):
    return {c.full_name: c for c in result.contacts}


def reasons(result):
    return sorted(i.reason for i in result.review)


# ---------- sections layout ----------

def test_sections_layout_reads_people_with_header_below_row_one(tmp_path, master):
    r = contacts.read_paths([sections_book(tmp_path)], master)
    people = by_name(r)
    assert set(people) == {"Anna Müller", "Bernd Schäfer", "Clara Beck", "Alan Test 测试", "Wu Test", "Wang Test"}
    anna = people["Anna Müller"]  # trailing space and umlaut kept as written, space trimmed
    assert (anna.kind, anna.company, anna.role) == ("internal", "FIBRO Testtische GmbH", "Contract management")
    assert anna.department == "Tables" and anna.phones == ["+49713473275"]
    assert anna.emails == ["a.mueller@fibro-test.example"]
    assert people["Bernd Schäfer"].department == "Tables"  # the team label carries down to the next row


def test_internal_people_share_one_company_named_by_the_heading(tmp_path, master):
    r = contacts.read_paths([sections_book(tmp_path)], master)
    assert {c.company for c in r.contacts if c.kind == "internal"} == {"FIBRO Testtische GmbH"}
    assert sum(o.kind == "internal" for o in r.orgs.values()) == 1


@pytest.mark.parametrize("label,company", [
    ("Motors (Volt Test)", "Volt Test"),
    ("Fast Post (Urgent Parts)", "Fast Post"),
    ("Gearbox (Drive Test GmbH)", "Drive Test GmbH"),
    ("Gearbox （Drive Test GmbH）", "Drive Test GmbH"),
    ("Dasher (Exhibits Trip) ?", "Dasher"),
    ("Chen Test 太仓测试运输有限公司", "太仓测试运输有限公司"),
    ("Motors (Siemens) 上海测试实业有限公司 (Siemens Dealer)", "上海测试实业有限公司"),
    ("ZT 中通物流", "ZT 中通物流"),
    ("苏州测试（太仓）运输有限公司", "苏州测试(太仓)运输有限公司"),
    ("Plain Supplier", "Plain Supplier"),
])
def test_company_from_label(label, company):
    assert contacts.company_from_label(label) == company


def test_supplier_rows_and_company_only_rows(tmp_path, master):
    r = contacts.read_paths([sections_book(tmp_path)], master)
    names = {o.name: o for o in r.orgs.values()}
    assert {"Volt Test", "Sensa", "Fast Post", "Drive Test GmbH", "苏州测试货运有限公司"} <= set(names)
    assert all(names[n].kind == "supplier" for n in ("Volt Test", "Fast Post", "苏州测试货运有限公司"))
    assert not any(c.company in ("Fast Post", "Drive Test GmbH") for c in r.contacts)  # "无" and no name: no contact


def test_skipped_rows_have_reasons_and_blank_rows_are_counted(tmp_path, master):
    r = contacts.read_paths([sections_book(tmp_path)], master)
    sections = [i for i in r.review if i.reason == "section"]
    assert len(sections) == 5  # two FIBRO headings, Key Parts, Shipping, the trailing question
    assert contacts.summarize(r)["blank_rows"] >= 3
    assert contacts.summarize(r)["skipped_by_reason"]["section"] == 5


def test_phone_without_country_stays_out_and_is_reported(tmp_path, master):
    r = contacts.read_paths([sections_book(tmp_path)], master)
    wu = by_name(r)["Wu Test"]  # 8 digits, no area code: not guessed
    assert wu.phones == []
    assert [i.value for i in r.review if i.reason == "phone_no_country"] == ["38938741"]
    assert by_name(r)["Wang Test"].phones == ["+8613900000002"]
    assert by_name(r)["Clara Beck"].emails == []  # no email is no problem


def test_german_number_without_plus_is_not_read_as_china(tmp_path, master):
    p = sections_book(tmp_path, extra={"D9": "07134 73213"})
    r = contacts.read_paths([p], master)
    assert by_name(r)["Clara Beck"].phones == []
    assert "phone_no_country" in reasons(r)


def test_fax_is_not_a_phone(tmp_path, master):
    p = sections_book(tmp_path, extra={"D9": "F +49 7134 73-999"})
    assert by_name(contacts.read_paths([p], master))["Clara Beck"].phones == []


def test_bad_email_is_reported_and_the_contact_kept(tmp_path, master):
    p = sections_book(tmp_path, extra={"C13": "ask reception"})
    r = contacts.read_paths([p], master)
    assert by_name(r)["Wu Test"].emails == []
    assert [i.value for i in r.review if i.reason == "bad_email"] == ["ask reception"]


def test_name_less_row_with_an_address_is_not_a_contact(tmp_path, master):
    p = sections_book(tmp_path, extra={"C15": "info@drive-test.example"})
    r = contacts.read_paths([p], master)
    assert "Drive Test GmbH" in {o.name for o in r.orgs.values()}
    assert "no_contact_name" in reasons(r) and not any(c.company == "Drive Test GmbH" for c in r.contacts)


# ---------- list layout ----------

def test_list_layout_with_unlabelled_type_columns(tmp_path, master):
    r = contacts.read_paths([zh_book(tmp_path)], master)
    assert by_name(r)["关测试"].kind == "customer"
    assert by_name(r)["俞测试"].kind == "supplier" and by_name(r)["俞测试"].phones == ["+8613621659714"]
    assert by_name(r)["Test Yao"].phones == ["+862160000000"]
    assert {o.name: o.address for o in r.orgs.values()}["常州测试智能有限公司"] == "北京市测试路1号"
    assert by_name(r)["关测试"].ledger == "应收"


def test_columns_in_any_order(tmp_path, master):
    cells = {"A2": "Phone", "B2": "Company", "C2": "Contact Name", "D2": "Email",
             "A3": "+49 30 1234567", "B3": "Orderly Test GmbH", "C3": "Eva Test", "D3": "eva@orderly-test.example"}
    r = contacts.read_paths([write_book(tmp_path / "x.xlsx", {"S": {"cells": cells}})], master)
    assert [(c.company, c.full_name, c.phones, c.emails) for c in r.contacts] == [
        ("Orderly Test GmbH", "Eva Test", ["+49301234567"], ["eva@orderly-test.example"])]


def test_row_without_company_is_skipped_with_a_reason(tmp_path, master):
    cells = zh_rows(True)
    cells.update({"A9": "9", "F9": "孤立测试", "G9": "13600000009"})
    r = contacts.read_paths([write_book(tmp_path / "x.xlsx", {"Sheet1": {"cells": cells}})], master)
    assert "no_company" in reasons(r) and "孤立测试" not in by_name(r)


def test_a_visit_list_is_not_a_contact_list(book, master):
    path = book("visits.xlsx", {"S": {"cells": {"B8": "Date", "C8": "Customer", "D8": "Contact Name",
                                                 "E8": "Phone No.", "F8": "Email"}}})
    r = contacts.read_paths([path], master)
    assert r.contacts == [] and r.per_file_stats[0].layout == "none"


# ---------- copies and duplicates ----------

def test_blank_copy_does_not_win_or_add_contacts(tmp_path, master):
    r = contacts.read_paths([zh_book(tmp_path, "a.xlsx", filled=False), zh_book(tmp_path, "b.xlsx", filled=True)], master)
    assert len(r.contacts) == 3 and len(r.orgs) == 3
    assert all(c.phones and c.company_key in r.orgs for c in r.contacts)
    assert {o.address for o in r.orgs.values()} >= {"北京市测试路1号"}  # the blank copy never blanks a field


def test_blank_copy_alone_gives_companies_and_no_contacts(tmp_path, master):
    r = contacts.read_paths([zh_book(tmp_path, filled=False)], master)
    assert r.contacts == [] and len(r.orgs) == 3


def test_richer_sections_copy_wins_and_poorer_one_adds_nothing(tmp_path, master):
    poor = sections_book(tmp_path, "poor.xlsx", extra={"B12": "Alan Test 测试", "C12": None, "D12": None, "E12": None})
    rich = sections_book(tmp_path, "rich.xlsx")
    for order in ([poor, rich], [rich, poor]):
        r = contacts.read_paths(order, master)
        alan = by_name(r)["Alan Test 测试"]
        assert alan.emails == ["alan@volt-test.example"] and alan.phones == ["+8615800000001"] and alan.role == "Sales"
        assert alan.copies == 2 and len(r.contacts) == 6


def test_byte_identical_copy_is_read_once(tmp_path, master):
    a = sections_book(tmp_path, "a.xlsx")
    shutil.copy(a, tmp_path / "copy of a.xlsx")
    r = contacts.read_paths([a, str(tmp_path / "copy of a.xlsx")], master)
    assert len(r.per_file_stats) == 1 and r.skipped_files[0]["error"].startswith("identical to")


def test_duplicate_rows_in_one_sheet_make_one_contact(tmp_path, master):
    extra = {"A20": "苏州测试货运有限公司 (Goods)", "B20": "Wang  Test", "D20": "13900000002", "E20": "Driver"}
    r = contacts.read_paths([sections_book(tmp_path, extra=extra)], master)
    wang = [c for c in r.contacts if c.full_name == "Wang Test"]
    assert len(wang) == 1 and wang[0].role == "Driver" and wang[0].copies == 2


def test_same_person_at_two_companies_stays_two_contacts(tmp_path, master):
    cells = {"A1": "Company", "B1": "Contact", "C1": "Phone",
             "A2": "One Test Ltd", "B2": "Pat Test", "C2": "13600000011",
             "A3": "Two Test Ltd", "B3": "Pat Test", "C3": "13600000011"}
    r = contacts.read_paths([write_book(tmp_path / "x.xlsx", {"S": {"cells": cells}})], master)
    assert len(r.contacts) == 2


# ---------- never abort ----------

def test_bad_file_is_recorded_and_the_rest_read(tmp_path, master):
    bad = tmp_path / "broken.xlsx"
    bad.write_text("not a workbook")
    r = contacts.read_paths([str(bad), zh_book(tmp_path)], master)
    assert len(r.contacts) == 3
    assert r.skipped_files and r.skipped_files[0]["error"].startswith("could not be read")


def test_empty_and_unrecognised_sheets_do_not_raise(tmp_path, master):
    cells = {f"A{i}": f"x{i}" for i in range(1, 15)}
    path = write_book(tmp_path / "x.xlsx", {"Empty": {"cells": {}}, "Notes": {"cells": cells}})
    r = contacts.read_paths([path], master)
    assert r.contacts == [] and r.per_file_stats[0].unrecognised == ["Notes"]


# ---------- company matching ----------

def test_known_company_is_filed_under_the_masters_name(tmp_path, master):
    cells = {"A1": "公司名称", "B1": "联系人", "C1": "电话", "A2": "宇宙星辰科技集团", "B2": "测试员", "C2": "13600000021"}
    r = contacts.read_paths([write_book(tmp_path / "x.xlsx", {"S": {"cells": cells}})], master)
    org = r.orgs[r.contacts[0].company_key]
    assert org.name == "宇宙星辰科技有限公司" and org.match["code"] == "KD00001" and not org.queued


def test_similar_company_is_queued_with_candidates_and_its_people_stay_out(tmp_path, master):
    cells = {"A1": "Company", "B1": "Contact", "C1": "Phone", "A2": "Cosmos Stellar Works",
             "B2": "Tess Test", "C2": "13600000022", "A3": "Unrelated Test Ltd", "B3": "Una Test", "C3": "13600000023"}
    r = contacts.read_paths([write_book(tmp_path / "x.xlsx", {"S": {"cells": cells}})], master)
    queued = [i for i in r.review if i.reason == "company"]
    assert [i.value for i in queued] == ["Cosmos Stellar Works"] and queued[0].candidates
    assert [c.full_name for c in r.accepted_contacts()] == ["Una Test"]
    s = contacts.summarize(r)
    assert (s["contacts"], s["contacts_accepted"], s["contacts_queued"], s["companies_queued"]) == (2, 1, 1, 1)
    assert s["queued_by_reason"] == {"company": 1}


def test_internal_colleagues_are_not_matched_against_customers(tmp_path):
    odd = Master.from_records([{"source": "master", "code": "KD1", "cn": "", "en": "FIBRO Testtische GmbH", "short": ""}])
    r = contacts.read_paths([sections_book(tmp_path)], odd)
    assert not any(o.queued for o in r.orgs.values() if o.kind == "internal")


# ---------- bank and tax details stay out ----------

SECRETS = ["6222 0200 1234 5678", "91310000MA1K3XYZ12", "DE89 3704 0044 0532 0130 00", "Tax ID 123456789",
           "开户行 工商银行 测试支行", "VAT DE123456789"]


def secret_book(tmp_path):
    cells = {"A1": "公司名称", "B1": "联系人", "C1": "电话", "D1": "银行账号", "E1": "税号", "F1": "备注", "G1": "职位",
             "H1": "Bank account", "I1": "IBAN", "J1": "地址",
             "A2": "密测试有限公司", "B2": "丙测试", "C2": "13600000041", "D2": SECRETS[0], "E2": SECRETS[1],
             "F2": SECRETS[4], "G2": "Tax ID 123456789", "H2": SECRETS[2], "I2": SECRETS[2], "J2": "VAT DE123456789",
             "A3": "密测试二有限公司", "B3": "丁测试", "C3": "13600000042", "F3": "wants a call back", "G3": "Buyer"}
    return write_book(tmp_path / "secret.xlsx", {"S": {"cells": cells}})


def test_bank_and_tax_columns_and_text_never_reach_a_record_or_the_crm(tmp_path, master):
    result = contacts.read_paths([secret_book(tmp_path)], master)
    crm = FakeCRM()
    commit(crm, result)
    dump = repr((result.contacts, result.orgs, result.review, crm.companies, crm.contacts, crm.fields))
    for secret in SECRETS + ["6222", "MA1K3XYZ12", "0532", "123456789"]:
        assert secret not in dump
    assert sorted(i.field for i in result.review if i.reason == "withheld_sensitive") == ["address", "position", "remark"]
    ok = by_name(result)["丁测试"]
    assert (ok.remark, ok.role) == ("wants a call back", "Buyer")  # ordinary text is kept
    assert person(crm, "丙测试")["phones"][0]["phone"] == "+8613600000041"


# ---------- Margince adapter, against the in-memory CRM ----------

def commit(crm, result):
    return excel_margince.contacts_commit(crm, result)


def full_result(tmp_path, master):
    return contacts.read_paths([sections_book(tmp_path), zh_book(tmp_path)], master)


def test_commit_writes_companies_contacts_titles(tmp_path, master):
    crm, result = FakeCRM(), full_result(tmp_path, master)
    out = commit(crm, result)
    names = {c["display_name"] for c in crm.companies.values()}
    assert {"FIBRO Testtische GmbH (internal)", "Volt Test", "常州测试智能有限公司"} <= names
    assert "FIBRO Testtische GmbH" not in names
    assert len(crm.contacts) == len(result.contacts) == 9
    assert out["contacts_created"] == 9 and out["titles_set"] == 2 and out["contacts_without_company"] == 0
    anna = next(c for c in crm.contacts.values() if c["full_name"] == "Anna Müller")
    assert anna["phones"][0]["phone"] == "+49713473275" and anna["title"] == "Contract management"
    kinds = {c["display_name"]: c.get("relationship_types") for c in crm.companies.values()}
    assert kinds["Volt Test"] == ["supplier"] and kinds["常州测试智能有限公司"] == ["customer"]
    assert kinds["FIBRO Testtische GmbH (internal)"] == ["other"]
    assert next(c for c in crm.companies.values() if c["display_name"] == "常州测试智能有限公司")["lifecycle"] == "customer"
    assert next(c for c in crm.companies.values() if c["display_name"] == "上海测试软件有限公司")["address"] == {"line1": "上海市测试路2号"}


def company(crm, name):
    return next(c for c in crm.companies.values() if c["display_name"] == name)


def person(crm, name):
    return next(c for c in crm.contacts.values() if c["full_name"] == name)


def test_origin_and_ledger_tags_let_the_owner_filter(tmp_path, master):
    crm = FakeCRM()
    commit(crm, full_result(tmp_path, master))
    assert company(crm, "常州测试智能有限公司")["cf_contact_list_origin"] == "Finance list: customer"
    assert company(crm, "常州测试智能有限公司")["cf_ledger_type"] == "应收"
    assert company(crm, "上海测试软件有限公司")["cf_ledger_type"] == "预付"
    assert company(crm, "Volt Test")["cf_contact_list_origin"] == "Handover list: supplier"
    assert "cf_ledger_type" not in company(crm, "Volt Test")
    assert company(crm, "FIBRO Testtische GmbH (internal)")["cf_contact_list_origin"] == "Handover list: internal colleague"
    assert person(crm, "Anna Müller")["cf_contact_origin"] == "Handover list: internal colleague"
    assert person(crm, "Anna Müller")["cf_department"] == "Tables"
    assert person(crm, "Alan Test 测试")["cf_remark"] == "China"
    assert person(crm, "关测试")["cf_contact_origin"] == "Finance list: customer"
    assert {f["label"] for f in crm.fields} == {"Contact list origin", "Ledger type", "Contact origin", "Department", "Remark"}


def test_only_the_picklist_options_the_tags_use_are_offered(tmp_path, master):
    crm = FakeCRM()
    commit(crm, full_result(tmp_path, master))
    used = {c["cf_contact_origin"] for c in crm.contacts.values()}
    assert used <= set(next(f for f in crm.fields if f["label"] == "Contact origin")["options"])


def test_internal_staff_never_land_on_a_customer_of_the_same_name(tmp_path, master):
    crm = FakeCRM()
    crm.companies["pre9"] = {"id": "pre9", "display_name": "FIBRO Testtische GmbH", "lifecycle": "customer"}
    commit(crm, contacts.read_paths([sections_book(tmp_path)], master))
    assert crm.employment and all(co != "pre9" for co, _ in crm.employment)
    assert crm.companies["pre9"] == {"id": "pre9", "display_name": "FIBRO Testtische GmbH", "lifecycle": "customer"}


def test_company_on_both_ledgers_is_customer_and_supplier(tmp_path, master):
    cells = {"A1": "公司名称", "B1": "应收/应付", "C1": "联系人", "D1": "电话",
             "A2": "双向测试有限公司", "B2": "应收", "C2": "甲测试", "D2": "13600000031",
             "A3": "双向测试有限公司", "B3": "应付", "C3": "乙测试", "D3": "13600000032"}
    crm = FakeCRM()
    commit(crm, contacts.read_paths([write_book(tmp_path / "x.xlsx", {"S": {"cells": cells}})], master))
    both = company(crm, "双向测试有限公司")
    assert both["relationship_types"] == ["customer", "supplier"] and both["cf_ledger_type"] == "应收, 应付"


def test_existing_values_are_never_overwritten(tmp_path, master):
    crm = FakeCRM()
    crm.companies["pre1"] = {"id": "pre1", "display_name": "Volt Test", "lifecycle": "prospect",
                            "relationship_types": ["partner"], "address": {"line1": "Typed by hand"},
                            "cf_contact_list_origin": "Typed by hand"}
    commit(crm, contacts.read_paths([sections_book(tmp_path)], master))
    volt = crm.companies["pre1"]
    assert volt["relationship_types"] == ["partner", "supplier"] and volt["lifecycle"] == "prospect"
    assert volt["cf_contact_list_origin"] == "Typed by hand" and volt["address"] == {"line1": "Typed by hand"}


def test_second_commit_writes_nothing(tmp_path, master):
    crm = FakeCRM()
    commit(crm, full_result(tmp_path, master))
    snapshot = {k: dict(v) for k, v in crm.contacts.items()}, {k: dict(v) for k, v in crm.companies.items()}
    again = commit(crm, full_result(tmp_path, master))
    assert (again["contacts_created"], again["contacts_updated"], again["titles_set"], again["companies_enriched"]) == (0, 0, 0, 0)
    assert again["companies"] == {"already_present": len(crm.companies)}
    assert ({k: dict(v) for k, v in crm.contacts.items()}, {k: dict(v) for k, v in crm.companies.items()}) == snapshot
    assert len(crm.employment) == len(set(crm.employment))


def test_a_later_file_adds_only_what_is_new(tmp_path, master):
    crm = FakeCRM()
    commit(crm, contacts.read_paths([zh_book(tmp_path)], master))
    extra = {"A21": "苏州测试货运有限公司 (Goods)", "B21": "New Test", "D21": "13900000099"}
    out = commit(crm, contacts.read_paths([zh_book(tmp_path), sections_book(tmp_path, extra=extra)], master))
    assert out["contacts_created"] == 7  # 6 + New Test - the 3 already there


def test_queued_company_is_not_written(tmp_path, master):
    cells = {"A1": "Company", "B1": "Contact", "C1": "Phone", "A2": "Cosmos Stellar Works",
             "B2": "Tess Test", "C2": "13600000022"}
    crm = FakeCRM()
    commit(crm, contacts.read_paths([write_book(tmp_path / "x.xlsx", {"S": {"cells": cells}})], master))
    assert crm.companies == {} and crm.contacts == {}


def test_duplicate_email_reuses_only_when_the_name_matches(tmp_path, master):
    crm = FakeCRM()
    crm.contacts["held1"] = {"id": "held1", "full_name": "Anna Müller", "emails": [], "phones": []}
    crm.contacts["held2"] = {"id": "held2", "full_name": "Somebody Else", "emails": [], "phones": []}
    crm.contact_conflict["a.mueller@fibro-test.example"] = "held1"
    crm.contact_conflict["b.schaefer@fibro-test.example"] = "held2"
    commit(crm, contacts.read_paths([sections_book(tmp_path)], master))
    assert sum(c["full_name"] == "Anna Müller" for c in crm.contacts.values()) == 1  # reused, linked to the new company
    bernd = next(c for c in crm.contacts.values() if c["full_name"] == "Bernd Schäfer")
    assert bernd["emails"] == []  # the address belongs to somebody else


def test_preview_reads_only(tmp_path, master):
    crm = FakeCRM()
    crm.run_csv_import = mock.Mock(side_effect=AssertionError("preview must not write"))
    crm.log_activity = crm.create_contact = crm.patch = crm.run_csv_import
    plan, titles = excel_margince.contacts_plan(full_result(tmp_path, master))
    assert excel_margince.contacts_preview(crm, plan) == {
        "companies_new": len(plan.companies), "companies_present": 0, "contacts_new": 9, "contacts_present": 0}
    assert sum("title" in e for e in titles.values()) == 2


def test_preview_after_commit_finds_everything_present(tmp_path, master):
    crm, result = FakeCRM(), full_result(tmp_path, master)
    commit(crm, result)
    plan, _ = excel_margince.contacts_plan(result)
    p = excel_margince.contacts_preview(crm, plan)
    assert (p["companies_new"], p["contacts_new"], p["contacts_present"]) == (0, 0, 9)


class _LoginCRM(FakeCRM):
    def login(self, email, password):
        pass


def test_cli_preview_and_commit(tmp_path, master, capsys):
    crm = _LoginCRM()
    zh_book(tmp_path, "list.xlsx")
    args = ["--family", "contacts", "--base-url", "http://fake", str(tmp_path / "list.xlsx")]
    with mock.patch("excel_margince.MargenceClient", return_value=crm), mock.patch("excel.run._master", return_value=master):
        excel_margince.main([*args, "--mode", "preview"])
        assert crm.companies == {}
        assert "would write: 3 companies (0 already there), 3 contacts" in capsys.readouterr().out
        excel_margince.main([*args, "--mode", "commit"])
        assert len(crm.contacts) == 3
        out = capsys.readouterr().out
        assert "3 created" in out
        excel_margince.main([*args, "--mode", "commit"])
        assert "0 created" in capsys.readouterr().out and len(crm.contacts) == 3


# ---------- the real sample folder (counts only) ----------

REAL = os.path.join(SAMPLE_DIR, "06 Contact Lists")


@pytest.fixture
def real_files():
    files = sorted(glob.glob(os.path.join(REAL, "*.xlsx")))
    if not files:
        pytest.skip(f"sample folder not found: {REAL}")
    return files


def test_real_folder_counts(real_files, master):
    r = contacts.read_paths(real_files, master)
    s = contacts.summarize(r)
    assert (s["files"], s["read_sheets"], s["skipped_files"]) == (4, 4, 0)
    assert (s["contacts"], s["companies"], s["companies_queued"]) == (27, 27, 0)
    assert r.skipped_files == [] and all(st.layout != "none" for st in r.per_file_stats)


def test_real_folder_blank_copy_adds_no_contacts(real_files, master):
    blank = [f for f in real_files if "客户供应商" in f and contacts.read_paths([f], master).contacts == []]
    assert len(blank) == 1
    with_blank = contacts.read_paths(real_files, master)
    without = contacts.read_paths([f for f in real_files if f not in blank], master)
    assert len(with_blank.contacts) == len(without.contacts) and len(with_blank.orgs) == len(without.orgs)


def test_real_folder_against_the_real_company_master(real_files):
    from excel.run import IMPORTER_DIR, _master
    real_master = _master(IMPORTER_DIR)
    if real_master.size == 0:
        pytest.skip("no company master on this machine")
    r = contacts.read_paths(real_files, real_master)
    s = contacts.summarize(r)
    assert (s["contacts"], s["companies"], s["contacts_accepted"], s["companies_queued"]) == (27, 27, 23, 4)
    assert s["queued_by_reason"] == {"company": 4}
    assert sum(o.kind == "internal" for o in r.orgs.values()) == 1 and all(o.name.endswith(" (internal)") for o in r.orgs.values() if o.kind == "internal")
    crm = FakeCRM()
    first = commit(crm, r)
    second = commit(crm, r)
    assert first["contacts_created"] == 23 and (second["contacts_created"], second["fields_set"], second["titles_set"]) == (0, 0, 0)
