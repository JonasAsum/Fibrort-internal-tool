"""A contact of folder 06 carries every +86 / international number of its phone cell, a mobile first, and
a number nobody can place is reported, not guessed. Invented workbooks, written to the in-memory CRM."""
import pytest

import excel_margince
from excel import contacts
from excel.companies import Master
from tests.excel_book import write_book
from tests.fake_crm import FakeCRM


@pytest.fixture
def master():
    # one company that looks like none of the invented names: an empty master would queue every company for a person's answer
    return Master.from_records([{"source": "master", "code": "KD00001", "cn": "宇宙星辰科技有限公司", "en": "Cosmos Stellar Co., Ltd.", "short": ""}])


def book(tmp_path, phone_cells):
    cells = {"A1": "公司名称", "B1": "联系人", "C1": "电话", "D1": "Email"}
    for i, (company, person, phone, mail) in enumerate(phone_cells, start=2):
        cells.update({f"A{i}": company, f"B{i}": person, f"C{i}": phone, f"D{i}": mail})
    return write_book(tmp_path / "list.xlsx", {"S": {"cells": cells}})


def test_several_numbers_of_one_cell_reach_the_crm_contact(tmp_path, master):
    r = contacts.read_paths([book(tmp_path, [("测试甲有限公司", "王五", "021-55550001 / 13800000001 / +49 7134 73-275",
                                              "wang@example.invalid")])], master)
    [c] = r.contacts
    assert c.phones == ["+862155550001", "+8613800000001", "+49713473275"]
    crm = FakeCRM()
    excel_margince.contacts_commit(crm, r)
    [made] = crm.contacts.values()
    assert [(p["phone"], p["phone_type"], p["is_primary"]) for p in made["phones"]] == [
        ("+8613800000001", "mobile", True), ("+862155550001", "work", False), ("+49713473275", "work", False)]
    assert [e["email"] for e in made["emails"]] == ["wang@example.invalid"]


def test_same_name_at_two_companies_stays_two_contacts_with_their_own_numbers(tmp_path, master):
    r = contacts.read_paths([book(tmp_path, [("测试甲有限公司", "李明", "13800000001", None),
                                             ("测试乙有限公司", "李明", "13800000002", None)])], master)
    assert sorted((c.company, c.phones[0]) for c in r.contacts) == [("测试乙有限公司", "+8613800000002"),
                                                                      ("测试甲有限公司", "+8613800000001")]
    crm = FakeCRM()
    excel_margince.contacts_commit(crm, r)
    assert sorted(c["phones"][0]["phone"] for c in crm.contacts.values()) == ["+8613800000001", "+8613800000002"]


def test_a_local_number_is_reported_not_guessed(tmp_path, master):
    r = contacts.read_paths([book(tmp_path, [("测试甲有限公司", "王五", "38938741 / 13800000001", None)])], master)
    [c] = r.contacts
    assert c.phones == ["+8613800000001"]
    assert [(i.field, i.value, i.reason) for i in r.review if i.reason == "phone_no_country"] == [("phone", "38938741", "phone_no_country")]
