"""The contact people of the ERP customer lists (联系人, 电话, 手机, Email地址), on invented workbooks:
read into the account by its customer code, several numbers per person, nothing guessed, nothing lost."""
import pytest

from common.client import ApiError
from excel import accounts, accounts_margince as am
from excel.companies import Master
from tests.test_excel_accounts import FakeCRM, erp_book
from tests.test_excel_accounts import erp_row as _erp_row

NO_CONTACT = {"联系人": None, "电话": None, "手机": None, "Email地址": None}


def erp_row(code, name, **over):
    """An ERP row with the contact columns empty unless the test gives them."""
    return _erp_row(code, name, **{**NO_CONTACT, **over})


def problem(status, code):
    import json
    import requests
    resp = requests.Response()
    resp.status_code = status
    resp._content = json.dumps({"title": "boom", "code": code}).encode()
    resp.request = requests.Request("POST", "http://fake/v1/contacts").prepare()
    return ApiError(resp)


@pytest.fixture(autouse=True)
def _decisions(tmp_path, monkeypatch):
    monkeypatch.setenv("COMPANY_DECISIONS_FILE", str(tmp_path / "company_decisions.json"))


def read(tmp_path, *rows):
    return accounts.read_paths([erp_book(tmp_path, list(rows))], Master.from_records([]))


def person(acc, name):
    return next(p for p in acc.contacts if p.name == name)


def test_numbers_and_address_are_kept_on_the_person_of_the_account_with_that_code(tmp_path):
    r = read(tmp_path, erp_row("CR100001", "测试甲机电有限公司", 联系人="王五", 电话="0351-6262916 / 0351-6262917",
                               手机="13834581008", Email地址="wang@example.invalid"))
    [a] = r.accounts
    p = person(a, "王五")
    assert a.code == "CR100001" and p.phones == ["+8613834581008", "+863516262916", "+863516262917"]
    assert p.emails == ["wang@example.invalid"] and p.other == []


def test_same_name_at_two_companies_is_two_people(tmp_path):
    r = read(tmp_path, erp_row("CR100001", "测试甲机电有限公司", 联系人="李明", 手机="13800000001"),
             erp_row("CR100002", "测试乙机床有限公司", 联系人="李明", 手机="13800000002"))
    by = {a.code: a for a in r.accounts}
    assert person(by["CR100001"], "李明").phones == ["+8613800000001"]
    assert person(by["CR100002"], "李明").phones == ["+8613800000002"]
    crm = FakeCRM()
    am.commit(crm, r)
    assert sorted(p["phones"][0]["phone"] for p in crm.contacts.values()) == ["+8613800000001", "+8613800000002"]


def test_the_same_person_in_two_rows_gets_all_his_numbers_once(tmp_path):
    r = read(tmp_path, erp_row("CR100001", "测试甲机电有限公司", 联系人="王五", 手机="13834581008"),
             erp_row("CR100001", "测试甲机电有限公司", 联系人="王五 ", 手机="13834581008", 电话="021-55550001"))
    [a] = r.accounts
    assert len(a.contacts) == 1 and a.contacts[0].phones == ["+8613834581008", "+862155550001"]


def test_what_cannot_be_placed_is_not_guessed_and_is_not_lost(tmp_path):
    r = read(tmp_path, erp_row("CR100001", "测试甲机电有限公司", 联系人="王五", 电话="25190887 / 010-62984726-1596 / 稍后再打",
                               手机="1365917891", Email地址="wang@example"))
    [a] = r.accounts
    p = a.contacts[0]
    assert p.phones == ["+861062984726"]  # the extension is kept as text, a local number is not given an area code
    assert p.other == ["1365917891", "25190887", "010-62984726-1596", "稍后再打", "wang@example"]  # the mobile cell first
    [note] = [n for n in a.notes if n.startswith("Contact 王五")]
    assert "25190887" in note and "1596" in note and "稍后再打" in note and "wang@example" in note


def test_fullwidth_and_long_dashes_and_400_lines(tmp_path):
    r = read(tmp_path, erp_row("CR100001", "测试甲机电有限公司", 联系人="王五", 电话="010—84254081", 手机="+86 (411) 8670 3123"),
             erp_row("CR100002", "测试乙机床有限公司", 联系人="赵六", 电话="0791－83962762 / 400-6789-532"))
    by = {a.code: a for a in r.accounts}
    assert by["CR100001"].contacts[0].phones == ["+8641186703123", "+861084254081"]
    assert by["CR100002"].contacts[0].phones == ["+867918396 2762".replace(" ", ""), "+864006789532"]


def test_contact_details_without_a_name_become_a_note_not_a_contact(tmp_path):
    r = read(tmp_path, erp_row("CR100001", "测试甲机电有限公司", 电话="010-63185480", 手机="17744592679",
                               Email地址="info@example.invalid"))
    [a] = r.accounts
    assert a.contacts == []
    assert a.notes == ["Contact details of the customer list that name no person: phone 010-63185480; "
                       "mobile 17744592679; e-mail info@example.invalid"]


def test_a_unlabelled_column_inside_the_table_is_a_sales_note(tmp_path):
    from tests.test_excel_accounts import ERP_HEADER, book
    head = list(ERP_HEADER)
    head.insert(head.index("电话"), None)  # no header text over this column
    row = erp_row("CR100001", "测试甲机电有限公司", 联系人="王五")
    row.insert(ERP_HEADER.index("电话"), "加微信")
    p = book(tmp_path, "erp2.xlsx", {"North": [head, row]})
    [a] = accounts.read_paths([p], Master.from_records([])).accounts
    assert "Sales note (unlabelled column of the ERP list): 加微信" in a.notes


def test_commit_writes_every_number_and_address_and_is_idempotent(tmp_path):
    r = read(tmp_path, erp_row("CR100001", "测试甲机电有限公司", 联系人="王五", 电话="021-55550001", 手机="13834581008",
                               Email地址="wang@example.invalid"))
    crm = FakeCRM()
    am.commit(crm, r)
    [c] = crm.contacts.values()
    assert [(p["phone"], p["phone_type"], p["is_primary"]) for p in c["phones"]] == [
        ("+8613834581008", "mobile", True), ("+862155550001", "work", False)]
    assert c["emails"] == [{"email": "wang@example.invalid", "email_type": "work", "is_primary": True, "position": 0}]
    again = am.commit(crm, r)
    assert again["contacts_created"] == 0 and again["contacts_present"] == 1 and len(crm.contacts) == 1


def test_an_address_the_crm_holds_on_another_contact_is_not_taken_from_it(tmp_path):
    r = read(tmp_path, erp_row("CR100001", "测试甲机电有限公司", 联系人="王五", 手机="13834581008", Email地址="wang@example.invalid"))

    class Clash(FakeCRM):
        def create_contact(self, body):
            if body.get("emails"):
                raise problem(409, "duplicate_email")
            return super().create_contact(body)
    crm = Clash()
    report = am.commit(crm, r)
    assert report["failed"] == [] and report["contacts_created"] == 1
    assert report["emails_held_elsewhere"] == [{"id": "account:CR100001", "contact": "王五", "emails": ["wang@example.invalid"]}]
    [c] = crm.contacts.values()
    assert "emails" not in c and c["phones"][0]["phone"] == "+8613834581008"
