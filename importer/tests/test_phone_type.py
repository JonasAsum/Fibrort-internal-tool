"""A phone number is "mobile" only when the number itself is a Chinese mobile (+86 1[3-9]...); landlines, 400/800 lines
and foreign numbers are "work", whichever Excel column they stood in."""
from types import SimpleNamespace as NS

from excel.accounts_margince import contact_body
from excel.split import phone_type
from visits.pipeline import contact_phones


def test_phone_type_by_number():
    assert phone_type("+8613701217825") == "mobile"
    assert phone_type("+861052649208") == "work"      # Beijing 010 landline: starts +861 but is not a mobile
    assert phone_type("+862039206114") == "work"      # Guangzhou 020
    assert phone_type("+864008001234") == "work"
    assert phone_type("+4971234567") == "work"
    assert phone_type("+861234567890") == "work"      # 12x is no mobile prefix


def test_accounts_path_labels_beijing_landline_work():
    body = contact_body(NS(name="x", phones=["+8613701217825", "+861052649208"], emails=[]))
    assert [(p["phone"], p["phone_type"]) for p in body["phones"]] == [("+8613701217825", "mobile"), ("+861052649208", "work")]


def test_visits_path_ignores_the_column():
    c = NS(phones=["010-52649208", "13701217825"], office_phones=["13911496998"])
    got = {p["phone"]: p["phone_type"] for p in contact_phones(c)}
    assert got == {"+861052649208": "work", "+8613701217825": "mobile", "+8613911496998": "mobile"}
