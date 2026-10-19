"""What happens when the same Excel data is imported again, or changed between imports.
Invented data only. Tests named test_gap_* pin CURRENT behaviour the owner may want to change;
xfail(strict) tests are real bugs: they assert the wanted behaviour and fail today."""
import json
import os
import time
from datetime import datetime
from unittest import mock

import pytest

from excel import run as excel_run
from excel.run import read_paths
from excel_margince import accepted
from tests.excel_book import LUCAS_HEADER, WEEKLY_HEADER, write_book
from tests.fake_crm import FakeCRM
from visits import pipeline
from visits.match import match

NAME21 = "Template for Weekly Sales Reporting 2026_CW21_Rep.xlsx"
NAME22 = "Template for Weekly Sales Reporting 2026_CW22_Rep.xlsx"
MON, TUE, WED = datetime(2026, 5, 18), datetime(2026, 5, 19), datetime(2026, 5, 20)
NEXT_MON = datetime(2026, 5, 25)
ACME, BETA, GAMMA = "Acme Test Co.", "Beta Test GmbH", "Gamma Test Ltd."


@pytest.fixture
def tmp(tmp_path):
    decisions = tmp_path / "company_decisions.json"
    decisions.write_text(json.dumps({n: None for n in [
        "Acme Test Co.", "Beta Test Co.", "Beta Test GmbH", "Gamma Test Ltd.", "Gamma Test Limited",
        *[f"Company {i} Test Ltd." for i in range(10)]]}), encoding="utf-8")
    with mock.patch.dict(os.environ, {"COMPANY_DECISIONS_FILE": str(decisions)}):
        excel_run._master.cache_clear()
        yield tmp_path
        excel_run._master.cache_clear()


def row(day, cust, step="", purpose="", contact="Wang Test", phone="", email="", typ="exist", owner="Jason Zhao"):
    return dict(day=day, cust=cust, step=step, purpose=purpose, contact=contact, phone=phone,
                email=email, typ=typ, owner=owner)


def book(tmp, name, sheets, owner="Jason Zhao"):
    """sheets: {sheet name: [row(...)]}; one rep per workbook."""
    out = {}
    for sheet, rows in sheets.items():
        cells = dict(WEEKLY_HEADER)
        cells.update({"E4": "Owner：", "F4": owner})
        for i, r in enumerate(rows):
            n = 9 + i
            cells.update({f"B{n}": r["owner"], f"D{n}": r["day"], f"F{n}": r["cust"], f"H{n}": r["typ"],
                          f"I{n}": r["contact"], f"J{n}": r["phone"], f"L{n}": r["email"],
                          f"N{n}": r["purpose"], f"P{n}": r["step"]})
        out[sheet] = {"cells": cells}
    return write_book(str(tmp / name), out)


def load(crm, *paths, opts=None):
    parsed = accepted(read_paths(list(paths)))
    return pipeline.run(crm, parsed, match(parsed.visits), opts)


def visits(crm):
    return [a for a in crm.activities.values() if a["kind"] in ("meeting", "note") and "visit:" in a["source_id"]]


def tasks(crm):
    return [a for a in crm.activities.values() if a["kind"] == "task"]


def snapshot(crm, with_done=True):
    """The CRM state without its generated ids, comparable between two CRMs. The "From <file> > ..."
    provenance line of a visit is left out (it names whichever file wrote the visit first)."""
    acts = {a["source_id"]: (a["kind"], a["subject"],
                             "\n".join(l for l in a["body"].split("\n") if "›" not in l), a["occurred_at"],
                             a.get("is_done") if with_done else None) for a in crm.activities.values()}
    names = sorted(c["display_name"] for c in crm.companies.values())
    people = sorted((c["full_name"], tuple(sorted(p["phone"] for p in c.get("phones") or [])),
                     tuple(sorted(e["email"] for e in c.get("emails") or []))) for c in crm.contacts.values())
    return acts, names, people


def counts(crm):
    return (len(crm.companies), len(crm.contacts), len(crm.activities), len(crm.employment))


# ---------- same data again ----------

def test_same_file_twice_creates_nothing_the_second_time(tmp):
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer", "demo", phone="13600001111", email="w@test.example"),
                                    row(TUE, BETA)]})
    crm = FakeCRM()
    first = load(crm, p)
    before = snapshot(crm)
    second = load(crm, p)
    assert (first.activities_created, first.tasks_created, first.weeks_created) == (2, 1, 0)
    assert (second.activities_created, second.tasks_created, second.weeks_created, second.contacts_created) == (0, 0, 0, 0)
    assert (second.activities_idempotent, second.tasks_idempotent) == (2, 1)
    assert snapshot(crm) == before


def test_same_file_in_one_run_twice_via_two_paths(tmp):
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer")]})
    crm = FakeCRM()
    load(crm, p)
    n = counts(crm)
    load(crm, p, p)
    assert counts(crm) == n


def test_later_file_repeating_earlier_weeks_plus_new_ones(tmp):
    a = book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer"), row(TUE, BETA)]})
    b = book(tmp, NAME22, {"CW21": [row(MON, ACME, "Send offer"), row(TUE, BETA)],
                           "CW22": [row(NEXT_MON, GAMMA, "Call back")]})
    crm = FakeCRM()
    load(crm, a)
    r = load(crm, b)
    assert (r.activities_created, r.activities_idempotent, r.tasks_created) == (1, 2, 1)
    assert len(visits(crm)) == 3 and len(tasks(crm)) == 2
    assert len(crm.companies) == 3


def test_same_file_renamed_creates_nothing(tmp):
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer", phone="13600001111")]})
    crm = FakeCRM()
    load(crm, p)
    n = counts(crm)
    q = str(tmp / "Template for Weekly Sales Reporting 2026_CW21_Rep_copy final.xlsx")
    os.rename(p, q)
    r = load(crm, q)
    assert counts(crm) == n
    assert r.activities_created == 0 and r.tasks_created == 0


def test_two_copies_with_different_mtimes_in_one_run(tmp):
    a = book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer"), row(TUE, BETA)]})
    other = tmp / "copy"
    other.mkdir()
    b = book(other, NAME21, {"CW21": [row(MON, ACME, "Send offer"), row(TUE, BETA)]})
    os.utime(a, (time.time() - 86400 * 3, time.time() - 86400 * 3))
    os.utime(b, (time.time(), time.time()))
    crm = FakeCRM()
    load(crm, a, b)
    assert len(visits(crm)) == 2 and len(tasks(crm)) == 1
    assert len(crm.companies) == 2 and len(crm.contacts) == 2  # one Wang Test per company
    n = counts(crm)
    load(crm, b, a)
    assert counts(crm) == n


# ---------- text of a known visit / task changed ----------

def test_gap_changed_purpose_is_not_applied_to_the_existing_visit(tmp):
    """CURRENT BEHAVIOUR, may change: the visit is found by id and left as first written."""
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, purpose="Intro demo")]}))
    r = load(crm, book(tmp, "Template for Weekly Sales Reporting 2026_CW21_Rep_v2.xlsx",
                       {"CW21": [row(MON, ACME, purpose="Price negotiation")]}))
    (v,) = visits(crm)
    assert r.activities_created == 0 and r.activities_idempotent == 1
    assert "Intro demo" in v["body"] and "Price negotiation" not in v["body"]


def test_gap_changed_next_step_text_adds_a_second_task_and_keeps_the_old_one(tmp):
    """CURRENT BEHAVIOUR, may change: the task id hashes the text, so an edit is a new task
    and the old one stays open."""
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, step="Send offer")]}))
    r = load(crm, book(tmp, "Template for Weekly Sales Reporting 2026_CW21_Rep_v2.xlsx",
                       {"CW21": [row(MON, ACME, step="Send revised offer")]}))
    assert r.tasks_created == 1 and len(visits(crm)) == 1
    assert {t["subject"] for t in tasks(crm)} == {"Send offer", "Send revised offer"}


def test_task_text_change_in_case_or_spacing_only_is_the_same_task(tmp):
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, step="Send offer")]}))
    load(crm, book(tmp, "Template for Weekly Sales Reporting 2026_CW21_Rep_v2.xlsx",
                   {"CW21": [row(MON, ACME, step="send   OFFER")]}))
    assert len(tasks(crm)) == 1


# ---------- contacts ----------

def one_contact(crm):
    (c,) = crm.contacts.values()
    return c


def test_new_phone_for_existing_contact_is_added_and_old_kept_primary(tmp):
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, phone="13600001111")]}))
    load(crm, book(tmp, NAME22, {"CW22": [row(NEXT_MON, ACME, phone="13600002222")]}))
    ph = one_contact(crm)["phones"]
    assert [(p["phone"], p["is_primary"], p["position"]) for p in ph] == [
        ("+8613600001111", True, 0), ("+8613600002222", False, 1)]


def test_gap_changed_phone_is_added_not_replaced(tmp):
    """CURRENT BEHAVIOUR, may change: a corrected number leaves the wrong old one primary."""
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, phone="13600001111")]}))
    load(crm, book(tmp, "Template for Weekly Sales Reporting 2026_CW21_Rep_v2.xlsx",
                   {"CW21": [row(MON, ACME, phone="13600001112")]}))
    ph = one_contact(crm)["phones"]
    assert {p["phone"] for p in ph} == {"+8613600001111", "+8613600001112"}
    assert [p["phone"] for p in ph if p["is_primary"]] == ["+8613600001111"]


def test_same_phone_written_differently_is_not_added_twice(tmp):
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, phone="13600001111")]}))
    load(crm, book(tmp, NAME22, {"CW22": [row(NEXT_MON, ACME, phone="136 0000 1111")]}))
    assert len(one_contact(crm)["phones"]) == 1


def test_new_email_for_existing_contact_is_added(tmp):
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, email="w@test.example")]}))
    load(crm, book(tmp, NAME22, {"CW22": [row(NEXT_MON, ACME, email="W@Test.example; w2@test.example")]}))
    em = one_contact(crm)["emails"]
    assert [(e["email"].lower(), e["is_primary"]) for e in em] == [("w@test.example", True), ("w2@test.example", False)]


def test_second_commit_with_an_address_another_contact_holds_does_not_fail_or_write(tmp):
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME, contact="Wang Test", email="shared@test.example"),
                                    row(TUE, BETA, contact="Li Test", email="shared@test.example")]})
    crm = FakeCRM()
    load(crm, p)
    li = [c for c in crm.contacts.values() if c["full_name"] == "Li Test"][0]
    assert not li["emails"]  # made without the address Wang holds
    before = snapshot(crm)
    second = load(crm, p)
    assert (second.contacts_created, second.contacts_updated, second.activities_created) == (0, 0, 0)
    assert snapshot(crm) == before


def test_contact_found_in_crm_with_different_letter_case(tmp):
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, phone="13600001111")]}))
    one_contact(crm)["full_name"] = "WANG TEST"
    load(crm, book(tmp, NAME22, {"CW22": [row(NEXT_MON, ACME, phone="13600001111")]}))
    assert len(crm.contacts) == 1


def test_contact_found_in_crm_with_extra_inner_space(tmp):
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME)]}))
    one_contact(crm)["full_name"] = "Wang  Test"
    load(crm, book(tmp, NAME22, {"CW22": [row(NEXT_MON, ACME)]}))
    assert len(crm.contacts) == 1


def test_contact_name_with_extra_space_in_the_sheet_is_the_same_contact(tmp):
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME)]}))
    load(crm, book(tmp, NAME22, {"CW22": [row(NEXT_MON, ACME, contact=" wang   TEST ")]}))
    assert len(crm.contacts) == 1


# ---------- companies changed in the CRM between runs ----------

@pytest.mark.xfail(strict=True, reason="BUG: a company renamed in the CRM is not found by name, so a re-import "
                                       "creates a second company (and a second contact when there is no email)")
def test_company_renamed_in_crm_is_not_created_again(tmp):
    crm = FakeCRM()
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer")]})
    load(crm, p)
    (cid,) = crm.companies
    crm.companies[cid]["display_name"] = "Acme Renamed AG"
    load(crm, p)
    assert len(crm.companies) == 1
    assert len(crm.contacts) == 1


def test_company_renamed_in_crm_does_not_duplicate_the_visits(tmp):
    crm = FakeCRM()
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer")]})
    load(crm, p)
    (cid,) = crm.companies
    crm.companies[cid]["display_name"] = "Acme Renamed AG"
    load(crm, p)
    assert len(visits(crm)) == 1 and len(tasks(crm)) == 1


def test_company_deleted_in_crm_is_recreated_and_old_visit_stays(tmp):
    crm = FakeCRM()
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer", email="w@test.example")]})
    load(crm, p)
    (cid,) = crm.companies
    del crm.companies[cid]
    crm.employment[:] = [e for e in crm.employment if e[0] != cid]
    load(crm, p)
    assert [c["display_name"] for c in crm.companies.values()] == [ACME]
    assert list(crm.companies) != [cid]
    assert len(visits(crm)) == 1 and len(tasks(crm)) == 1
    assert len(crm.contacts) == 1  # found again through its email


@pytest.mark.xfail(strict=True, reason="BUG: after the company was deleted in the CRM, a contact without email "
                                       "is not recognised again and is created a second time")
def test_company_deleted_in_crm_does_not_duplicate_contact_without_email(tmp):
    crm = FakeCRM()
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME)]})
    load(crm, p)
    (cid,) = crm.companies
    del crm.companies[cid]
    crm.employment[:] = [e for e in crm.employment if e[0] != cid]
    load(crm, p)
    assert len(crm.contacts) == 1


def test_gap_visit_of_deleted_company_is_not_relinked(tmp):
    """CURRENT BEHAVIOUR, may change: the existing visit keeps pointing at the contact written the first time."""
    crm = FakeCRM()
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME)]})
    load(crm, p)
    (v,) = visits(crm)
    old_links = list(v["links"])
    (cid,) = crm.companies
    del crm.companies[cid]
    crm.employment[:] = []
    load(crm, p)
    (v,) = visits(crm)
    assert v["links"] == old_links


# ---------- the sheet changed between imports ----------

def test_gap_corrected_visit_date_leaves_the_old_visit_and_task_behind(tmp):
    """CURRENT BEHAVIOUR, may change: the date is part of the id, so the corrected row is a new visit
    and the wrong-day visit and its task stay in the CRM."""
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer")]}))
    r = load(crm, book(tmp, "Template for Weekly Sales Reporting 2026_CW21_Rep_v2.xlsx",
                       {"CW21": [row(WED, ACME, "Send offer")]}))
    assert (r.activities_created, r.tasks_created) == (1, 1)
    assert sorted(v["occurred_at"][:10] for v in visits(crm) if "Weekly" not in v["subject"]) == \
        ["2026-05-18", "2026-05-20"]
    assert len(tasks(crm)) == 2


def test_gap_row_deleted_from_sheet_keeps_old_visit_and_task(tmp):
    """CURRENT BEHAVIOUR, may change: nothing is retracted."""
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer"), row(TUE, BETA, "Call")]}))
    r = load(crm, book(tmp, "Template for Weekly Sales Reporting 2026_CW21_Rep_v2.xlsx",
                       {"CW21": [row(MON, ACME, "Send offer")]}))
    assert (r.activities_created, r.tasks_created) == (0, 0)
    assert len(visits(crm)) == 2
    assert any("Beta" in v["subject"] for v in visits(crm))
    assert len(tasks(crm)) == 2


def test_gap_customer_spelling_corrected_makes_a_second_company_contact_and_visit(tmp):
    """CURRENT BEHAVIOUR, may change: a corrected customer name is a different customer; the old
    company, its contact and its visit stay next to the new ones."""
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, GAMMA, "Call back")]}))
    r = load(crm, book(tmp, "Template for Weekly Sales Reporting 2026_CW21_Rep_v2.xlsx",
                       {"CW21": [row(MON, "Gamma Test Limited", "Call back")]}))
    assert r.activities_created == 1 and r.tasks_created == 1
    assert sorted(c["display_name"] for c in crm.companies.values()) == ["Gamma Test Limited", GAMMA]
    assert len(visits(crm)) == 2 and len(crm.contacts) == 2


def test_customer_spelling_differing_only_in_case_and_spaces_is_the_same_visit(tmp):
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer")]}))
    r = load(crm, book(tmp, "Template for Weekly Sales Reporting 2026_CW21_Rep_v2.xlsx",
                       {"CW21": [row(MON, "  ACME   test co. ", "Send offer")]}))
    assert r.activities_created == 0 and len(crm.companies) == 1 and len(visits(crm)) == 1


def test_two_reps_visiting_the_same_company_the_same_day(tmp):
    crm = FakeCRM()
    a = book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer", phone="13600001111")]})
    other = tmp / "rep2"
    other.mkdir()
    b = book(other, "Template for Weekly Sales Reporting 2026_CW21_Devin.xlsx", {"CW21": [row(MON, ACME, "Visit again", owner="Devin Guo", contact="Wang Test",
                                          phone="13600001111")]}, owner="Devin Guo")
    load(crm, a, b)
    assert len(visits(crm)) == 2 and len(crm.companies) == 1 and len(crm.contacts) == 1
    (co,) = crm.companies.values()
    assert {v["source_author_name"] for v in visits(crm)} == {"Jason Zhao", "Devin Guo"}
    n = counts(crm)
    load(crm, b, a)
    assert counts(crm) == n


# ---------- ordering ----------

def two_files(tmp):
    a = book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer", phone="13600001111"), row(TUE, BETA)]})
    b = book(tmp, NAME22, {"CW21": [row(MON, ACME, "Send offer", phone="13600001111"), row(TUE, BETA)],
                           "CW22": [row(NEXT_MON, GAMMA, "Call back"), row(NEXT_MON, ACME, phone="13600002222")]})
    return a, b


def test_order_of_files_gives_the_same_crm_state(tmp):
    a, b = two_files(tmp)
    ab, ba = FakeCRM(), FakeCRM()
    load(ab, a)
    load(ab, b)
    load(ba, b)
    load(ba, a)
    assert snapshot(ab, with_done=False) == snapshot(ba, with_done=False)


def test_gap_visit_provenance_names_the_file_that_came_first(tmp):
    """CURRENT BEHAVIOUR, may change: a visit repeated in a later file keeps the first file's name
    in its text; the same data in the opposite order names the other file."""
    a, b = two_files(tmp)
    ab, ba = FakeCRM(), FakeCRM()
    load(ab, a)
    load(ab, b)
    load(ba, b)
    load(ba, a)
    src = lambda crm: {v["source_id"]: v["body"].split("\n")[-1].split(" — ")[1].split(" › ")[0] for v in visits(crm)
                       if "2026-05-18:acme" in v["source_id"]}
    assert list(src(ab).values()) == [NAME21] and list(src(ba).values()) == [NAME22]


def test_order_inside_one_run_gives_the_same_crm_state(tmp):
    a, b = two_files(tmp)
    one, two = FakeCRM(), FakeCRM()
    load(one, a, b)
    load(two, b, a)
    assert snapshot(one, with_done=False) == snapshot(two, with_done=False)


def test_company_fields_keep_the_newest_visit_whatever_the_order(tmp):
    a, b = two_files(tmp)
    states = []
    for order in ((a, b), (b, a)):
        crm = FakeCRM()
        for p in order:
            load(crm, p)
        acme = next(c for c in crm.companies.values() if c["display_name"] == ACME)
        states.append({k: v for k, v in acme.items() if k.startswith("cf_")})
    assert states[0] == states[1]
    assert states[0]["cf_last_visit"] == "2026-05-25"


def test_gap_task_closing_depends_on_how_the_files_are_batched(tmp):
    """CURRENT BEHAVIOUR, may change: the default cutoff is the newest week of the RUN (only lucas-layout
    sheets have a week). A CW10 task stays open when its file is imported alone, closed when a CW11 file
    is in the same run."""
    a = lucas_book(tmp, "Weekly Sales Planning and Reporting 2026_CW10-Lucas.xlsx", datetime(2026, 3, 2), ACME, "Old")
    b = lucas_book(tmp, "Weekly Sales Planning and Reporting 2026_CW11-Lucas.xlsx", datetime(2026, 3, 9), BETA, "New")
    apart, together = FakeCRM(), FakeCRM()
    load(apart, a)
    load(apart, b)
    load(together, a, b)
    done = lambda crm: {t["subject"]: t.get("is_done") for t in tasks(crm)}
    assert done(apart) == {"Old": False, "New": False}
    assert done(together) == {"Old": True, "New": False}


def test_gap_weekly_layout_has_no_week_so_no_default_cutoff_and_old_tasks_stay_open(tmp):
    """CURRENT BEHAVIOUR, may change: without a week the cutoff is None and nothing is written done."""
    crm = FakeCRM()
    load(crm, book(tmp, NAME22, {"CW21": [row(MON, ACME, "Old thing")], "CW22": [row(NEXT_MON, ACME, "New thing")]}))
    assert {t["subject"]: t["is_done"] for t in tasks(crm)} == {"Old thing": False, "New thing": False}


def test_tasks_before_cutoff_are_written_done_and_stay_so(tmp):
    p = book(tmp, NAME22, {"CW21": [row(MON, ACME, "Old thing")], "CW22": [row(NEXT_MON, ACME, "New thing")]})
    crm = FakeCRM()
    opts = pipeline.Options(close_tasks_due_before="2026-05-25")
    load(crm, p, opts=opts)
    load(crm, p, opts=opts)
    assert {t["subject"]: t["is_done"] for t in tasks(crm)} == {"Old thing": True, "New thing": False}


# ---------- partial failure ----------

class Boom(Exception):
    pass


class FailingCRM(FakeCRM):
    """Raises on the Nth call of one method (1-based), once."""
    def __init__(self, method, nth):
        super().__init__()
        self.method, self.nth, self.calls, self.armed = method, nth, 0, True
        orig = getattr(FakeCRM, method)

        def wrapper(*a, **k):
            self.calls += 1
            if self.armed and self.calls == self.nth:
                raise Boom(method)
            return orig(self, *a, **k)
        setattr(self, method, wrapper)


def full_book(tmp):
    return book(tmp, NAME21, {"CW21": [row(MON, ACME, "Send offer", phone="13600001111", email="w@test.example"),
                                       row(TUE, BETA, "Call back", contact="Li Test"),
                                       row(WED, GAMMA, contact="Zhao Test")]})


def test_failure_free_baseline_counts(tmp):
    crm = FakeCRM()
    load(crm, full_book(tmp))
    assert counts(crm) == (3, 3, 3 + 2, 3)


@pytest.mark.parametrize("method,nth", [("log_activity", 1), ("log_activity", 3), ("log_activity", 5),
                                        ("create_contact", 2),
                                        ("run_csv_import", 1), ("create_custom_field", 3)])
def test_run_that_dies_midway_is_finished_by_a_clean_rerun_without_duplicates(tmp, method, nth):
    p = full_book(tmp)
    clean = FakeCRM()
    load(clean, p)
    crm = FailingCRM(method, nth)
    with pytest.raises(Boom):
        load(crm, p)
    crm.armed = False
    load(crm, p)
    assert snapshot(crm, with_done=False)[1:] == snapshot(clean, with_done=False)[1:]
    assert {k: v[:4] for k, v in snapshot(crm)[0].items()} == {k: v[:4] for k, v in snapshot(clean)[0].items()}
    assert len(crm.employment) == len(set(crm.employment))
    third = load(crm, p)
    assert (third.activities_created, third.tasks_created, third.contacts_created) == (0, 0, 0)


@pytest.mark.xfail(strict=True, reason="BUG: a contact created but not yet linked to its company when the run died is "
                                       "not found again (lookup is by employment), so the re-run creates it a second time")
def test_contact_created_but_link_failed_is_not_created_twice_on_rerun(tmp):
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME)]})
    crm = FailingCRM("link_employment", 1)
    with pytest.raises(Boom):
        load(crm, p)
    crm.armed = False
    load(crm, p)
    assert len(crm.contacts) == 1 and len(crm.employment) == 1


def test_company_created_but_run_died_before_contacts_leaves_no_duplicate_company(tmp):
    p = full_book(tmp)
    crm = FailingCRM("create_contact", 1)
    with pytest.raises(Boom):
        load(crm, p)
    assert len(crm.companies) == 3 and not crm.contacts
    crm.armed = False
    load(crm, p)
    assert len(crm.companies) == 3 and len(crm.contacts) == 3


@pytest.mark.xfail(strict=True, reason="BUG: a task logged but not yet closed when the run died is found by "
                                       "its id on the re-run and never closed, so an old task stays open")
def test_task_logged_but_close_failed_is_closed_by_the_rerun(tmp):
    p = book(tmp, NAME22, {"CW21": [row(MON, ACME, "Old thing")], "CW22": [row(NEXT_MON, ACME)]})
    p = p

    class CloseFails(FakeCRM):
        armed = True

        def patch(self, path, body, **kw):
            if self.armed and body == {"is_done": True}:
                raise Boom("close")
            return super().patch(path, body, **kw)
    crm = CloseFails()
    opts = pipeline.Options(close_tasks_due_before="2026-05-25")
    with pytest.raises(Boom):
        load(crm, p, opts=opts)
    crm.armed = False
    load(crm, p, opts=opts)
    assert {t["subject"]: t["is_done"] for t in tasks(crm)} == {"Old thing": True}


# ---------- timezone ----------

@pytest.mark.parametrize("stamp", [datetime(2026, 5, 18, 0, 0, 1), datetime(2026, 5, 18, 12, 0),
                                   datetime(2026, 5, 18, 23, 59, 59)])
def test_date_cell_near_midnight_stays_on_its_day(tmp, stamp):
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(stamp, ACME, "Send offer")]}))
    (v,) = visits(crm)
    assert v["source_id"].split(":")[2] == "2026-05-18"
    assert v["raw"]["visit_list_row"]["date_iso"] == "2026-05-18"
    from datetime import timezone, timedelta
    local = datetime.fromisoformat(v["occurred_at"]).astimezone(timezone(timedelta(hours=8)))
    assert local.date().isoformat() == "2026-05-18"


def test_same_day_with_a_different_time_in_the_cell_is_the_same_visit(tmp):
    crm = FakeCRM()
    load(crm, book(tmp, NAME21, {"CW21": [row(datetime(2026, 5, 18, 0, 0, 1), ACME, "Send offer")]}))
    r = load(crm, book(tmp, "Template for Weekly Sales Reporting 2026_CW21_Rep_v2.xlsx",
                       {"CW21": [row(datetime(2026, 5, 18, 23, 59, 59), ACME, "Send offer")]}))
    assert r.activities_created == 0 and len(visits(crm)) == 1


# ---------- many files ----------

def test_ten_files_in_one_run_and_again(tmp):
    paths = []
    for i in range(10):
        owner = ["Jason Zhao", "Devin Guo", "Mark Wang"][i % 3]
        day = datetime(2026, 5, 18 + (i % 5))
        paths.append(book(tmp, f"Template for Weekly Sales Reporting 2026_CW21_R{i}.xlsx", {"CW21": [row(day, f"Company {i} Test Ltd.", f"Step {i}", owner=owner,
                                                    contact=f"Person {i}")]}, owner=owner))
    # the first and the last file are the same visit twice
    paths.append(paths[0])
    crm = FakeCRM()
    r = load(crm, *paths)
    assert len(visits(crm)) == 10 and len(tasks(crm)) == 10
    assert len(crm.companies) == 10 and len(crm.contacts) == 10
    n = counts(crm)
    again = load(crm, *reversed(paths))
    assert counts(crm) == n and again.activities_created == 0 and again.tasks_created == 0


# ---------- week notes (lucas layout is the only one that makes a week) ----------

def lucas_book(tmp, name, day, customer, step="", plan="Alpha trip"):
    cells = dict(LUCAS_HEADER)
    cells.update({"C4": "Owner：", "D4": "Lucas Guo", "C6": "Monday", "D6": plan})
    cells.update({"C15": day, "E15": customer, "G15": "exist", "H15": "Wu Test", "L15": "Phone call", "M15": step})
    return write_book(str(tmp / name), {"W01": {"cells": cells}})


def weeks(crm):
    return [a for a in crm.activities.values() if a["source_id"].startswith("week:")]


def test_week_note_is_written_once(tmp):
    p = lucas_book(tmp, "Weekly Sales Planning and Reporting 2026_CW10-Lucas.xlsx", datetime(2026, 3, 2), ACME)
    crm = FakeCRM()
    assert load(crm, p).weeks_created == 1
    r = load(crm, p)
    assert (r.weeks_created, r.weeks_idempotent) == (0, 1) and len(weeks(crm)) == 1


def test_gap_changed_key_plan_does_not_update_the_week_note(tmp):
    """CURRENT BEHAVIOUR, may change: the week note keeps the plan text of the first import."""
    crm = FakeCRM()
    load(crm, lucas_book(tmp, "Weekly Sales Planning and Reporting 2026_CW10-Lucas.xlsx", datetime(2026, 3, 2), ACME))
    load(crm, lucas_book(tmp, "Weekly Sales Planning and Reporting 2026_CW10-Lucas2.xlsx", datetime(2026, 3, 2), ACME,
                         plan="Beta trip"))
    (w,) = weeks(crm)
    assert "Alpha trip" in w["body"] and "Beta trip" not in w["body"]


# ---------- two files, one visit ----------

def test_two_files_with_one_visit_the_newer_file_wins_whatever_the_order(tmp):
    old = book(tmp, "Template for Weekly Sales Reporting 2026_CW21_Old.xlsx", {"CW21": [row(MON, ACME, purpose="Old text")]})
    new = book(tmp, "Template for Weekly Sales Reporting 2026_CW21_New.xlsx", {"CW21": [row(MON, ACME, purpose="New text")]})
    t = time.time()
    os.utime(old, (t - 5 * 86400, t - 5 * 86400))
    os.utime(new, (t, t))
    for order in ((old, new), (new, old)):
        crm = FakeCRM()
        load(crm, *order)
        (v,) = visits(crm)
        assert "New text" in v["body"], order


def test_gap_two_different_files_with_one_name_the_second_is_skipped(tmp):
    """CURRENT BEHAVIOUR, may change: the file name is part of the id, so a second, different file
    with the same name (another folder) is left out of the run."""
    a = book(tmp, NAME21, {"CW21": [row(MON, ACME)]})
    d = tmp / "other"
    d.mkdir()
    b = book(d, NAME21, {"CW21": [row(TUE, BETA)]})
    result = read_paths([a, b])
    assert len(result.parsed.visits) == 1
    crm = FakeCRM()
    load(crm, a, b)
    assert len(visits(crm)) == 1


def test_second_commit_does_not_fail_when_a_contact_was_made_without_a_taken_email(tmp):
    """Two different people write the same address (a shared mailbox). The second is created without
    it; every later run must not try to add it again (the CRM answers PATCH with 409 duplicate_email)."""
    from tests.fake_site_crm import SiteCRM
    p = book(tmp, NAME21, {"CW21": [row(MON, ACME, contact="Wang Test", email="shared@test.example"),
                                    row(TUE, BETA, contact="Li Test", email="shared@test.example")]})
    crm = SiteCRM()
    load(crm, p)
    before = len(crm.writes)
    second = load(crm, p)
    assert len(crm.writes) == before
    assert second.contacts_created == 0
