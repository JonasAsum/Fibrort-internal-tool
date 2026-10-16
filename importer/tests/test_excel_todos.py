"""To-do and sales-topic importer (folder 07): layouts, templates, status / priority / due variants,
the IT-list switch, company links, copies and ids, and the Margince adapter against tests/fake_crm.py.
Every value here is invented; the real-folder test asserts counts only."""
import datetime as dt
import json
import os
import shutil

import pytest

from common.client import ApiError
from excel import todos, todos_margince
from excel.companies import Master
from excel.todos import map_status
from tests.conftest import SAMPLE_DIR, make_book
from tests.fake_crm import FakeCRM, _Conflict

HEAD = ["No.", "Categorie", "Date", "Topics", "To Do / Next Steps", "Responsible", "Due Date", "Status", "Prio",
        "Remarks", "Documents or Websites (Hyperlink)"]
D = dt.datetime(2026, 4, 7)
MASTER = [{"source": "master", "code": "KD00001", "cn": "宇宙星辰科技有限公司", "en": "Cosmos Stellar Co., Ltd.", "short": ""},
          {"source": "master", "code": "KD00002", "cn": "", "en": "Nova Quantum GmbH", "short": ""}]


@pytest.fixture
def master():
    return Master.from_records(MASTER)


def cells(rows, head=HEAD, top=3, title="TOPIC - Invented Call", owner="Project Responsible:", ncols=None):
    """Projekt-OPL sheet: merged title rows, header on `top`, data below it."""
    out = {"A1": title, "A2": owner}
    for c, label in enumerate(head, start=1):
        out[f"{_col(c)}{top}"] = label
    for r, row in enumerate(rows, start=top + 1):
        for c, value in enumerate(row, start=1):
            if value is not None:
                out[f"{_col(c)}{r}"] = value
    return out


def _col(c):
    return "ABCDEFGHIJKLMNOP"[c - 1]


def row(topic="Weekly reports", text="send the weekly report", who="Ethan", due=D, status="WIP", prio=1,
        remarks=None, link=None, date=D, cat=None, no=1):
    return [no, cat, date, topic, text, who, due, status, prio, remarks, link]


def lst(tmp_path, rows, name="list.xlsx", head=HEAD, sheets=None, **kw):
    sheets = sheets or {"Projekt-OPL": {"cells": cells(rows, head, **kw), "merges": ["A1:K1", "A2:K2"]}}
    return make_book(tmp_path / name, sheets)


def read(paths, master=None, **kw):
    paths = [paths] if isinstance(paths, str) else paths
    return todos.read_paths(paths, master=master or Master.from_records([]), **kw)


def reasons(result):
    return sorted(i.reason for i in result.review)


# ---------- the to-do layout ----------

def test_todo_layout_reads_a_row(tmp_path):
    r = read(lst(tmp_path, [row(remarks="in progress", link="https://x.example/f", cat="Reporting")]))
    t, = r.todos
    assert (t.subject, t.text, t.owner, t.owner_raw) == ("Weekly reports", "send the weekly report", "Ethan Lv", "Ethan")
    assert (t.state, t.done, t.prio, t.due_iso, t.date_iso) == ("wip", False, "1", "2026-04-07", "2026-04-07")
    assert (t.remarks, t.link, t.category, t.layout, t.kind, t.row) == (
        "in progress", "https://x.example/f", "Reporting", "todo", "todo", 4)
    assert t.tags == [] and t.source_id.startswith("todo:")
    assert reasons(r) == []


def test_done_open_and_unknown_status(tmp_path):
    r = read(lst(tmp_path, [row("A", status="Done"), row("B", status="Open"), row("C", status=None),
                             row("D", status="Sometime"), row("E", status="已完成"), row("F", status="Cancelled")]))
    by = {t.subject: t for t in r.todos}
    assert [by[k].state for k in "ABCDEF"] == ["done", "open", "open", "open", "done", "cancelled"]
    assert [by[k].done for k in "ABCDEF"] == [True, False, False, False, True, True]
    assert reasons(r) == ["unknown_status"]  # only the word nobody knows; an empty cell is just open


@pytest.mark.parametrize("raw,state", [
    ("Done", "done"), ("DONE ", "done"), ("Done (30.03.)", "done"), ("closed", "done"), ("erledigt", "done"),
    ("完成", "done"), ("已完成", "done"), ("WIP", "wip"), ("In progress", "wip"), ("进行中", "wip"),
    ("Open", "open"), ("to do", "open"), ("未完成", "open"), ("未开始", "open"), ("offen", "open"), ("", "open"),
    ("On hold", "waiting"), ("等待", "waiting"), ("canceled", "cancelled"), ("取消", "cancelled")])
def test_map_status(raw, state):
    assert map_status(raw) == (state, True)


def test_chinese_headers_in_any_order_and_a_lower_header_row(tmp_path):
    head = ["状态", "优先级", "负责人", "截止日期", "待办事项", "主题", "备注"]
    rows = [["已完成", 2, "Ethan", D, "整理客户清单", "客户清单", "已发"],
            ["进行中", 1, "Ethan", "TBC", "准备展会", "展会", None]]
    r = read(lst(tmp_path, rows, head=head, top=6, title="待办清单", owner=""))
    a, b = r.todos
    assert (a.subject, a.text, a.state, a.prio, a.due_iso, a.remarks, a.owner) == (
        "客户清单", "整理客户清单", "done", "2", "2026-04-07", "已发", "Ethan Lv")
    assert (b.state, b.due_iso, b.due_raw) == ("wip", "", "TBC")


def test_a_todo_without_topic_uses_its_first_line_and_a_topic_without_todo_stands_alone(tmp_path):
    r = read(lst(tmp_path, [row(topic=None, text="first line\nsecond line"), row(topic="Only a topic", text=None)]))
    a, b = r.todos
    assert (a.subject, a.text) == ("first line", "first line\nsecond line")
    assert (b.subject, b.text) == ("Only a topic", "")


def test_sheet_owner_label_fills_rows_without_a_responsible(tmp_path):
    r = read(lst(tmp_path, [row(who=None)], owner="Project Responsible: Mark Wang"))
    assert r.todos[0].owner == "Mark Wang"
    r = read(lst(tmp_path, [row(who=None)], owner="Project Responsible: XXXX", name="b.xlsx"))
    assert r.todos[0].owner == "" and r.todos[0].owner_raw == ""


def test_owner_names_that_are_not_a_single_rep_stay_unassigned(tmp_path):
    r = read(lst(tmp_path, [row("A", who="Chenley"), row("B", who="Ethan, Mark"), row("C", who="Mark Wang")]))
    by = {t.subject: t for t in r.todos}
    assert (by["A"].owner, by["A"].owner_raw, by["B"].owner, by["C"].owner) == ("", "Chenley", "", "Mark Wang")
    assert reasons(r) == ["owner_unknown", "owner_unknown"]


def test_section_heading_rows_are_not_tasks_and_name_the_category(tmp_path):
    sheet = cells([row("A"), ["Open points"], row("B", cat=None, no=2)])  # the heading is row 5
    r = read(lst(tmp_path, None, sheets={"Projekt-OPL": {"cells": sheet, "merges": ["A5:K5"]}}))
    assert [t.subject for t in r.todos] == ["A", "B"] and r.todos[1].category == "Open points"
    assert "section_heading" in reasons(r)


# ---------- templates ----------

def test_blank_template_gives_no_task_and_says_why(tmp_path):
    stray = [[None, None, None, None, None, None, D, "Done", None, None, None],
             [None, None, None, None, None, None, D, "WIP", None, None, None]]
    dropdown = {"A1": "Status", "B1": "Prio", "C1": "Status", "A2": "Open", "B2": 1, "C2": "Open", "A3": "WIP",
                "B3": 2, "C3": "WIP", "A4": "Done", "B4": 3, "C4": "Done"}
    path = lst(tmp_path, None, sheets={"Projekt-OPL": {"cells": cells(stray), "merges": ["A1:K1"]},
                                       "Dropdown": {"cells": dropdown}})
    r = read(path)
    s = r.per_file_stats[0]
    assert r.todos == [] and s.error == "" and s.unrecognised == []  # the dropdown list is not an unread sheet
    assert s.rows == 2 and "blank template" in s.reason
    assert reasons(r) == ["no_tasks", "no_text", "no_text"]
    assert todos.summarize(r)["skipped_by_reason"] == {"no_text": 2}


def test_header_only_file_is_a_blank_template(tmp_path):
    r = read(lst(tmp_path, []))
    assert r.todos == [] and r.per_file_stats[0].layout == "todo" and "blank template" in r.per_file_stats[0].reason


def test_non_todo_sheet_is_unrecognised_not_an_error(tmp_path):
    other = {"A1": "Date", "B1": "Customer", "C1": "Potential", "D1": "Topics", "E1": "Next Step", **{
        f"A{i}": f"row {i}" for i in range(2, 12)}}
    r = read(make_book(tmp_path / "v.xlsx", {"S": {"cells": other}}))
    assert r.todos == [] and r.per_file_stats[0].layout == "none" and r.per_file_stats[0].unrecognised == ["S"]


# ---------- sales topics ----------

TOPIC_HEAD = {"A3": "Topic", "B3": "Due date", "C3": "Date provided"}


def topics(tmp_path, who="Name: XXX", name="topics.xlsx"):
    sheet = {"A1": who, **TOPIC_HEAD,
             "A4": "weekly report CW 13", "B4": D, "C4": "2026-x-xx",
             "A5": "business development PPT", "B5": dt.datetime(2026, 3, 24), "C5": dt.datetime(2026, 3, 20),
             "A6": "customer list_contacted", "B6": "2026-x-xx", "C6": "2026-x-xx"}
    return make_book(tmp_path / name, {"Sheet1": {"cells": sheet}})


def test_sales_topics_done_when_a_real_date_was_provided(tmp_path):
    r = read(topics(tmp_path))
    a, b, c = r.todos
    assert (a.layout, a.kind, a.state, a.due_iso) == ("topics", "topic", "open", "2026-04-07")
    assert (b.state, b.provided_iso, b.due_iso) == ("done", "2026-03-20", "2026-03-24")
    assert (c.state, c.due_iso, c.due_raw) == ("open", "", "")  # "2026-x-xx" is a date nobody filled in
    assert reasons(r) == ["owner_placeholder"]  # and no due-date review item for it


def test_topics_template_with_no_rep_is_tagged_and_a_named_rep_is_not(tmp_path):
    anonymous = read(topics(tmp_path))
    assert all(t.tags == [todos.TEMPLATE_TAG] and t.owner == "" for t in anonymous.todos)
    named = read(topics(tmp_path, who="Name: Mark Wang", name="n.xlsx"))
    assert all(t.tags == [] and t.owner == "Mark Wang" for t in named.todos) and reasons(named) == []


# ---------- due dates ----------

def test_due_variants(tmp_path):
    rows = [row("datetime", due=dt.datetime(2026, 5, 4)), row("tbc", due="TBC"), row("empty", due=None),
            row("decimal", due="5.25"), row("ambiguous", due="03/04"), row("phrase", due="next week"),
            row("weird", due="whenever"), row("text", due="2026-06-01"), row("far", due="02/03")]
    r = read(lst(tmp_path, rows))
    by = {t.subject: t for t in r.todos}
    assert by["datetime"].due_iso == "2026-05-04"
    assert (by["tbc"].due_iso, by["tbc"].due_raw) == ("", "TBC")
    assert (by["empty"].due_iso, by["empty"].due_raw) == ("", "")
    assert by["decimal"].due_iso == "2026-05-25"
    assert by["phrase"].due_iso == "2026-04-13"  # the Monday after the list date, as in the visit lists
    assert (by["weird"].due_iso, by["weird"].due_raw) == ("", "whenever")
    assert by["text"].due_iso == "2026-06-01"
    # 03/04 is 3 April or 4 March: 3 April is the day beside the list; 02/03 is the far one of 2 March / 3 February
    assert by["ambiguous"].due_iso == "2026-04-03"
    assert by["far"].due_iso == "" and by["far"].due_raw == "02/03"
    assert reasons(r) == ["ambiguous_due", "due_decimal"]


def test_due_without_any_reference_date_reads_only_what_is_unambiguous(tmp_path):
    rows = [row("iso", date=None, due="2026-05-04"), row("slash", date=None, due="03/04"), row("dec", date=None, due="5.25")]
    r = read(lst(tmp_path, rows, head=HEAD))
    by = {t.subject: t for t in r.todos}
    assert by["iso"].due_iso == "2026-05-04"
    assert by["slash"].due_iso == "" and by["dec"].due_iso == ""


# ---------- the IT project list and the switch ----------

def it_rows():
    return [row("Add users", who="Chenley", status="WIP"), row("Server move", who="Team, JET-IT", status="Done")]


def test_it_list_is_imported_tagged_and_appears_as_one_review_item(tmp_path):
    r = read(lst(tmp_path, it_rows(), name="43ac30b0ac15_to_do_list__JET-IT.xlsx"))
    assert len(r.todos) == 2 and all(t.tags == [todos.IT_TAG] for t in r.todos)
    assert reasons(r).count("it_project_list") == 1 and r.per_file_stats[0].it_list
    assert todos.summarize(r)["needs_check"] == 2


def test_it_list_is_found_by_its_content_too(tmp_path):
    r = read(lst(tmp_path, it_rows(), name="plain_list.xlsx"))
    assert all(t.tags == [todos.IT_TAG] for t in r.todos)
    r = read(lst(tmp_path, [row(who="Ethan")], name="itinerary_list.xlsx"))  # "it" inside a word is not IT
    assert r.todos[0].tags == []


def test_one_switch_leaves_the_it_list_out_and_keeps_the_rest(tmp_path):
    it = lst(tmp_path, it_rows(), name="to_do_list__JET-IT.xlsx")
    sales = lst(tmp_path, [row("Sales item")], name="sales_list.xlsx")
    r = read([it, sales], include_it_list=False)
    assert [t.subject for t in r.todos] == ["Sales item"]
    assert todos.summarize(r)["skipped_by_reason"] == {"it_list_excluded": 2}
    assert read([it, sales]).todos.__len__() == 3


# ---------- companies in the text ----------

def test_company_is_linked_only_on_one_sure_match(tmp_path, master):
    rows = [row("Quote", text="Send the quote to Cosmos Stellar Co., Ltd. this week"),
            row("Both", text="Compare Cosmos Stellar Co., Ltd. with Nova Quantum GmbH"),
            row("Unknown", text="Visit Invented Machines Inc to talk"),
            row("Chinese", text="跟进宇宙星辰科技有限公司的报价"),
            row("Plain", text="Weekly reports for everybody")]
    r = read(lst(tmp_path, rows), master=master)
    by = {t.subject: t for t in r.todos}
    assert by["Quote"].company["code"] == "KD00001" and by["Quote"].company["score"] >= 0.9
    assert by["Chinese"].company["code"] == "KD00001"
    assert not by["Both"].company and not by["Unknown"].company and not by["Plain"].company
    assert reasons(r) == ["company_ambiguous"]


def test_the_it_list_is_never_linked_to_a_company(tmp_path, master):
    r = read(lst(tmp_path, [row(text="Laptop for Cosmos Stellar Co., Ltd.")], name="to_do_list__JET-IT.xlsx"), master=master)
    assert not r.todos[0].company


# ---------- ids, copies, re-exports ----------

def source_ids(result):
    return [t.source_id for t in result.todos]


def test_ids_are_content_based_and_survive_status_due_and_row_changes(tmp_path):
    first = read(lst(tmp_path, [row("A", status="WIP"), row("B")]))
    later = read(lst(tmp_path, [row("Z", text="new row"), row("A", status="Done", due="2026-09-09"), row("B")],
                     name="later.xlsx"))
    a = {t.subject: t.source_id for t in first.todos}
    b = {t.subject: t.source_id for t in later.todos}
    assert a["A"] == b["A"] and a["B"] == b["B"] and b["Z"] not in a.values()


def test_identical_rows_in_one_sheet_stay_two_tasks(tmp_path):
    r = read(lst(tmp_path, [row("Same"), row("Same")]))
    ids = source_ids(r)
    assert len(set(ids)) == 2 and ids[1] == ids[0] + "#2"


def test_the_same_task_for_two_people_is_two_tasks(tmp_path):
    r = read(lst(tmp_path, [row("Same", who="Ethan"), row("Same", who="Mark")]))
    assert len(set(source_ids(r))) == 2


def test_copies_merge_and_the_richer_copy_wins(tmp_path):
    thin = lst(tmp_path, [row("A", text="do it", remarks=None, link=None, status=None, prio=None, due=None)], name="copy1.xlsx")
    rich = lst(tmp_path, [row("A", text="do it", remarks="done on site", link=None, status="Done")],
               name="aa11bb22cc33_copy1.xlsx")
    blank = lst(tmp_path, [], name="blank_copy.xlsx")
    r = read([thin, rich, blank])
    t, = r.todos
    assert (t.remarks, t.state, t.copies) == ("done on site", "done", 2)
    assert [f.reason for f in r.per_file_stats if f.path.endswith("blank_copy.xlsx")][0]  # the blank copy says why


def test_byte_identical_copies_are_read_once(tmp_path):
    one = lst(tmp_path, [row("A")], name="one.xlsx")
    shutil.copy(one, tmp_path / "copy_of_one.xlsx")
    r = read([one, str(tmp_path / "copy_of_one.xlsx")])
    assert len(r.todos) == 1 and r.todos[0].copies == 1
    assert [f["error"] for f in r.skipped_files] == ["identical to copy_of_one.xlsx"]


def test_a_broken_file_is_skipped_and_the_rest_is_read(tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a workbook")
    r = read([str(bad), lst(tmp_path, [row("A")], name="ok.xlsx")])
    assert len(r.todos) == 1 and len(r.skipped_files) == 1 and "could not be read" in r.skipped_files[0]["error"]


def test_only_mapped_columns_are_read(tmp_path):
    head = HEAD + ["Bank Account", "Tax ID"]
    r = read(lst(tmp_path, [row("A") + ["DE00 0000 0000 0000", "TAX-123"]], head=head))
    dump = json.dumps([vars(t) for t in r.todos], ensure_ascii=False, default=str)
    assert "DE00" not in dump and "TAX-123" not in dump
    assert "DE00" not in json.dumps(todos_margince.payload(r.todos[0]), ensure_ascii=False)


def test_bank_and_tax_details_in_text_cells_are_withheld(tmp_path):
    rows = [row("Pay", text="transfer to IBAN DE89 3704 0044 0532 0130 00 today", remarks="tax no. 91310000MA1FL0ABCD"),
            row("Ok", text="check the tax invoice layout", remarks="账号 6222 0000 1111"), row("Bank", text="Bank account details")]
    r = read(lst(tmp_path, rows))
    dump = json.dumps([todos_margince.payload(t) for t in r.todos], ensure_ascii=False)
    for secret in ("DE89", "91310000MA1FL0ABCD", "6222"):
        assert secret not in dump
    pay, ok, bank = r.todos  # the task stays, only the cell with the details goes
    assert (pay.text, pay.remarks, ok.text, ok.remarks, bank.text) == ("", "", "check the tax invoice layout", "", "")
    assert reasons(r).count("withheld_sensitive") == 4


# ---------- the CRM mapping ----------

class Crm(FakeCRM):
    def __init__(self, seats=None):
        super().__init__()
        self.seats = seats or {}
        self.refuse = set()

    def seats_by_name(self):
        return dict(self.seats)

    def log_activity(self, body):
        if body["source_id"] in self.refuse:
            raise _Conflict("x")
        return super().log_activity(body)


def make_crm(**kw):
    crm = Crm(seats={"Ethan Lv": "seat-ethan"})
    crm.companies["co9"] = {"id": "co9", "display_name": "宇宙星辰科技有限公司"}
    return crm


def todo_list(tmp_path, master, extra=()):
    rows = [row("Quote", text="Send the quote to Cosmos Stellar Co., Ltd.", status="WIP", prio=1, due=dt.datetime(2026, 5, 4)),
            row("Report", text="send it", status="Done", prio=2, who="Chenley", due="TBC"),
            row("Nova", text="Ring Nova Quantum GmbH", status="Open", who="Mark Wang", prio=None), *extra]
    return read(lst(tmp_path, rows, name="crm_list.xlsx"), master=master)


def test_payload_fits_the_task_contract_and_keeps_priority_in_the_body(tmp_path, master):
    r = todo_list(tmp_path, master)
    p = todos_margince.payload(r.todos[0], company_id="co9", seat_id="seat-ethan")
    allowed = {"kind", "subject", "body", "occurred_at", "due_at", "assignee_id", "links", "source", "source_system",
               "source_id", "raw"}
    assert set(p) <= allowed and "priority" not in p and p["kind"] == "task"
    assert p["assignee_id"] == "seat-ethan" and p["links"] == [{"entity_type": "company", "entity_id": "co9"}]
    assert p["due_at"].startswith("2026-05-04T09:00:00") and p["occurred_at"].startswith("2026-04-07")
    first = p["body"].splitlines()[0]
    assert first == "[Prio 1] [Status WIP] [Responsible Ethan Lv]"
    assert "Source: crm_list.xlsx, sheet Projekt-OPL, row 4" in p["body"]
    no_due = todos_margince.payload(r.todos[1])
    assert "due_at" not in no_due and 'Due as written: "TBC"' in no_due["body"] and "assignee_id" not in no_due


def test_tagged_task_carries_the_tag_in_subject_body_and_raw(tmp_path):
    r = read(lst(tmp_path, it_rows(), name="to_do_list__JET-IT.xlsx"))
    p = todos_margince.payload(r.todos[0])
    assert p["subject"].startswith("[needs-check: IT project list] ") and "Tag: needs-check: IT project list" in p["body"]
    assert p["raw"]["tags"] == [todos.IT_TAG]


def test_a_long_topic_is_cut_in_the_subject_and_whole_in_the_body(tmp_path):
    topic = "Very long topic " * 20
    p = todos_margince.payload(read(lst(tmp_path, [row(topic)])).todos[0])
    assert len(p["subject"]) <= todos_margince.SUBJECT_MAX and topic.strip() in p["body"]


def test_commit_writes_each_task_once_assigns_links_and_closes(tmp_path, master):
    crm, r = make_crm(), todo_list(tmp_path, master)
    out = todos_margince.commit(crm, r)
    assert (out["tasks_created"], out["tasks_present"], out["tasks_closed"], out["failed"]) == (3, 0, 1, [])
    assert (out["linked"], out["assigned"]) == (1, 1)
    tasks = {a["subject"]: a for a in crm.activities.values()}
    assert tasks["Quote"]["assignee_id"] == "seat-ethan" and tasks["Quote"]["links"][0]["entity_id"] == "co9"
    assert tasks["Report"]["is_done"] is True and tasks["Quote"]["is_done"] is False
    assert "assignee_id" not in tasks["Nova"] and tasks["Nova"]["links"] == []  # no seat, company not in the CRM
    assert all(a["source_system"] == "todolist" and a["kind"] == "task" for a in crm.activities.values())


def test_a_second_run_writes_nothing_and_an_added_row_writes_one(tmp_path, master):
    crm = make_crm()
    todos_margince.commit(crm, todo_list(tmp_path, master))
    before = json.dumps(crm.activities, sort_keys=True)
    again = todos_margince.commit(crm, todo_list(tmp_path, master))
    assert (again["tasks_created"], again["tasks_present"]) == (0, 3) and json.dumps(crm.activities, sort_keys=True) == before
    more = todos_margince.commit(crm, todo_list(tmp_path, master, extra=[row("Added", text="new line", who="Ethan")]))
    assert (more["tasks_created"], more["tasks_present"]) == (1, 3) and len(crm.activities) == 4


def test_preview_counts_and_writes_nothing(tmp_path, master):
    crm, r = make_crm(), todo_list(tmp_path, master)
    p = todos_margince.preview(crm, r)
    assert (p["tasks"], p["new"], p["present"], p["would_close"], p["linked"], p["assigned"]) == (3, 3, 0, 1, 1, 1)
    assert p["company_not_in_crm"] == 1 and p["reps_without_seat"] == ["Mark Wang"]
    assert crm.activities == {}
    todos_margince.commit(crm, r)
    after = todos_margince.preview(crm, r)
    assert (after["new"], after["present"], after["would_close"]) == (0, 3, 0)


def test_one_refused_task_is_reported_and_the_run_goes_on(tmp_path, master):
    crm, r = make_crm(), todo_list(tmp_path, master)
    crm.refuse = {r.todos[1].source_id}
    out = todos_margince.commit(crm, r)
    assert out["tasks_created"] == 2 and out["failed"] == [{"source_id": r.todos[1].source_id, "status": 409}]


def test_an_empty_read_commits_nothing(tmp_path):
    out = todos_margince.commit(make_crm(), read(lst(tmp_path, [])))
    assert (out["tasks_created"], out["tasks_present"], out["failed"]) == (0, 0, [])


def test_plan_needs_no_crm(tmp_path, master):
    p = todos_margince.plan(todo_list(tmp_path, master))
    assert (p["tasks"], p["done"], p["open"], p["with_due"], p["with_owner"], p["linked_to_company"]) == (3, 1, 2, 2, 2, 2)


# ---------- the real folder: counts only ----------

FOLDER = os.path.join(SAMPLE_DIR, "07 To-do and Sales Topics")


@pytest.mark.skipif(not os.path.isdir(FOLDER), reason="sample folder not present")
def test_real_folder_counts():
    paths = [os.path.join(FOLDER, f) for f in sorted(os.listdir(FOLDER)) if f.endswith(".xlsx")]
    r = todos.read_paths(paths, master=Master.from_records([]))
    s = todos.summarize(r)
    assert (s["files"], s["read_sheets"], s["skipped_files"]) == (4, 4, 0)
    assert (s["tasks"], s["done"], s["needs_check"]) == (29, 9, 17)
    per_file = {os.path.basename(f.path)[:16]: f.tasks for f in r.per_file_stats}
    assert sorted(per_file.values()) == [0, 4, 12, 13]  # the blank to-do template, topics template, one-on-one, IT list
    assert s["skipped_by_reason"] == {"no_text": 6}
    without_it = todos.summarize(todos.read_paths(paths, master=Master.from_records([]), include_it_list=False))
    assert without_it["tasks"] == 16 and without_it["skipped_by_reason"]["it_list_excluded"] == 13
