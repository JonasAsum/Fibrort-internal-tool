"""Quality checks for the to-do / sales-topic reader (folder 07), run on every row and every cell.

The oracle (oracle_todos.py) is built from the cell ledger alone. Each rule compares the reader with it over the
whole population (4 workbooks, every sheet, every row, every cell).

Declared drop rules (a cell, row or sheet that is not in a record must be named here, with its reason):
  sheets: a value list (the `Dropdown` sheet: status words and priority numbers only)
  rows:   a row with no topic and no to-do text (`no_text`: the stray status / date cells of a blank template)
  cells:  the header row; the title / owner-label cells above it; the `No.` column (a list sequence number);
          a date placeholder such as `2026-x-xx` (a date nobody filled in); a white-space-only cell
"""
import os
import re

import pytest

from excel import todos
from tests.quality import fam3_common as oc
from tests.quality import oracle_todos as oo
from tests.quality import snapshot, snapshot_todos

KNOWN_BUGS = {}
ALL_RULES = sorted(set(KNOWN_BUGS) | {
    "sheet-kind", "row-set", "cell-found", "owner", "status", "blank-template", "it-list-tag", "template-tag",
    "field-shape", "idempotent", "sheet-labels",
})
STATE = {"open": "open", "wip": "wip", "done": "done", "": "open"}
ISO = re.compile(r"^\d{4}-\d{2}-\d{2}$")
TEXT_FIELDS = {"topic": "topic", "todo": "text", "owner": "owner_raw", "status": "status_raw", "prio": "prio",
               "remarks": "remarks", "link": "link", "category": "category"}


def _v(rule, file, sheet, row, detail):
    return {"rule": rule, "file": file, "sheet": sheet, "row": row, "detail": str(detail)}


def _n(text):
    return re.sub(r"\s+", " ", str(text)).strip()


def _placeholder_date(text):
    t = re.sub(r"\s+", "", text).casefold()
    return "x" in t and oo.PLACEHOLDER_DATE.match(t) is not None


@pytest.fixture(scope="session")
def facts(sample_dir):
    if not os.path.isdir(sample_dir):
        pytest.skip("sample folder not found")
    return oo.all_facts(sample_dir)


@pytest.fixture(scope="session")
def result(facts):
    return todos.read_paths([f.path for f in facts])


def _find_cell(cell_field, cell, rec):
    """Is the ledger cell's content in the record, in the field it belongs to?"""
    text = _n(cell.text)
    if cell_field in TEXT_FIELDS:
        mine = _n(getattr(rec, TEXT_FIELDS[cell_field]))
        if cell_field == "todo" and not mine:  # a to-do text that is one line and has no topic is the subject
            mine = _n(rec.subject)
        return mine == text
    if cell_field == "date":
        return rec.date_iso == text[:10] and ISO.match(text[:10]) is not None
    if cell_field == "due":
        if _placeholder_date(text):
            return rec.due_iso == "" and rec.due_raw == ""  # declared: a date nobody filled in
        return rec.due_iso == text[:10] if ISO.match(text[:10]) else _n(rec.due_raw) == text
    if cell_field == "provided":
        if _placeholder_date(text):
            return rec.provided_iso == ""
        return rec.provided_iso == text[:10]
    if cell_field == "seq":
        return True  # declared drop: the No. column is the list's own sequence number
    return False


def _checks(facts, result, out):
    by_cell = {}
    for t in result.todos:
        by_cell.setdefault((t.file, t.sheet, t.row), []).append(t)
    reviews = {}
    for i in result.review:
        reviews.setdefault((i.file, i.sheet, i.row), set()).add(i.reason)
    stats = {os.path.basename(s.path): s for s in result.per_file_stats}
    expected = set()
    for f in facts:
        stat = stats.get(f.file)
        listed = {n for n, s in f.sheets.items() if s.kind in ("todo", "topics")}
        if stat is None:
            out.append(_v("sheet-kind", f.file, "", 0, "file not read"))
            continue
        if stat.read_sheets != len(listed):
            out.append(_v("sheet-kind", f.file, "", 0, f"reader reads {stat.read_sheets} sheets, oracle {len(listed)}"))
        if stat.unrecognised:
            out.append(_v("sheet-kind", f.file, "", 0, f"unread sheets {stat.unrecognised} (oracle: value lists {[n for n, s in f.sheets.items() if s.kind == 'value_list']})"))
        for name, sh in f.sheets.items():
            if sh.kind == "other":
                out.append(_v("sheet-kind", f.file, name, 0, "a sheet that is neither a task list nor a value list"))
                continue
            if sh.kind == "value_list":
                continue
            if stat.layout != sh.kind:
                out.append(_v("sheet-kind", f.file, name, 0, f"layout {stat.layout!r}, oracle {sh.kind!r}"))
            rows_with_text = [r for r in sh.rows if r.has_text]
            for row in sh.rows:
                key = (f.file, name, row.row)
                recs = by_cell.get(key, [])
                expected.add(key)
                if not row.has_text:
                    if recs:
                        out.append(_v("blank-template", f.file, name, row.row, "a row without topic or to-do text became a task"))
                    elif "no_text" not in reviews.get((f.file, name, row.row), set()):
                        out.append(_v("row-set", f.file, name, row.row, "row without text is not reported as no_text"))
                    continue
                if len(recs) != 1:
                    out.append(_v("row-set", f.file, name, row.row, f"{len(recs)} tasks for the row, expected 1"))
                    continue
                rec = recs[0]
                for fld, cell in row.cells.items():
                    if not _n(cell.text):
                        continue  # declared: white space only
                    if not _find_cell(fld, cell, rec):
                        out.append(_v("cell-found", f.file, name, row.row, f"{cell.coord} ({fld}) {_n(cell.text)[:50]!r} is not in the task"))
                _row_rules(f, sh, row, rec, reviews, out)
            if stat.rows != len(sh.rows) and len({s for s in f.sheets.values() if s.kind in ('todo', 'topics')}) == 1:
                out.append(_v("row-set", f.file, name, 0, f"reader counts {stat.rows} rows under the header, oracle {len(sh.rows)}"))
            _label_checks(f, sh, rows_with_text, by_cell, out)
        if not any(r.has_text for s in f.sheets.values() if s.kind in ("todo", "topics") for r in s.rows):
            if stat.tasks != 0 or any(t.file == f.file for t in result.todos):
                out.append(_v("blank-template", f.file, "", 0, "a template with no text gave tasks"))
            if "no_tasks" not in reviews.get((f.file, "", 0), set()):
                out.append(_v("blank-template", f.file, "", 0, "a blank template is not reported as such"))
        if f.is_it_list:
            for t in result.todos:
                if t.file == f.file and todos.IT_TAG not in t.tags:
                    out.append(_v("it-list-tag", f.file, t.sheet, t.row, "task of the IT list without its tag"))
        for t in result.todos:
            if t.file != f.file:
                continue
            sh = f.sheets.get(t.sheet)
            if sh and sh.kind == "topics" and sh.owner_is_placeholder and todos.TEMPLATE_TAG not in t.tags:
                out.append(_v("template-tag", f.file, t.sheet, t.row, "topic of an unnamed template without the template tag"))
    for t in result.todos:
        if (t.file, t.sheet, t.row) not in expected:
            out.append(_v("row-set", t.file, t.sheet, t.row, "task from a row the oracle does not have"))


def _row_rules(f, sh, row, rec, reviews, out):
    got = row.cells
    status_cell = _n(got["status"].text).lower() if "status" in got else ""
    if sh.kind == "todo":
        want = STATE.get(status_cell)
        if want is None:
            out.append(_v("status", f.file, sh.sheet, row.row, f"status word {status_cell!r} is not in the oracle's table"))
        elif rec.state != want:
            out.append(_v("status", f.file, sh.sheet, row.row, f"status {status_cell!r} -> state {rec.state!r}, expected {want!r}"))
    else:
        provided = "provided" in got and not _placeholder_date(_n(got["provided"].text))
        if rec.state != ("done" if provided else "open"):
            out.append(_v("status", f.file, sh.sheet, row.row, f"provided={got.get('provided') and got['provided'].text!r} -> state {rec.state!r}"))
    # owner: the Responsible cell, else the sheet's owner label; the reps' names are the only owners we assign
    owner_cell = _n(got["owner"].text) if "owner" in got else ""
    want_raw = owner_cell or ("" if sh.owner_is_placeholder else sh.owner_label)
    if _n(rec.owner_raw) != want_raw:
        out.append(_v("owner", f.file, sh.sheet, row.row, f"owner text {rec.owner_raw!r}, expected {want_raw!r}"))
    if rec.owner:
        first = rec.owner.split()[0].lower()
        if rec.owner not in oo.REPS or first not in want_raw.lower():
            out.append(_v("owner", f.file, sh.sheet, row.row, f"owner {rec.owner!r} does not come from {want_raw!r}"))
    elif want_raw and "owner_unknown" not in reviews.get((f.file, sh.sheet, row.row), set()):
        out.append(_v("owner", f.file, sh.sheet, row.row, f"unknown owner {want_raw!r} is not reported"))
    for d in (rec.date_iso, rec.due_iso, rec.provided_iso):
        if d and not ISO.match(d):
            out.append(_v("field-shape", f.file, sh.sheet, row.row, f"date {d!r} is not ISO"))
    if rec.prio and rec.prio not in ("1", "2", "3"):
        out.append(_v("field-shape", f.file, sh.sheet, row.row, f"prio {rec.prio!r}"))
    if rec.state not in ("open", "wip", "waiting", "done", "cancelled"):
        out.append(_v("field-shape", f.file, sh.sheet, row.row, f"state {rec.state!r}"))


def _label_checks(f, sh, rows, by_cell, out):
    for row in rows:
        recs = by_cell.get((f.file, sh.sheet, row.row), [])
        if len(recs) == 1 and _n(recs[0].list_title) != sh.title:
            out.append(_v("sheet-labels", f.file, sh.sheet, row.row, f"list title {recs[0].list_title!r}, the title cell says {sh.title!r}"))
            break


def _idempotent(facts, out):
    paths = [f.path for f in facts]
    a, b = snapshot_todos.records_of(todos.read_paths(paths)), snapshot_todos.records_of(todos.read_paths(paths))
    if a != b:
        out.append(_v("idempotent", "", "", 0, snapshot.report(snapshot.diff(a, b))))


def _shape(result, out):
    ids = [t.source_id for t in result.todos]
    for i in {x for x in ids if ids.count(x) > 1}:
        out.append(_v("field-shape", "", "", 0, f"source id {i} used twice"))
    for t in result.todos:
        if not t.subject:
            out.append(_v("field-shape", t.file, t.sheet, t.row, "task without subject"))


@pytest.fixture(scope="session")
def violations(facts, result):
    out = []
    _checks(facts, result, out)
    _shape(result, out)
    _idempotent(facts, out)
    return out


def _param(rule):
    if rule in KNOWN_BUGS:
        return pytest.param(rule, marks=pytest.mark.xfail(strict=True, reason=KNOWN_BUGS[rule]))
    return rule


@pytest.mark.parametrize("rule", [_param(r) for r in ALL_RULES])
def test_rule_has_no_violations(rule, violations):
    bad = [v for v in violations if v["rule"] == rule]
    assert not bad, f"{len(bad)} violation(s) of '{rule}':\n" + "\n".join(
        f"  {v['file'][:34]} {v['sheet'][:12]} row {v['row']}: {v['detail'][:170]}" for v in bad[:6])


def test_population(facts, result):
    """4 workbooks; 3 task lists with 6 + 12 + 13 data rows (+ 4 topic rows); the blank template yields nothing."""
    assert len(facts) == 4
    rows = {f.file[:16]: sum(len(s.rows) for s in f.sheets.values() if s.kind in ("todo", "topics")) for f in facts}
    assert sorted(rows.values()) == [4, 6, 12, 13]
    assert len(result.todos) == 29
