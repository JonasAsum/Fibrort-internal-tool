"""Margince adapter for the to-do reader (excel/todos.py): the only place that knows Margince.

A to-do or sales topic becomes a CRM task (an activity of kind `task`, written through
POST /v1/activities so it can carry source_system + source_id). The CRM task has an assignee, a
due time and done/open, but no priority, so priority, the status as written, the remarks and the
list's own words go in a plain first line of the body. A task whose Responsible cell names a rep
with a seat is assigned to that seat; any other name stays in the body, never guessed at. A task
links to a company only when the reader found exactly one master company in its text and the CRM
holds that company; this module never creates companies.

Idempotent: the source id is built from the task's content (todos.py), so a second run finds the
tasks it wrote and writes nothing; the same list with one more row writes one task. Status and due
date are not part of the id, and an existing task is never changed.

  preview(client, result) - reads the CRM only; what a commit would write
  commit(client, result)  - writes the tasks that are not there yet; done ones are closed right after
"""
from zoneinfo import ZoneInfo

from common.client import ApiError
from visits import when
from excel.split import fold_key

from . import todos

SOURCE_SYSTEM = "todolist"
TIMEZONE = "Asia/Shanghai"
SUBJECT_MAX = 200


def subject(todo) -> str:
    text = (f"[{todo.tags[0]}] " if todo.tags else "") + todo.subject
    return text if len(text) <= SUBJECT_MAX else text[:SUBJECT_MAX - 1].rstrip() + "…"


def body(todo) -> str:
    """The first line carries what the CRM task has no field for; the rest is the row's own words."""
    head = []
    if todo.prio:
        head.append(f"Prio {todo.prio}")
    if todo.layout == "topics":
        head.append(f"Provided {todo.provided_iso}" if todo.provided_iso else "Not provided yet")
    elif todo.status_raw:
        head.append(f"Status {todo.status_raw}")
    if todo.owner or todo.owner_raw:
        head.append(f"Responsible {todo.owner or todo.owner_raw}")
    lines = ["[" + "] [".join(head) + "]"] if head else []
    if len(todo.subject) > SUBJECT_MAX - 40:  # the subject line is cut: keep the whole topic here
        lines.append(todo.subject)
    if todo.text:
        lines.append(todo.text)
    for label, value in (("Category", todo.category), ("Remarks", todo.remarks), ("Link or note", todo.link)):
        if value:
            lines.append(f"{label}: {value}")
    if todo.due_raw:
        lines.append(f'Due as written: "{todo.due_raw}"')
    if todo.date_iso:
        lines.append(f"Listed on {todo.date_iso}")
    if todo.list_title:
        lines.append(f"List: {todo.list_title}")
    lines.append(f"Source: {todo.file}, sheet {todo.sheet}, row {todo.row}")
    lines += [f"Tag: {tag}" for tag in todo.tags]
    return "\n".join(lines)


def payload(todo, company_id=None, seat_id=None, tz=TIMEZONE, source_system=SOURCE_SYSTEM) -> dict:
    zone = ZoneInfo(tz)
    out = {"source": "import", "source_system": source_system, "source_id": todo.source_id, "kind": "task",
           "subject": subject(todo), "body": body(todo), "links": [],
           "raw": {"todo_list_row": {"file": todo.file, "sheet": todo.sheet, "row": todo.row, "kind": todo.kind,
                                     "status": todo.status_raw, "prio": todo.prio, "responsible": todo.owner_raw,
                                     "due": todo.due_iso or todo.due_raw, "list_date": todo.date_iso},
                   "tags": list(todo.tags)}}
    if todo.date_iso:
        out["occurred_at"] = when.day_instant(todo.date_iso, zone).isoformat()
    if todo.due_iso:
        out["due_at"] = when.day_instant(todo.due_iso, zone).isoformat()
    if company_id:
        out["links"].append({"entity_type": "company", "entity_id": company_id})
    if seat_id:
        out["assignee_id"] = seat_id
    return out


def plan(result) -> dict:
    """Counts straight from the read: no CRM involved."""
    return todos.summarize(result)


class _Lookup:
    """What the CRM already holds, read once: this source's tasks, the seats and the companies."""

    def __init__(self, client, source_system):
        self.existing = {a.get("source_id") for a in client.paginate("/v1/activities", {"limit": 200, "kind": "task"})
                         if a.get("source_system") == source_system}
        self.seats = {when.fold_person(n): i for n, i in client.seats_by_name().items()}
        self.companies = {fold_key(n): i for n, i in client.company_id_by_name().items()}

    def seat(self, todo):
        return self.seats.get(when.fold_person(todo.owner)) if todo.owner else None

    def company(self, todo):
        names = [todo.company.get("cn"), todo.company.get("en")] if todo.company else []
        return next((self.companies[fold_key(n)] for n in names if n and fold_key(n) in self.companies), None)


def _counts(result, lookup, only_new) -> dict:
    new = [t for t in result.todos if t.source_id not in lookup.existing]
    scope = new if only_new else result.todos
    return {"tasks": len(result.todos), "new": len(new), "present": len(result.todos) - len(new),
            "would_close": sum(t.done for t in new),
            "linked": sum(lookup.company(t) is not None for t in scope),
            "company_not_in_crm": sum(bool(t.company) and lookup.company(t) is None for t in scope),
            "assigned": sum(lookup.seat(t) is not None for t in scope),
            "unassigned": sum(lookup.seat(t) is None for t in scope),
            "reps_without_seat": sorted({t.owner for t in scope if t.owner and lookup.seat(t) is None})}


def preview(client, result, source_system=SOURCE_SYSTEM) -> dict:
    """What a commit would write, from reads only."""
    return _counts(result, _Lookup(client, source_system), only_new=True)


def commit(client, result, source_system=SOURCE_SYSTEM, on_progress=None) -> dict:
    """Writes the tasks the CRM does not hold yet; a second run writes nothing. One task the CRM
    refuses is reported in "failed" (source id and status) and the run goes on."""
    emit = on_progress or (lambda step, **info: None)
    lookup = _Lookup(client, source_system)
    out = {"tasks_created": 0, "tasks_present": 0, "tasks_closed": 0, "linked": 0, "assigned": 0, "failed": []}
    emit("tasks_start", total=len(result.todos))
    for i, todo in enumerate(result.todos):
        if todo.source_id in lookup.existing:
            out["tasks_present"] += 1
        else:
            company_id, seat_id = lookup.company(todo), lookup.seat(todo)
            try:
                created, task = client.log_activity(payload(todo, company_id, seat_id, source_system=source_system))
                lookup.existing.add(todo.source_id)
                if not created:
                    out["tasks_present"] += 1
                else:
                    out["tasks_created"] += 1
                    out["linked"] += company_id is not None
                    out["assigned"] += seat_id is not None
                    if todo.done and not task.get("is_done"):
                        client.patch(f"/v1/activities/{task['id']}", {"is_done": True})
                        out["tasks_closed"] += 1
            except ApiError as exc:
                out["failed"].append({"source_id": todo.source_id, "status": exc.response.status_code})
        emit("tasks_progress", done=i + 1, total=len(result.todos), created=out["tasks_created"])
    return out
