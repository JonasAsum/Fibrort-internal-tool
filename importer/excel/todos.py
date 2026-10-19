"""Read-only reader for to-do and sales-topic workbooks (folder "07 To-do and Sales Topics"). Pure
parsing: nothing here knows the CRM. Headers are found by label (English, Chinese or German, any
column order, header row anywhere in the first rows); only the mapped columns are read, so a
stray column (bank details, tax ids ...) can never reach a record. Two layouts:

  todo    a project to-do list (Projekt-OPL): No., Categorie, Date, Topics, To Do / Next Steps,
          Responsible, Due Date, Status (Open / WIP / Done), Prio, Remarks, Documents or Websites.
  topics  a sales-topics list: Topic, Due date, Date provided. A topic is done once a real date
          is provided (a placeholder such as 2026-x-xx means not yet).

Rows without topic or to-do text (the filled-in status and date cells of an empty template) are
skipped with a reason; a sheet that ends up with no task says so, it is not an error. Every task
gets a content-based source id (who, topic, text), so the same list re-exported with one more row
adds only that row; copies of one list merge on that id and the richer copy wins.

Tasks are tagged "needs-check: ..." where the owner asked to look first: the IT project list, and
a topics template that names no rep. A task links to a company only when the master list knows a
name in its text with a score of 0.9 or more and exactly one company does."""
import hashlib
import os
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime

from visits.when import resolve_due

from .dates import parse_date, parse_due
from .headers import field_for as visit_field
from .headers import label_field, normalize
from .readers import ReviewItem, normalize_owner
from .files import unique_contents
from .sheet import load
from .split import cell_text, fold_key, is_placeholder

FAMILY = "todos"
SCAN_ROWS = 30
COMPANY_OK = 0.9
IT_TAG = "needs-check: IT project list"
TEMPLATE_TAG = "needs-check: topics template, no rep named"
IT_FILE = re.compile(r"(?<![A-Za-z])IT(?![A-Za-z])")
UNRECOGNISED_MIN_CELLS = 10
SKIP_REASONS = {"no_text", "section_heading", "it_list_excluded"}  # rows left out; every other reason is a note

# field -> labels; a leading "=" means the whole label must equal it, otherwise it may start with it.
# Order matters: the first field that claims a label wins ("due date" is not the list's "date").
SYNONYMS = {
    "due": ("due date", "due", "deadline", "target date", "截止日期", "截止时间", "截止", "到期日", "期限", "计划完成",
            "fällig", "faellig", "termin"),
    "provided": ("date provided", "provided", "delivered", "提交日期", "已提供"),
    "status": ("status", "状态", "进度", "完成情况"),
    "prio": ("prio", "priority", "优先级", "priorität"),
    "owner": ("responsible", "owner", "assignee", "负责人", "责任人", "跟进人", "执行人", "verantwortlich"),
    "todo": ("to do / next steps", "to do", "todo", "next steps", "next step", "action", "待办", "下一步",
             "行动", "措施", "maßnahme", "massnahme"),
    "topic": ("topics", "topic", "subject", "主题", "议题", "事项", "项目", "thema"),
    "category": ("categorie", "category", "类别", "分类", "kategorie"),
    "remarks": ("remarks", "remark", "comment", "notes", "note", "备注", "说明", "bemerkung"),
    "link": ("documents or websites", "document", "link", "文档", "链接", "dokument"),
    "seq": ("=no", "=no.", "=#", "=nr", "=序号", "=编号"),
    "date": ("=date", "=日期", "=datum", "=list date", "=raised"),
}
OWNER_LABEL = re.compile(r"^(?:project responsible|responsible|owner|name|sales|rep|负责人|责任人|姓名)\s*[:：]\s*(.*)$", re.I)
TITLE_LABEL = re.compile(r"^topic\s*[-–:]\s*(.*)$", re.I)
# Bank and tax details stay out of the CRM: a text cell holding some is withheld (and reported without its value).
SENSITIVE = re.compile(r"税号|纳税人识别|开户|银行[账帐]|[账帐]号|\b(?:iban|swift|bic|ust-?id|vat\s*(?:id|no|number))\b|"
                       r"bank\s*(?:acc|detail)|\b[A-Z]{2}\d{2}(?:\s?[0-9A-Z]{4}){3,7}\b", re.I)
CREDIT_CODE = re.compile(r"\b(?=[0-9A-Z]*\d)[0-9A-Z]{18}\b")  # a Chinese unified social credit code (tax ID)
PLACEHOLDER_NAME = re.compile(r"^(?:x+|_+|\.+|-+|n/?a|tbd|tbc)$", re.I)

# state -> words; checked in this order, "open" before "done" so 未完成 is not read as 完成.
STATUS_WORDS = {
    "cancelled": ("cancelled", "canceled", "dropped", "abandoned", "abgebrochen", "取消", "已取消", "作废"),
    "open": ("open", "new", "todo", "to do", "pending", "planned", "not started", "offen", "neu", "未开始",
             "未完成", "待办", "待处理", "未处理", "计划"),
    "done": ("done", "closed", "complete", "completed", "finished", "ok", "resolved", "erledigt", "fertig",
             "abgeschlossen", "完成", "已完成", "已关闭", "关闭", "已解决"),
    "wip": ("wip", "in progress", "ongoing", "doing", "started", "in arbeit", "laufend", "进行中", "处理中",
            "在做", "跟进中", "执行中"),
    "waiting": ("waiting", "on hold", "hold", "blocked", "waiting for", "paused", "wartet", "等待", "暂停",
                "搁置", "阻塞"),
}
CLOSED = {"done", "cancelled"}
DATE_FIELDS = ("date", "due", "provided")


@dataclass
class Todo:
    """One task as read from one row (then merged across copies)."""
    kind: str  # todo | topic
    subject: str  # topic, else the first line of the to-do text
    text: str = ""  # the to-do / next-steps text
    topic: str = ""
    category: str = ""
    owner: str = ""  # one of readers.REPS, "" when the cell names nobody we know
    owner_raw: str = ""  # the Responsible cell (or the sheet's owner label) as written
    status_raw: str = ""
    state: str = "open"  # open | wip | waiting | done | cancelled
    prio: str = ""
    remarks: str = ""
    link: str = ""  # Documents or Websites cell: a link, or a short status note
    date_iso: str = ""  # the list's Date column
    due_iso: str = ""
    due_raw: str = ""  # the due cell as written when it gave no date ("TBC", "next week")
    provided_iso: str = ""
    list_title: str = ""
    section: str = ""  # a heading row above this one
    tags: list = field(default_factory=list)
    company_text: str = ""  # the name found in the text
    company: dict = field(default_factory=dict)  # code, cn, en, score, method of the unique master match
    file: str = ""
    sheet: str = ""
    row: int = 0
    layout: str = ""
    source_id: str = ""
    copies: int = 1  # rows (in any file) that said the same task

    @property
    def done(self) -> bool:
        return self.state in CLOSED

    def richness(self) -> int:
        return sum(bool(v) for v in (self.subject, self.text, self.topic, self.category, self.owner_raw,
                                     self.status_raw, self.prio, self.remarks, self.link, self.date_iso,
                                     self.due_iso, self.due_raw, self.provided_iso))


@dataclass
class TodosFileStat:
    path: str
    layout: str = "none"
    sheets: int = 0
    read_sheets: int = 0
    rows: int = 0  # non-blank rows under a header
    tasks: int = 0  # tasks read from this file, before merging copies
    blank_rows: int = 0
    it_list: bool = False
    reason: str = ""  # why a recognised list gave no task
    error: str = ""
    unrecognised: list = field(default_factory=list)
    sheet_errors: list = field(default_factory=list)


@dataclass
class TodosResult:
    todos: list
    review: list  # ReviewItem: rows left out and notes on rows that went in
    skipped_files: list  # [{"file", "error"}]
    per_file_stats: list = field(default_factory=list)


# ---------- cells ----------

def _placeholder_date(value) -> bool:
    """"2026-x-xx" and the like: a date that was never filled in."""
    t = re.sub(r"\s+", "", str(value)).casefold()
    return isinstance(value, str) and "x" in t and re.fullmatch(r"[\dx\-/.]+", t) is not None


def map_status(raw) -> tuple[str, bool]:
    """(state, known): an empty cell is an open task; a word we do not know is open and not known."""
    t = normalize(unicodedata.normalize("NFKC", "" if raw is None else str(raw)))
    if not t:
        return "open", True
    word = (re.match(r"[^\W\d_]+", t) or [""])[0]
    for state, words in STATUS_WORDS.items():
        if t in words or word in words or any(w in t for w in words if not w.isascii()):
            return state, True
    return "open", False


# ---------- header and sheet ----------

def find_todo_header(sheet):
    """(row, {field: column}) of the first row that labels a topic or to-do column and at least two
    more fields; a visit list's header (customer, week, potential) never counts."""
    for r in range(1, min(SCAN_ROWS, sheet.max_row) + 1):
        cols, visit = {}, False
        for c in range(1, sheet.max_col + 1):
            if sheet.is_merged_continuation(r, c):
                continue
            value = sheet.value(r, c)
            visit = visit or visit_field(value) in ("customer", "cw", "potential", "customer_type")
            name = label_field(value, SYNONYMS)
            if name and name not in cols:
                cols[name] = c
        if not visit and len(cols) >= 3 and ("topic" in cols or "todo" in cols):
            return r, cols
    return None


def detect(sheet) -> str | None:
    """"todo", "topics" or None."""
    found = find_todo_header(sheet)
    if found is None:
        return None
    return "todo" if {"todo", "owner", "status", "prio"} & found[1].keys() else "topics"


def _is_value_list(sheet) -> bool:
    """The Dropdown sheet of the template: only status words and priority numbers."""
    words = {w for ws in STATUS_WORDS.values() for w in ws} | {"status", "prio", "priority"}
    cells = [normalize(str(v)) for v in sheet.grid.values()]
    return bool(cells) and all(c in words or c.isdigit() for c in cells)


def _sheet_labels(sheet, header_row):
    """(owner, title) from the rows above the header: "Project Responsible: <name>" or a label with
    the name in the next cell, and "TOPIC - <title>". Placeholders (XXXX) come back empty."""
    owner, title = "", ""
    for r in range(1, header_row):
        for c in range(1, sheet.max_col + 1):
            if sheet.is_merged_continuation(r, c):
                continue
            text = cell_text(sheet.value(r, c), keep_lines=True)
            if not text:
                continue
            if m := TITLE_LABEL.match(text):
                title = title or m.group(1).strip()
            elif m := OWNER_LABEL.match(text):
                rest = m.group(1).strip()
                if not rest:  # the name sits in the next cell of the row
                    rest = next((cell_text(sheet.value(r, cc), keep_lines=True) for cc in range(c + 1, sheet.max_col + 1)
                                 if not sheet.is_merged_continuation(r, cc) and cell_text(sheet.value(r, cc), keep_lines=True)), "")
                owner = owner or rest
    placeholder = lambda s: not s or bool(PLACEHOLDER_NAME.match(s))  # noqa: E731
    return ("" if placeholder(owner) else owner), ("" if placeholder(title) else title), bool(owner) and placeholder(owner)


def _year_of(path) -> int | None:
    m = re.search(r"(20\d{2})", os.path.basename(path))
    return int(m.group(1)) if m else None


def _reference_day(sheet, header_row, cols) -> date | None:
    """The earliest date in the list's date columns: what an undated "5.25" or "next week" counts from."""
    days = []
    for f in DATE_FIELDS:
        c = cols.get(f)
        for r in range(header_row + 1, sheet.max_row + 1) if c else ():
            v = sheet.value(r, c)
            if isinstance(v, datetime):
                days.append(v.date())
            elif isinstance(v, date):
                days.append(v)
    return min(days) if days else None


def _due(value, ref, year):
    """(iso, raw, flag): iso from a date, a decimal or a phrase; raw is the text when it gave no date."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return "", "", ""
    if _placeholder_date(value) or is_placeholder(value):
        return "", "", ""
    if ref is not None:
        iso, phrase, flag = parse_due(value, ref)
        if iso is None and phrase:
            when = resolve_due(phrase, ref)
            if when:
                return when.isoformat(), "", "phrase_resolved"
        return iso or "", phrase, flag
    iso, flag = parse_date(value, None, year)
    return (iso or "", "", flag) if iso else ("", cell_text(value, keep_lines=True), "phrase")


def read_todos_sheet(sheet, path):
    """(layout, todos, review items, {"rows", "blank"}); layout None when the sheet is not a to-do list."""
    found = find_todo_header(sheet)
    if found is None:
        return None, [], [], {}
    header_row, cols = found
    layout = "todo" if {"todo", "owner", "status", "prio"} & cols.keys() else "topics"
    file = os.path.basename(path)
    sheet_owner, title, owner_placeholder = _sheet_labels(sheet, header_row)
    ref, year = _reference_day(sheet, header_row, cols), _year_of(path)
    items, todos, blank, rows, section = [], [], 0, 0, ""

    def item(r, fld, value, reason):
        items.append(ReviewItem(FAMILY, file, sheet.name, r, fld, "" if value is None else str(value), reason))

    for r in range(header_row + 1, sheet.max_row + 1):
        own = [c for c in range(1, sheet.max_col + 1)
               if not sheet.is_merged_continuation(r, c) and sheet.value(r, c) is not None
               and str(sheet.value(r, c)).strip() != ""]
        if not own:
            blank += 1
            continue
        rows += 1
        cell = {f: sheet.value(r, c) for f, c in cols.items()}

        def free(fld):
            t = cell_text(cell.get(fld), keep_lines=True)
            if t and (SENSITIVE.search(t) or CREDIT_CODE.search(t)):
                item(r, fld, "", "withheld_sensitive")
                return ""
            return t

        topic, text = free("topic"), free("todo")
        if len(own) == 1 and any(sheet.is_merged_continuation(r, c) and sheet.top.get((r, c)) == (r, own[0])
                                 for c in range(1, sheet.max_col + 1)):
            section = cell_text(sheet.value(r, own[0]), keep_lines=True)
            item(r, "topic", section, "section_heading")
            continue
        if not topic and not text:
            item(r, "topic", "", "no_text")
            continue
        owner_raw = cell_text(cell.get("owner"), keep_lines=True) or sheet_owner
        owner = normalize_owner(owner_raw)
        if owner_raw and not owner:
            item(r, "owner", owner_raw, "owner_unknown")
        status_raw = cell_text(cell.get("status"), keep_lines=True)
        state, known = map_status(status_raw)
        if not known:
            item(r, "status", status_raw, "unknown_status")
        row_day = None
        date_iso = ""
        if cols.get("date"):
            date_iso, flag = parse_date(cell.get("date"), ref, year)
            date_iso = date_iso or ""
            if flag in ("ambiguous", "unparsed") and cell.get("date") is not None:
                item(r, "date", cell.get("date"), "ambiguous_date" if flag == "ambiguous" else "unparsed_date")
                date_iso = "" if flag == "ambiguous" else date_iso
            row_day = date.fromisoformat(date_iso) if date_iso else None
        due_iso, due_raw, flag = _due(cell.get("due"), row_day or ref, year)
        if flag == "ambiguous":  # a guessed day is worse than none: the text stays, the date waits
            item(r, "due", cell.get("due"), "ambiguous_due")
            due_iso, due_raw = "", cell_text(cell.get("due"), keep_lines=True)
        elif flag == "decimal":
            item(r, "due", cell.get("due"), "due_decimal")
        provided_iso = ""
        if "provided" in cols and cell.get("provided") is not None and not _placeholder_date(cell.get("provided")):
            provided_iso, _ = parse_date(cell.get("provided"), ref, year)
            provided_iso = provided_iso or ""
        if layout == "topics":
            state = "done" if provided_iso else "open"
        prio = cell_text(cell.get("prio"), keep_lines=True)
        todos.append(Todo(
            kind="todo" if layout == "todo" else "topic",
            subject=topic or text.split("\n")[0], text=text if topic or "\n" in text else "",
            topic=topic, category=free("category") or section, owner=owner, owner_raw=owner_raw,
            status_raw=status_raw, state=state, prio=prio, remarks=free("remarks"),
            link=free("link"), date_iso=date_iso, due_iso=due_iso, due_raw=due_raw,
            provided_iso=provided_iso, list_title=title, section=section, file=file, sheet=sheet.name, row=r,
            layout=layout))
        if layout == "topics" and not sheet_owner:
            todos[-1].tags.append(TEMPLATE_TAG)
    if layout == "topics" and owner_placeholder and todos:
        item(header_row, "owner", "", "owner_placeholder")
    return layout, todos, items, {"rows": rows, "blank": blank}


# ---------- company in text ----------

LATIN_TOKEN = re.compile(r"[A-Za-zÀ-ÿ0-9&][A-Za-zÀ-ÿ0-9&'’.\-]*")
MIN_LETTERS = 5


def _phrases(text):
    tokens = [t.strip(".-'’") for t in LATIN_TOKEN.findall(text)]
    tokens = [t for t in tokens if t]
    for n in (4, 3, 2, 1):
        for i in range(len(tokens) - n + 1):
            window = tokens[i:i + n]
            if not (window[0][0].isupper() or window[0][0].isdigit()):
                continue
            phrase = " ".join(window)
            if sum(ch.isalpha() for ch in phrase) >= MIN_LETTERS:
                yield phrase
    for run in re.findall(r"[一-鿿]{4,}", text):
        yield run
        for m in re.finditer("公司", run):
            end = m.end()
            for start in range(0, min(9, end - 3)):
                yield run[start:end]


def find_company(text, master, cache=None):
    """({code, cn, en, score, method} | None, ambiguous): the one company the master knows by a
    name in the text with a score >= 0.9. Two different companies, or none, link nothing."""
    if master is None or master.size == 0 or not text:
        return None, False
    cache = {} if cache is None else cache
    hits = {}
    for phrase in _phrases(text):
        if phrase not in cache:
            m = master.match(phrase)
            ok = m.score >= COMPANY_OK and m.method not in ("new", "none", "review", "ambiguous") and (m.cn or m.en)
            cache[phrase] = {"code": m.code, "cn": m.cn, "en": m.en, "score": m.score, "method": m.method,
                             "phrase": phrase} if ok else None
        if cache[phrase]:
            hit = cache[phrase]
            hits.setdefault(hit["code"] or (hit["cn"], hit["en"]), hit)
    if len(hits) == 1:
        return next(iter(hits.values())), False
    return None, len(hits) > 1


# ---------- files ----------



def _is_it_list(path, todos) -> bool:
    return bool(IT_FILE.search(os.path.splitext(os.path.basename(path))[0])) or any(
        re.search(r"jet-?it", t.owner_raw, re.I) for t in todos)


def _name_day(path) -> str:
    m = re.match(r"(\d{4}-\d{2}-\d{2})", os.path.basename(path))
    return m.group(1) if m else ""


def _assign_ids(todos):
    """Content-based ids: who, topic and text. Identical tasks in one sheet get #2, #3 ... in sheet order."""
    seen = Counter()
    for t in todos:
        who = t.owner or fold_key(t.owner_raw)
        digest = hashlib.sha1("|".join((t.kind, who, fold_key(t.topic), fold_key(t.text or t.subject))).encode()).hexdigest()[:12]
        base = f"{t.kind}:{digest}"
        seen[(t.file, t.sheet, base)] += 1
        n = seen[(t.file, t.sheet, base)]
        t.source_id = base if n == 1 else f"{base}#{n}"


def _merge(todos):
    """One task per id. The richest copy is the base, then the newer file; a copy never adds a blank record."""
    groups = defaultdict(list)
    for t in todos:
        groups[t.source_id].append(t)
    out = []
    for group in groups.values():
        group.sort(key=lambda t: (t.file, t.sheet, t.row))
        group.sort(key=lambda t: (t.richness(), _name_day(t.file)), reverse=True)  # stable: ties keep file order
        base = group[0]
        base.copies = len(group)
        for other in group[1:]:
            for f in ("owner_raw", "prio", "remarks", "link", "category", "date_iso", "due_iso", "due_raw"):
                if not getattr(base, f) and getattr(other, f):
                    setattr(base, f, getattr(other, f))
            for tag in other.tags:
                if tag not in base.tags:
                    base.tags.append(tag)
        out.append(base)
    return sorted(out, key=lambda t: (t.file, t.sheet, t.row))


def read_paths(paths, master=None, include_it_list=True) -> TodosResult:
    """Every to-do list in the files. include_it_list=False leaves out the IT project list (rows are
    skipped with the reason "it_list_excluded"). A bad file or sheet is recorded and the run goes on."""
    if master is None:
        from .run import IMPORTER_DIR, _master
        master = _master(IMPORTER_DIR)
    skipped, stats, todos, items, cache = [], [], [], [], {}
    for path in unique_contents(paths, skipped):
        stat = TodosFileStat(path=path)
        stats.append(stat)
        try:
            sheets = load(path)
        except Exception as exc:  # a corrupt or unsupported workbook must not stop the run
            stat.error = f"{type(exc).__name__}: {exc}"
            skipped.append({"file": path, "error": f"could not be read: {stat.error}"})
            continue
        stat.sheets = len(sheets)
        layouts, file_todos, file_items, recognised = Counter(), [], [], 0
        for sheet in sheets:
            if sheet.cut:
                stat.sheet_errors.append((sheet.name, sheet.cut))
            try:
                layout, got, found, counts = read_todos_sheet(sheet, path)
            except Exception as exc:
                stat.error = (stat.error + "; " if stat.error else "") + f"{sheet.name}: {exc}"
                stat.sheet_errors.append((sheet.name, str(exc) or type(exc).__name__))
                continue
            if layout is None:
                if len(sheet.grid) >= UNRECOGNISED_MIN_CELLS and not _is_value_list(sheet):
                    stat.unrecognised.append(sheet.name)
                continue
            layouts[layout] += 1
            recognised += 1
            stat.read_sheets += 1
            stat.rows += counts["rows"]
            stat.blank_rows += counts["blank"]
            file_todos += got
            file_items += found
        stat.layout = layouts.most_common(1)[0][0] if layouts else "none"
        stat.it_list = bool(file_todos) and _is_it_list(path, file_todos)
        name = os.path.basename(path)
        if stat.it_list:
            for t in file_todos:
                t.tags.append(IT_TAG)
            file_items.append(ReviewItem(FAMILY, name, "", 0, "file", name, "it_project_list"))
            if not include_it_list:
                file_items += [ReviewItem(FAMILY, t.file, t.sheet, t.row, "topic", "", "it_list_excluded")
                               for t in file_todos]
                file_todos = []
        elif master is not None:
            for t in file_todos:
                t.company, ambiguous = find_company("\n".join((t.topic, t.text, t.remarks)), master, cache)
                t.company_text = t.company.get("phrase", "") if t.company else ""
                if ambiguous:
                    file_items.append(ReviewItem(FAMILY, t.file, t.sheet, t.row, "topic", "", "company_ambiguous"))
        stat.tasks = len(file_todos)
        if recognised and not file_todos:
            stat.reason = ("it_list_excluded" if stat.it_list and not include_it_list else
                           "blank template: no topic or to-do text in any row")
            file_items.append(ReviewItem(FAMILY, name, "", 0, "file", stat.reason, "no_tasks"))
        todos += file_todos
        items += file_items
    _assign_ids(todos)
    return TodosResult(todos=_merge(todos), review=items, skipped_files=skipped, per_file_stats=stats)


def summarize(result: TodosResult) -> dict:
    """Counts for a CLI line, the web UI and the every-file matrix. No cell values."""
    reasons = Counter(i.reason for i in result.review)
    todos = result.todos
    return {
        "files": len(result.per_file_stats),
        "sheets": sum(s.sheets for s in result.per_file_stats),
        "read_sheets": sum(s.read_sheets for s in result.per_file_stats),
        "rows": sum(s.rows for s in result.per_file_stats),
        "blank_rows": sum(s.blank_rows for s in result.per_file_stats),
        "tasks_read": sum(s.tasks for s in result.per_file_stats),
        "tasks": len(todos),
        "done": sum(t.done for t in todos),
        "open": sum(not t.done for t in todos),
        "with_due": sum(bool(t.due_iso) for t in todos),
        "with_owner": sum(bool(t.owner) for t in todos),
        "linked_to_company": sum(bool(t.company) for t in todos),
        "needs_check": sum(bool(t.tags) for t in todos),
        "skipped_by_reason": dict(sorted((k, v) for k, v in reasons.items() if k in SKIP_REASONS)),
        "notes_by_reason": dict(sorted((k, v) for k, v in reasons.items() if k not in SKIP_REASONS)),
        "skipped_files": len(result.skipped_files),
    }
