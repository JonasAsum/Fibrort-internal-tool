"""Reads many workbooks into one list of visits. Byte-identical copies are read
once; when several files carry the same visit (owner, date, customer) the copy that
would be auto-accepted wins; among those the one whose own dates match its file-name and
sheet CW (a renamed stale copy never does), then the newest file (higher CW, then mtime). A bad file or sheet is
recorded and the run goes on. Company matching runs only when excel/companies.py
exists."""
import copy
import functools
import hashlib
import os
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from datetime import date

from visits.reader import Parsed, Visit, assign_ordinals, merge_weeks

from .readers import ReviewItem, _sheet_cw, read_visit_sheet
from .split import SUFFIX
from .sheet import load
from .split import fold_key

IMPORTER_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_CW = re.compile(r"CW\s*(\d{1,2})", re.I)
AUTO_OK = {"due_decimal", "week_label_differs"}  # review reasons that do not keep a visit out of the auto-accepted set (a guessed month does: a wrong date is wrong data)
FILE_CACHE_MAX = 80  # files kept parsed between reads
UNRECOGNISED_MIN_CELLS = 10  # fewer cells than this is an empty template sheet, not a lost visit list
COMPANY_OK = 0.9


@dataclass
class RunFileStat:
    path: str
    family: str = "none"
    sheets: int = 0
    read_sheets: int = 0
    visits: int = 0
    error: str = ""
    unrecognised: list = field(default_factory=list)  # sheets with data but no visit layout, in a file where others were read
    sheet_errors: list = field(default_factory=list)  # [(sheet name, reason)]: sheets that raised while being read
    stale_copy: bool = False
    accepted: int = 0  # kept visits that go in without a check


@dataclass
class RunResult:
    parsed: Parsed
    review: list
    skipped_files: list  # [{"file": path, "error": text}]
    per_file_stats: list = field(default_factory=list)
    auto_accepted: set = field(default_factory=set)  # visit_key values
    visit_family: dict = field(default_factory=dict)  # visit_key -> layout
    review_by_visit: dict = field(default_factory=dict)  # visit_key -> set of review reasons
    # visits left out because another file or sheet holds the same one:
    # [{"date", "customer", "ignored", "ignored_sheet", "used", "used_sheet"}]
    superseded: list = field(default_factory=list)


@dataclass
class _Read:
    path: str
    layout: str
    cw: int
    mtime: float
    visit: Visit
    sheet: str = ""
    sheet_cw: int | None = None  # None: the sheet name carries no week (or does not count, as in the lucas layout)
    eligible: bool = False  # would be auto-accepted: has owner, date, customer and no blocking review reason

    @property
    def week_ok(self) -> bool:
        """The visit's date lies in the week the file name and the sheet say."""
        if not self.visit.date_iso:
            return False
        week = date.fromisoformat(self.visit.date_iso).isocalendar()[1]
        return all(week == cw for cw in (self.cw or None, self.sheet_cw) if cw)


def _key(r):
    """What makes two reads the same visit: owner, day, customer and which of that day's visits it is."""
    visit = r.visit
    if not visit.date_iso or not visit.customer_raw:
        return None
    return (visit.owner, visit.date_iso, fold_key(visit.customer_raw), visit.key_ordinal)


def _winner(candidates):
    return max(candidates, key=lambda r: (r.eligible, r.week_ok, r.cw, r.mtime, r.path, r.sheet))


def _unique(paths, skipped):
    """One path per distinct file content, preferring a copy under Weekly reports/."""
    groups: dict[str, list[str]] = {}
    for p in paths:
        try:
            with open(p, "rb") as f:
                digest = hashlib.sha256(f.read()).hexdigest()
        except OSError as exc:
            skipped.append({"file": p, "error": f"unreadable: {exc}"})
            continue
        groups.setdefault(digest, []).append(p)
    keep = []
    for members in groups.values():
        members.sort(key=lambda p: ("Weekly reports" not in p, p))
        keep.append(members[0])
        for dup in members[1:]:
            skipped.append({"file": dup, "error": f"identical to {os.path.basename(members[0])}"})
    # The file name is part of every record's source id; two different files with one name
    # would overwrite each other's visits, so the second waits until it is renamed.
    seen, unique = set(), []
    for p in sorted(keep, key=lambda p: ("Weekly reports" not in p, p)):
        name = os.path.basename(p)
        if name in seen:
            skipped.append({"file": p, "error": "same file name as another file with different content: rename it to import"})
            continue
        seen.add(name)
        unique.append(p)
    return unique


def _attach(items, reads):
    """visit_key of the visit each review item belongs to (None when it has none)."""
    ranges = defaultdict(list)
    for r in reads:
        ranges[(r.visit.workbook, r.visit.sheet)].append(
            (r.visit.excel_row_start, r.visit.excel_row_end, r.visit.visit_key))
    keys = []
    for item in items:
        keys.append(next((key for start, end, key in ranges.get((item.file, item.sheet), ())
                          if start <= item.row <= end), None))
    return keys


@functools.lru_cache(maxsize=None)
def _file_master(importer_dir: str):
    """The company master read from the customer list files, loaded once per process (the fallback)."""
    from .companies import Master
    return Master.load(importer_dir)


_CRM = {"master": None, "note": ""}


def set_crm_master(master, note: str = "") -> None:
    """The master read from the CRM for the upload in hand (None: use the files), and why when it is the files."""
    _CRM["master"], _CRM["note"] = master, note


def master_note() -> str:
    return _CRM["note"]


def _master(importer_dir: str):
    """The company master: the CRM's companies when the upload was matched against them, else the files."""
    return _CRM["master"] or _file_master(importer_dir)


_master.cache_clear = lambda: (_file_master.cache_clear(), set_crm_master(None))  # what the tests reset between runs
_master.cache_info = lambda: _file_master.cache_info()


def _name_parts(name: str) -> list[str]:
    """The name as written, then its Chinese part and its Latin part: a bilingual name "Baoji Geely Engine Co., Ltd.
    宝鸡吉利发动机有限公司" is one name for the CRM, but the company list holds each language on its own."""
    cjk = " ".join(re.sub(r"[A-Za-z][A-Za-z0-9 .,&'\-]*", " ", name).split())
    lat = " ".join(re.sub(r"[\u4e00-\u9fff（）]+", " ", name).split())
    parts = [name]
    m = SUFFIX.search(name)
    head = name[:m.end()].strip() if m and name[m.end():].strip() else ""  # "… Co., Ltd." of "… Co., Ltd. Xi'an Factory"
    for p in (head, cjk, lat):
        if p and p != name and len(re.findall(r"[\u4e00-\u9fffA-Za-z]", p)) >= 3:
            parts.append(p)
    return parts


def best_match(master, name: str, name_en: str, code: str | None = None):
    """The company-list match for a cell. The name as written decides unless one of its language parts finds a
    real company (a hit, not "new" or a question) with a higher score. The ERP code, when there is one, is added
    to each try and finds the company exactly."""
    tries = _name_parts(name)
    if name_en and name_en not in tries:
        tries.append(name_en)
    results = [master.match(f"{code} {t}" if code else t) for t in tries]
    hits = [m for m in results if m.method not in ("new", "review", "ambiguous") and m.score >= COMPANY_OK]
    if not hits:
        return results[0]
    return master.place_check(max(hits, key=lambda m: m.score), f"{name} {name_en}")


def _company_items(kept, master=None):
    """Review items for customers the company list does not know well enough, and
    the match itself in visit.extra["company_match"] for the ones it does."""
    if master is None:
        master = _master(IMPORTER_DIR)
    out = []
    for r in kept:
        name = r.visit.customer_raw
        if not name or r.visit.extra.get("kind"):  # internal and exhibition rows are not customers
            continue
        code = r.visit.extra.get("customer_code")
        m = best_match(master, name, r.visit.extra.get("name_en") or "", code)
        r.visit.extra["company_match"] = {"code": m.code, "cn": m.cn, "en": m.en,
                                          "score": m.score, "method": m.method}
        if m.score < COMPANY_OK:
            reason = "company"  # a person answers "same company?" (excel/decisions.py) before the visit goes in
            out.append((ReviewItem(r.layout, r.visit.workbook, r.visit.sheet, r.visit.excel_row_start,
                                   "customer", name, reason, [list(c) for c in (m.candidates or [])]),
                        r.visit.visit_key))
    return out + _pair_items(kept)


def _pair_items(kept):
    """Names in this upload that no customer in the list explains but that look like one company
    written two ways: merged when a person said "same", asked about otherwise."""
    from .companies import lookalike, norm
    from .decisions import DecisionStore, default_path
    unknown = {}  # normalised name -> (name as written, its rows)
    for r in kept:
        name = r.visit.customer_raw
        if not name or r.visit.extra.get("kind"):
            continue
        m = r.visit.extra.get("company_match") or {}
        if m.get("code") is None and m.get("method") == "new":
            unknown.setdefault(norm(name) or name, (name, []))[1].append(r)
    store = DecisionStore(default_path(IMPORTER_DIR))
    names = sorted(unknown.values(), key=lambda t: (-len(t[0]), t[0]))  # the fuller spelling first
    out, done = [], set()
    for i, (full, full_rows) in enumerate(names):
        for other, other_rows in names[i + 1:]:
            if (norm(full), norm(other)) in done or not lookalike(full, other):
                continue
            same = store.pair_decision(full, other)
            if same:
                for r in other_rows:
                    r.visit.extra["customer_as_written"] = r.visit.customer_raw
                    r.visit.customer_raw = full
            elif same is None:
                for rows, mine, theirs in ((full_rows, full, other), (other_rows, other, full)):
                    for r in rows:
                        out.append((ReviewItem(r.layout, r.visit.workbook, r.visit.sheet, r.visit.excel_row_start,
                                               "customer", mine, "company_pair", [[theirs]]), r.visit.visit_key))
    return out


def _file_cw(path):
    m = FILE_CW.search(os.path.basename(path))
    return int(m.group(1)) if m else 0


def _mtime(path):
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0


def _plain_error(exc, path, sheet_name) -> str:
    """The exception's message without the "file/sheet:" prefix the readers put on it."""
    text = str(exc) or type(exc).__name__
    prefix = f"{os.path.basename(path)}/{sheet_name}: "
    return text[len(prefix):] if text.startswith(prefix) else text


_FILE_CACHE: dict = {}  # (path, mtime, size) -> what one file read to; a company answer re-runs only the stages after it


def _read_file(path):
    """(RunFileStat, reads, weeks, review items) of one workbook. The result depends only on the file, so a
    re-read of an unchanged file (after a company answer, say) is a copy of the last one, not another parse."""
    try:
        st = os.stat(path)
        key = (path, st.st_mtime_ns, st.st_size)
    except OSError:
        key = None
    if key in _FILE_CACHE:
        return copy.deepcopy(_FILE_CACHE[key])
    got = _parse_file(path)
    if key is not None:
        if len(_FILE_CACHE) >= FILE_CACHE_MAX:
            _FILE_CACHE.pop(next(iter(_FILE_CACHE)))
        _FILE_CACHE[key] = copy.deepcopy(got)
    return got


def _parse_file(path):
    stat = RunFileStat(path=path)
    reads, weeks, items = [], [], []
    try:
        sheets = load(path)
    except Exception as exc:  # a corrupt or unsupported workbook must not stop the run
        stat.error = f"{type(exc).__name__}: {exc}"
        return stat, reads, weeks, items
    stat.sheets = len(sheets)
    layouts = Counter()
    for sheet in sheets:
        if sheet.cut:
            stat.sheet_errors.append((sheet.name, sheet.cut))
        try:
            layout, visits, week, review = read_visit_sheet(sheet, path)
        except Exception as exc:
            stat.error = (stat.error + "; " if stat.error else "") + f"{sheet.name}: {exc}"
            stat.sheet_errors.append((sheet.name, _plain_error(exc, path, sheet.name)))
            continue
        if layout is None:
            if len(sheet.grid) >= UNRECOGNISED_MIN_CELLS:
                stat.unrecognised.append(sheet.name)
            continue
        layouts[layout] += 1
        stat.read_sheets += 1
        assign_ordinals(visits)
        sheet_cw = None if layout == "lucas" else _sheet_cw(sheet.name)
        reads += [_Read(path, layout, _file_cw(path), _mtime(path), v, sheet.name, sheet_cw) for v in visits]
        weeks += [(path, week)] if week else []
        items += [(item, None) for item in review]
    stat.family = layouts.most_common(1)[0][0] if layouts else "none"
    return stat, reads, weeks, items


def read_paths(paths) -> RunResult:
    skipped, stats, reads, weeks, items = [], [], [], [], []
    for path in _unique(paths, skipped):
        stat, file_reads, file_weeks, file_items = _read_file(path)
        stats.append(stat)
        if stat.error and not stat.sheets:
            skipped.append({"file": path, "error": f"could not be read: {stat.error}"})
        reads += file_reads
        weeks += file_weeks
        items += file_items

    keys = _attach([item for item, _ in items], reads)
    items = [(item, key) for (item, _), key in zip(items, keys)]

    # Reasons known before company matching decide which copy of a visit may be the one kept.
    early = defaultdict(set)
    for item, key in items:
        if key is not None:
            early[key].add(item.reason)
    for r in reads:
        v = r.visit
        r.eligible = bool(v.date_iso and v.owner and v.customer_raw and early[v.visit_key] <= AUTO_OK)
    groups = defaultdict(list)
    for r in reads:
        if _key(r) is not None:
            groups[_key(r)].append(r)
    best = {key: _winner(group) for key, group in groups.items()}
    kept = [r for r in reads if _key(r) is None or best[_key(r)] is r]
    kept_keys = {r.visit.visit_key for r in kept}
    superseded = [{"date": r.visit.date_iso, "customer": r.visit.customer_raw,
                   "ignored": os.path.basename(r.path), "ignored_sheet": r.sheet,
                   "used": os.path.basename(best[_key(r)].path), "used_sheet": best[_key(r)].sheet}
                  for r in reads if _key(r) is not None and best[_key(r)] is not r]
    items += _company_items(kept)

    by_visit = defaultdict(set)
    for item, key in items:
        if key is not None:
            by_visit[key].add(item.reason)
    accepted = {r.visit.visit_key for r in kept
                if r.visit.date_iso and r.visit.owner and r.visit.customer_raw
                and by_visit[r.visit.visit_key] <= AUTO_OK}

    per_path = Counter(r.path for r in kept)
    ok_per_path = Counter(r.path for r in kept if r.visit.visit_key in accepted)
    read_paths_set = {r.path for r in reads}
    for stat in stats:
        stat.visits = per_path[stat.path]
        stat.accepted = ok_per_path[stat.path]
        stat.stale_copy = stat.path in read_paths_set and stat.visits == 0

    kept_sheets = {(r.path, r.sheet) for r in kept}
    stale = {(r.path, r.sheet) for r in reads} - kept_sheets
    return RunResult(
        parsed=Parsed(visits=[r.visit for r in kept],
                      weeks=merge_weeks([w for p, w in weeks if (p, w.sheet) not in stale])),
        review=[item for item, key in items if key is None or key in kept_keys],
        skipped_files=skipped,
        per_file_stats=stats,
        auto_accepted=accepted,
        visit_family={r.visit.visit_key: r.layout for r in kept},
        review_by_visit={k: v for k, v in by_visit.items() if k in kept_keys},
        superseded=superseded,
    )


def summarize(result: RunResult) -> dict:
    """Per-family counts: files, sheets, visits, auto-accepted, queued by reason, skipped files."""
    fams: dict[str, dict] = {}

    def fam(name):
        return fams.setdefault(name, {"files": 0, "sheets": 0, "read_sheets": 0, "visits": 0,
                                      "auto_accepted": 0, "queued": 0, "queued_by_reason": Counter(),
                                      "skipped_files": 0, "errors": []})

    for s in result.per_file_stats:
        f = fam(s.family)
        f["files"] += 1
        f["sheets"] += s.sheets
        f["read_sheets"] += s.read_sheets
        if s.error:
            f["errors"].append(os.path.basename(s.path))
    for s in result.skipped_files:
        f = fam("duplicate" if s["error"].startswith("identical to") else "unreadable")
        f["skipped_files"] += 1
        f["errors"].append(os.path.basename(s["file"]))
    for v in result.parsed.visits:
        f = fam(result.visit_family.get(v.visit_key, "none"))
        f["visits"] += 1
        if v.visit_key in result.auto_accepted:
            f["auto_accepted"] += 1
        else:
            f["queued"] += 1
    for item in result.review:
        fam(item.family)["queued_by_reason"][item.reason] += 1
    for f in fams.values():
        f["queued_by_reason"] = dict(f["queued_by_reason"])
    return fams
