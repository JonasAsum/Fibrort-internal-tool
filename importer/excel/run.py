"""Reads many workbooks into one list of visits. Byte-identical copies are read
once; when several files carry the same visit (owner, date, customer) the copy in
the newest file (higher CW in the name, then mtime) wins. A bad file or sheet is
recorded and the run goes on. Company matching runs only when excel/companies.py
exists."""
import hashlib
import importlib.util
import os
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from visits.reader import Parsed, Visit

from .readers import ReviewItem, read_sheet
from .sheet import load

FILE_CW = re.compile(r"CW\s*(\d{1,2})", re.I)
AUTO_OK = {"due_decimal"}  # review reasons that do not keep a visit out of the auto-accepted set
COMPANY_OK = 0.9


@dataclass
class FileStat:
    path: str
    family: str = "none"
    sheets: int = 0
    read_sheets: int = 0
    visits: int = 0
    error: str = ""
    stale_copy: bool = False


@dataclass
class RunResult:
    parsed: Parsed
    review: list
    skipped_files: list  # [{"file": path, "error": text}]
    per_file_stats: list = field(default_factory=list)
    auto_accepted: set = field(default_factory=set)  # visit_key values
    visit_family: dict = field(default_factory=dict)  # visit_key -> layout


@dataclass
class _Read:
    path: str
    layout: str
    cw: int
    mtime: float
    visit: Visit


def _fold(name: str) -> str:
    return re.sub(r"[\W_]+", "", name.lower())


def _key(visit):
    if not visit.date_iso or not visit.customer_raw:
        return None
    return (visit.owner, visit.date_iso, _fold(visit.customer_raw))


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
    return keep


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


def _company_items(kept):
    """Review items for customers the company list does not know well enough. Only
    runs when excel/companies.py exists."""
    if importlib.util.find_spec("excel.companies") is None:
        return []
    from . import companies

    match = getattr(companies, "match", None)
    if match is None:  # the module exists but its matcher is not written yet
        return []
    out = []
    for r in kept:
        name = r.visit.customer_raw
        if not name:
            continue
        m = match(name)
        if getattr(m, "score", 0) < COMPANY_OK:
            out.append((ReviewItem(r.layout, r.visit.workbook, r.visit.sheet, r.visit.excel_row_start,
                                   "customer", name, "company", list(getattr(m, "candidates", []))),
                        r.visit.visit_key))
    return out


def _file_cw(path):
    m = FILE_CW.search(os.path.basename(path))
    return int(m.group(1)) if m else 0


def _mtime(path):
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0


def read_paths(paths) -> RunResult:
    skipped, stats, reads, weeks, items = [], [], [], [], []
    for path in _unique(paths, skipped):
        stat = FileStat(path=path)
        stats.append(stat)
        try:
            sheets = load(path)
        except Exception as exc:  # a corrupt or unsupported workbook must not stop the run
            stat.error = f"{type(exc).__name__}: {exc}"
            skipped.append({"file": path, "error": stat.error})
            continue
        stat.sheets = len(sheets)
        layouts = Counter()
        for sheet in sheets:
            try:
                layout, visits, week, review = read_sheet(sheet, path)
            except Exception as exc:
                stat.error = (stat.error + "; " if stat.error else "") + f"{sheet.name}: {exc}"
                continue
            if layout is None:
                continue
            layouts[layout] += 1
            stat.read_sheets += 1
            reads += [_Read(path, layout, _file_cw(path), _mtime(path), v) for v in visits]
            weeks += [(path, week)] if week else []
            items += [(item, None) for item in review]
        stat.family = layouts.most_common(1)[0][0] if layouts else "none"

    keys = _attach([item for item, _ in items], reads)
    items = [(item, key) for (item, _), key in zip(items, keys)]

    best = {}
    for r in reads:
        key = _key(r.visit)
        if key is not None:
            rank = (r.cw, r.mtime, r.path)
            if key not in best or rank > best[key][0]:
                best[key] = (rank, r.path)
    kept = [r for r in reads if _key(r.visit) is None or best[_key(r.visit)][1] == r.path]
    kept_keys = {r.visit.visit_key for r in kept}
    items += _company_items(kept)

    by_visit = defaultdict(set)
    for item, key in items:
        if key is not None:
            by_visit[key].add(item.reason)
    accepted = {r.visit.visit_key for r in kept
                if r.visit.date_iso and r.visit.owner and r.visit.customer_raw
                and by_visit[r.visit.visit_key] <= AUTO_OK}

    per_path = Counter(r.path for r in kept)
    read_paths_set = {r.path for r in reads}
    for stat in stats:
        stat.visits = per_path[stat.path]
        stat.stale_copy = stat.path in read_paths_set and stat.visits == 0

    stale = {s.path for s in stats if s.stale_copy}
    return RunResult(
        parsed=Parsed(visits=[r.visit for r in kept],
                      weeks=[w for p, w in weeks if p not in stale]),
        review=[item for item, key in items if key is None or key in kept_keys],
        skipped_files=skipped,
        per_file_stats=stats,
        auto_accepted=accepted,
        visit_family={r.visit.visit_key: r.layout for r in kept},
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
        f = fam("unreadable")
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
