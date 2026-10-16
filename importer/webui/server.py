"""Local web UI for both importers: visit-list workbooks (visits/) and the email
archive (mail/). Upload, review what was matched, preview, then commit - without
typing the CLI modes by hand. Single-user, local-only; state lives in process
memory (CURRENT, MAIL) and is lost on restart, same as a CLI run.

    cd importer && .venv/bin/python -m webui.server      # http://localhost:5055
"""
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
from collections import Counter, defaultdict
from dataclasses import asdict
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)  # importer/
if ROOT not in sys.path:  # also allow `python webui/server.py`
    sys.path.insert(0, ROOT)

from common.client import ApiError, MargenceClient  # noqa: E402
from visits import pipeline  # noqa: E402
from visits.match import match  # noqa: E402
from excel.readers import REPS  # noqa: E402
from excel import run as excel_run  # noqa: E402
from excel.companies import norm as company_norm  # noqa: E402
from excel.decisions import DecisionStore, default_path  # noqa: E402
from excel.run import read_paths  # noqa: E402
from excel_margince import accepted  # noqa: E402
from webui import checks as site_checks, entry, families, selftest  # noqa: E402

app = Flask(__name__, static_folder="static")

# IMPORTER_CONFIG / IMPORTER_UPLOAD_DIR let a test server keep its own state apart from the real one.
CONFIG_PATH = os.environ.get("IMPORTER_CONFIG") or os.path.join(HERE, ".webui_config.json")
UPLOAD_DIR = os.environ.get("IMPORTER_UPLOAD_DIR") or os.path.join(HERE, ".uploads")

DEFAULT_CONFIG = {"base_url": "http://localhost:8080", "email": "admin@demo.test", "password": "demo-password-123"}

CURRENT = {"files": [], "parsed": None, "result": None, "run": None, "paths": []}  # the visit reports (01-03)
# Every other list family of the upload: {family: {"paths", "full" (all it read), "left_out" (flag ids), "result"}}
BATCH = {"entry": [], "fams": {}}
JOBS: dict[str, dict] = {}
COMMIT_LOCK = threading.Lock()  # one import into the CRM at a time
UPLOAD_LOCK = threading.Lock()  # one upload at a time: they share the upload folder and CURRENT


def _load_config() -> dict:
    if os.path.isfile(CONFIG_PATH):
        with open(CONFIG_PATH) as f:
            return {**DEFAULT_CONFIG, **json.load(f)}
    return dict(DEFAULT_CONFIG)


def _save_config(cfg: dict) -> None:
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f)


def _client() -> MargenceClient:
    cfg = _load_config()
    c = MargenceClient(cfg["base_url"])
    c.login(cfg["email"], cfg["password"])
    return c


SKIP_DIRS = {".venv", ".git", "__pycache__", "node_modules", ".uploads"}


HEAD_TTL = 30  # seconds the git HEAD is reused (unless HEAD or its branch moved): every open page polls /api/version
_HEAD_CACHE: dict = {}


def _head_stamp():
    """Modification times of the files a commit or checkout rewrites (cheap stats, no git call)."""
    here = ROOT
    while not os.path.exists(os.path.join(here, ".git")):
        if os.path.dirname(here) == here:
            return None
        here = os.path.dirname(here)
    git_dir = os.path.join(here, ".git")
    if os.path.isfile(git_dir):
        return None
    stamp = []
    try:
        with open(os.path.join(git_dir, "HEAD")) as f:
            ref = f.read().strip()
        names = ["HEAD", "packed-refs"] + ([ref[5:].strip()] if ref.startswith("ref:") else [])
        for name in names:
            path = os.path.join(git_dir, name)
            stamp.append(os.path.getmtime(path) if os.path.exists(path) else None)
    except OSError:
        return None
    return tuple(stamp)


def _git_head() -> str:
    now, stamp = time.monotonic(), _head_stamp()
    if "at" in _HEAD_CACHE and now - _HEAD_CACHE["at"] < HEAD_TTL and _HEAD_CACHE["stamp"] == stamp:
        return _HEAD_CACHE["head"]
    try:
        head = subprocess.run(["git", "-C", ROOT, "rev-parse", "HEAD"], capture_output=True, text=True,
                              timeout=5).stdout.strip()
    except (OSError, subprocess.SubprocessError):
        head = ""
    _HEAD_CACHE.update(at=now, head=head, stamp=stamp)
    return head


def _code_state() -> dict:
    """The checkout's git HEAD and the newest mtime of importer/**/*.py right now (tests/ do not count)."""
    head = _git_head()
    newest = 0.0
    for dirpath, dirs, names in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith("staging")
                   and not (dirpath == ROOT and d == "tests")]
        for n in names:
            if n.endswith(".py"):
                try:
                    newest = max(newest, os.path.getmtime(os.path.join(dirpath, n)))
                except OSError:
                    pass
    return {"head": head, "newest_py_mtime": newest}


STARTED = _code_state()  # what this process was started from


@app.get("/api/version")
def version():
    """Whether the files on disk moved on after this server started (it would run older code)."""
    now = _code_state()
    return jsonify({"started": STARTED, "current": now, "stale": now != STARTED})


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.get("/api/config")
def get_config():
    cfg = _load_config()
    return jsonify({"base_url": cfg["base_url"], "email": cfg["email"]})


@app.post("/api/config")
def set_config():
    body = request.get_json(force=True)
    cfg = _load_config()
    for key in ("base_url", "email", "password"):
        if body.get(key):
            cfg[key] = body[key]
    _save_config(cfg)
    return jsonify({"ok": True})


@app.post("/api/login_test")
def login_test():
    try:
        c = _client()
        me = c.get("/v1/me")
        return jsonify({"ok": True, "workspace_name": me.get("workspace_name"), "user": me["user"]["display_name"]})
    except ApiError as e:
        return jsonify({"ok": False, "error": str(e)}), 200
    except Exception as e:  # connection refused, DNS, etc.
        return jsonify({"ok": False, "error": str(e)}), 200


def _checks_context() -> dict:
    return {"load_config": _load_config, "make_client": _client, "upload_dir": UPLOAD_DIR, "config_path": CONFIG_PATH,
            "decisions_path": Path(default_path(excel_run.IMPORTER_DIR)),
            "master": lambda: excel_run._master(excel_run.IMPORTER_DIR), "stale": _code_state() != STARTED}


@app.get("/api/checks")
def api_checks():
    """The Checks page: can this site work right now, and do the readers still read? Writes nothing."""
    return jsonify({"environment": site_checks.environment(_checks_context()),
                    "selftest": site_checks.selftest_rows(selftest.run())})


def _dataset_summary() -> dict:
    parsed, result = CURRENT["parsed"], CURRENT["result"]
    if result is None:
        return {"loaded": False}
    visits = parsed.visits
    no_date = [v.visit_key for v in visits if not v.date_iso]
    counts = pipeline.plan_counts(parsed, result)
    return {
        "loaded": True,
        "files": CURRENT["files"],
        "visits": len(visits),
        "companies": len(result.companies),
        "contacts": len(result.contacts),
        "tasks": counts["tasks"],
        "weeks": counts["weeks"],
        "counts": counts,
        "no_date": len(no_date),
        "no_date_samples": no_date[:10],
        "company_rows": sorted(
            [{"name": c.display_name, "city": c.city, "dealer": c.is_dealer, "potential": c.potential,
              "type": c.customer_type, "reps": ", ".join(c.owners), "visits": c.visit_count,
              "last_visit": c.last_visit} for c in result.companies.values()],
            key=lambda r: r["name"]),
        "contact_rows": sorted(
            [{"full_name": c.full_name, "company": result.companies[c.company_key].display_name,
              "phones": ", ".join(c.phones + c.office_phones), "emails": ", ".join(c.emails)}
             for c in result.contacts.values() if c.company_key in result.companies],
            key=lambda r: (r["company"], r["full_name"])),
        "decisions": [d for d in result._namer.decisions if d["why"].startswith("kept separate")],
        **_unsure_summary(CURRENT.get("run")),
    }


# Plain words for the review-queue reasons that keep a visit out of the import.
REASON_TEXT = {
    "week_mismatch": "the visit dates are not in the week and year the file or sheet is named after (often a copy of an older report)",
    "no_year": "the file name has no year and the sheet has no week number to work it out from",
    "no_date": "no readable date",
    "ambiguous_date": "the date could be day/month or month/day",
    "range_date": "the date is a range",
    "no_customer": "no customer",
    "no_owner": "no sales rep",
    "company_pair": "two spellings that may be one company: answer \"Same company\" or \"Different companies\" below",
    "company": "not sure which company this is: answer \"New company\" or \"Existing company\" below",
    "customer_split": "unsure where the customer name ends and the address starts",
    "participants": "the participants could not be matched to the company",
}


def _file_problems(run) -> list[dict]:
    """Plain-language notes on every file or sheet that did not (fully) become visits."""
    out, named = [], set()
    for x in run.skipped_files:
        named.add(x["file"])
        out.append({"file": os.path.basename(x["file"]), "why": x["error"]})
    for st in run.per_file_stats:
        name = os.path.basename(st.path)
        if st.path in named:
            continue
        for sheet, why in st.sheet_errors:
            out.append({"file": name, "why": f"sheet '{sheet}' could not be read: {why}"})
        if st.read_sheets:
            out += [{"file": name, "why": f"sheet '{sheet}' was not used: no visit list recognised in it"}
                    for sheet in st.unrecognised]
        if st.read_sheets == 0 and not st.sheet_errors:
            out.append({"file": name, "why": "no visit list recognised (no sheet has the Date / Customer columns "
                                              "of a weekly report, a weekly plan or a visit-in-detail form)"})
    pairs = Counter((x["ignored"], x["ignored_sheet"], x["used"], x["used_sheet"]) for x in run.superseded)
    for (ignored, ignored_sheet, used, used_sheet), n in pairs.items():
        where = (f"sheet '{used_sheet}' of this file" if used == ignored else f"{used} (sheet '{used_sheet}')")
        out.append({"file": ignored, "why": f"{n} visit(s) in sheet '{ignored_sheet}' ignored: the same visit is in "
                                             f"{where}, which was used instead"})
    return out


def _company_questions(run) -> list[dict]:
    """One question per distinct customer name that is not clearly a known company or a new one."""
    master = excel_run._master(excel_run.IMPORTER_DIR)
    asked: dict[str, dict] = {}
    for item in run.review:
        if item.field == "customer" and item.reason == "company_pair":
            other = item.candidates[0][0]
            a, b = sorted((item.value, other), key=lambda n: (-len(n), n))  # the fuller spelling first
            q = asked.setdefault("pair:" + "|".join(sorted((company_norm(a), company_norm(b)))),
                                 {"kind": "pair", "a": a, "b": b, "visits": 0})
            q["visits"] += 1
            continue
        if item.field != "customer" or item.reason != "company":
            continue
        q = asked.setdefault(company_norm(item.value) or item.value, {
            "kind": "company", "name": item.value, "visits": 0, "candidates": [
                {"code": c[0], "name": c[1], "en": (master.get(c[0]) or {}).get("en", ""), "score": round(c[2], 2),
                 "uscc": (master.get(c[0]) or {}).get("uscc", ""), "legal": (master.get(c[0]) or {}).get("legal", "")}
                for c in (item.candidates or [])[:3]]})
        q["visits"] += 1
    return [families.with_suggestion(q, master) if q["kind"] == "company" else q for q in asked.values()]


def _unsure_summary(run) -> dict:
    """Visits found in the file but left out of the import until someone checks them."""
    if run is None:
        return {"found": 0, "unsure": 0, "unsure_reasons": {}, "unsure_rows": [], "skipped_files": []}
    reasons_by_visit = run.review_by_visit
    unsure = [v for v in run.parsed.visits if v.visit_key not in run.auto_accepted]
    counts = {}
    rows = []
    for v in unsure:
        reasons = sorted(reasons_by_visit.get(v.visit_key, set()) - {"due_decimal"}) or ["unsure"]
        for r in reasons:
            counts[REASON_TEXT.get(r, r)] = counts.get(REASON_TEXT.get(r, r), 0) + 1
        rows.append({"visit": v.visit_key, "date": v.date_iso or v.date_raw, "customer": v.customer_raw,
                     "file": os.path.basename(v.workbook), "sheet": v.sheet, "row": v.excel_row_start,
                     "needs_excel_fix": bool(set(reasons) - {"company", "company_pair"}),
                     "why": "; ".join(REASON_TEXT.get(r, r) for r in reasons)})
    rows.sort(key=lambda r: not r["needs_excel_fix"])  # the ones only a corrected file can clear come first
    return {"found": len(run.parsed.visits), "unsure": len(unsure), "unsure_reasons": counts,
            "needs_excel_fix": sum(r["needs_excel_fix"] for r in rows), "unsure_rows": rows[:50], "company_questions": _company_questions(run),
            "skipped_files": _file_problems(run)}


@app.get("/api/dataset")
def get_dataset():
    return jsonify(_dataset_summary())


EXCEL_TYPES = (".xlsx", ".xlsm", ".xls")


def _clear_uploads():
    """Empties the upload folder of workbooks (each upload lives in its own f<n> subfolder)."""
    for kind in EXCEL_TYPES:
        for stale in glob.glob(os.path.join(UPLOAD_DIR, "*" + kind)):
            os.remove(stale)
    for sub in glob.glob(os.path.join(UPLOAD_DIR, "f[0-9]*")):
        shutil.rmtree(sub, ignore_errors=True)


def _forget():
    CURRENT.update(files=[], parsed=None, result=None, run=None, paths=[])
    BATCH.update(entry=[], fams={})


def _folder_of(filename: str) -> str:
    """The folder a browser reported for the file ("Jason Zhao/x.xlsx" -> "Jason Zhao"), made safe as a name."""
    parts = [p for p in filename.replace("\\", "/").split("/") if p not in ("", ".", "..")]
    return re.sub(r"[^\w .-]", "_", parts[-2]).lstrip(".") if len(parts) > 1 else ""


@app.get("/api/reps")
def reps():
    return jsonify({"reps": REPS})


@app.post("/api/upload")
def upload():
    with UPLOAD_LOCK:
        _forget()  # whatever happens below, nothing from an earlier upload may be imported by mistake
        return _upload()


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def _upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "no files in the request"}), 400
    rep = (request.form.get("rep") or "").strip()
    if rep and rep not in REPS:
        return jsonify({"error": f"unknown sales rep: {rep}"}), 400

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    _clear_uploads()

    saved = []  # (saved path, file name, the path the browser reported)
    for i, f in enumerate(files):
        name = os.path.basename(f.filename.replace("\\", "/"))
        if not name.lower().endswith(EXCEL_TYPES):
            _clear_uploads()
            return jsonify({"error": f"{name} is not an Excel file (.xlsx, .xlsm or .xls)"}), 400
        # Its own folder per file: two files with one name never overwrite each other, and the
        # folder the browser reported (or the rep chosen on the page) names the sales rep.
        folder = os.path.join(UPLOAD_DIR, f"f{i}", rep or _folder_of(f.filename))
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, name)
        f.save(path)
        saved.append((path, name, f.filename))

    try:
        _classify(saved)
        _read_everything()
    except Exception as e:
        _forget()
        return jsonify({"error": f"could not read the workbook(s): {e}"}), 400
    return jsonify(_batch_response())


def _classify(saved):
    """BATCH["entry"]: what every uploaded workbook is, by its content. A byte-identical copy is named, not read twice."""
    seen, out = {}, []
    for path, _, reported in saved:
        e = entry.classify(path, entry.folder_hint(reported))
        digest = _sha256(path)
        e["path"], e["duplicate_of"] = path, seen.get(digest, "")
        seen.setdefault(digest, e["display_name"])
        out.append(e)
    BATCH["entry"] = out


def _refresh_master():
    """Matches this upload against the companies the CRM holds now. Without a usable CRM the customer list files
    are the master, and the note says why (shown under Checks)."""
    from excel import crm_master
    try:
        master = crm_master.load(_client, excel_run.IMPORTER_DIR)
    except Exception as e:  # not reachable, not logged in, an older CRM without the endpoint
        excel_run.set_crm_master(None, f"the CRM could not be read ({str(e)[:120]}); matching against the customer list files")
        return
    if not master.size:
        excel_run.set_crm_master(None, "the CRM holds no companies yet; matching against the customer list files")
    else:
        excel_run.set_crm_master(master)


def _read_everything():
    """Reads each family's files with that family's reader. One family failing never hides another."""
    _refresh_master()
    by_family: dict[str, list[str]] = {}
    for e in BATCH["entry"]:
        if e["kind"] == "import" and not e["duplicate_of"]:
            by_family.setdefault(e["family"], []).append(e["path"])
    if by_family.get("visits"):
        _read_visits(by_family["visits"])
    for key, paths in by_family.items():
        if key != "visits":
            BATCH["fams"][key] = {"paths": paths, "left_out": set(families.ADAPTERS[key].left_out_by_default), "full": None,
                                   "result": None, "error": ""}
            _read_family(key)
    CURRENT["files"] = [os.path.basename(p) for p in CURRENT["paths"]]


def _read_visits(paths):
    run = read_paths(paths)
    parsed = accepted(run)
    CURRENT.update(paths=paths, run=run, parsed=parsed, result=match(parsed.visits))


def _read_family(key):
    f, adapter = BATCH["fams"][key], families.ADAPTERS[key]
    try:
        f["full"] = adapter.read(f["paths"])
        f["result"] = adapter.apply(f["full"], f["left_out"])
        f["error"] = ""
    except Exception as e:  # shown on that family's panel; the other families are unaffected
        f["full"] = f["result"] = None
        f["error"] = f"{type(e).__name__}: {e}"


def _visits_ready() -> int:
    """How many visits a commit would write (0 = nothing to do, e.g. a blank template)."""
    parsed = CURRENT["parsed"]
    return len(parsed.visits) if CURRENT["result"] is not None and parsed is not None else 0


def _family_ready(key) -> int:
    if key == "visits":
        return _visits_ready()
    f = BATCH["fams"].get(key)
    return families.ADAPTERS[key].writable(f["result"]) if f and f["result"] is not None else 0


def _batch_loaded() -> bool:
    return CURRENT["run"] is not None or bool(BATCH["fams"])


def _file_records() -> tuple[dict, dict, dict]:
    """({saved path: (records, unit)}, {saved path: why it was not read},
    {saved path: (status, words) for a file another file covers wholly or partly})."""
    records, skipped, covered = {}, {}, {}
    run = CURRENT["run"]
    if run is not None:
        records.update({st.path: (st.visits, "visits") for st in run.per_file_stats})
        used = defaultdict(set)
        for x in run.superseded:
            used[x["ignored"]].add(x["used"])
        for st in run.per_file_stats:
            if st.stale_copy:
                by = sorted(used.get(os.path.basename(st.path), ()))
                covered[st.path] = ("superseded", "Every visit in it is also in " + (", ".join(by) if by else "another file of this upload")
                                    + ", which is used instead. Nothing extra is written from it.")
            elif st.visits:
                waiting = st.visits - st.accepted
                covered[st.path] = ("ready", f"{st.visits} visit(s) found: {st.accepted} go in"
                                    + (f", {waiting} wait for a check (see the list below)" if waiting else "") + ".")
        skipped.update({x["file"]: x["error"] for x in run.skipped_files})
    for key, f in BATCH["fams"].items():
        if f["result"] is not None:
            records.update(families.ADAPTERS[key].file_stats(f["result"]))
            covered.update(families.ADAPTERS[key].file_notes(f["result"]))
            skipped.update(families.ADAPTERS[key].skipped(f["result"]))
    return records, skipped, covered


def _entry_view(e, records, skipped, covered) -> dict:
    """One row of the front door's table: what the file is, where it goes in the CRM, what will happen."""
    n, unit = records.get(e["path"], (0, ""))
    out = {k: e[k] for k in ("display_name", "kind", "family", "folder", "what", "layout", "target", "reason",
                             "sheets", "sheets_used", "duplicate_of")}
    out.update(records=n, unit=unit, outcome=e["outcome"], status=e["kind"])
    f = BATCH["fams"].get(e["family"]) if e["family"] else None
    if e["kind"] != "import":
        return out
    if e["duplicate_of"]:
        out.update(status="duplicate", outcome=f"Identical copy of {e['duplicate_of']}: read once, nothing extra.")
    elif e["family"] != "visits" and (f is None or f["error"]):
        out.update(status="not_read", outcome=f"Could not be read: {f['error'] if f else 'unknown error'}")
    elif e["path"] in skipped:
        out.update(status="not_read", outcome=f"Not read: {skipped[e['path']]}")
    elif e["path"] in covered and covered[e["path"]][0] == "superseded":
        out.update(status="superseded", outcome=covered[e["path"]][1])
    elif n == 0:
        out.update(status="empty", outcome=f"Recognised as {e['what']}, but it holds nothing to import "
                                           "(a blank template?). Nothing is written from it.")
    else:
        note = covered.get(e["path"], ("", ""))[1]
        out.update(status="ready", outcome=(note + " " if note else f"{n} {unit} found. ") + f"They go into the CRM as: {e['target']}.")
    return out


def _family_view(key, records) -> dict:
    label, target, _ = entry.FAMILIES[key]
    names = [e["display_name"] for e in BATCH["entry"] if e["kind"] == "import" and e["family"] == key and not e["duplicate_of"]]
    view = {"key": key, "label": label, "target": target, "files": names, "ready": _family_ready(key), "waiting": 0,
            "counts": [], "notes": [], "questions": [], "flags": [], "error": ""}
    if key == "visits":
        return view
    f, adapter = BATCH["fams"][key], families.ADAPTERS[key]
    view["error"] = f["error"]
    if f["result"] is None:
        return view
    counts = adapter.flag_counts(f["full"])
    view.update(waiting=adapter.waiting(f["result"]), counts=adapter.counts(f["result"]), notes=adapter.notes(f["result"]),
                questions=adapter.questions(f["result"]),
                flags=[{"id": fid, "label": label_, "help": help_, "count": counts.get(fid, 0), "included": fid not in f["left_out"]}
                       for fid, label_, help_ in adapter.flags if counts.get(fid, 0)])
    return view


def _batch_summary() -> dict:
    records, skipped, covered = _file_records()
    views = [_family_view(k, records) for k in entry.COMMIT_ORDER
             if any(e["kind"] == "import" and e["family"] == k for e in BATCH["entry"])]
    return {"entry": [_entry_view(e, records, skipped, covered) for e in BATCH["entry"]], "families": views,
            "ready_families": [v["key"] for v in views if v["ready"]]}


def _batch_response() -> dict:
    """The visit summary (its keys stay at the top, as before) plus the whole upload under "batch"."""
    return {**_dataset_summary(), "batch": _batch_summary()}


@app.get("/api/batch")
def get_batch():
    return jsonify(_batch_response())


def _asked_names() -> set:
    """Every customer name some family is asking about (a pair of spellings is answered separately)."""
    names = set()
    if CURRENT["run"] is not None:
        names |= {q["name"] for q in _company_questions(CURRENT["run"]) if q.get("name")}
    for key, f in BATCH["fams"].items():
        if f["result"] is not None:
            names |= {q["name"] for q in families.ADAPTERS[key].questions(f["result"])}
    return names


@app.get("/api/company_search")
def company_search():
    master = excel_run._master(excel_run.IMPORTER_DIR)
    return jsonify({"results": master.search(request.args.get("q", ""), limit=8)})


@app.post("/api/company_answer")
def company_answer():
    """{"name": <customer as written>, "code": <customer code> | null}. null = a new company."""
    body = request.get_json(silent=True) or {}
    return _save_answers([{"name": body.get("name"), "code": body.get("code")}])


@app.post("/api/company_answers")
def company_answers():
    """{"answers": [{"name", "code"}, ...]}: several answers at once, saved together and the files read once."""
    body = request.get_json(silent=True) or {}
    answers = body.get("answers")
    if not isinstance(answers, list) or not answers:
        return jsonify({"error": "answers must be a list of {name, code}"}), 400
    return _save_answers(answers)


def _save_answers(answers):
    """Checks every answer, saves them all, and reads the files once. Nothing is saved if one is wrong."""
    if not _batch_loaded():
        return jsonify({"error": "nothing is loaded; upload a file first"}), 400
    master = excel_run._master(excel_run.IMPORTER_DIR)
    asked = _asked_names()
    for a in answers:
        name, code = a.get("name") if isinstance(a, dict) else None, a.get("code") if isinstance(a, dict) else None
        if not isinstance(name, str) or not (code is None or isinstance(code, str)):
            return jsonify({"error": "name must be text and code must be text or null"}), 400
        if name not in asked:
            return jsonify({"error": "that company was not in the list of questions"}), 400
        if code is not None and not master.has_code(code):
            return jsonify({"error": f"unknown company code {code}"}), 400
    try:
        store = DecisionStore(default_path(excel_run.IMPORTER_DIR))
        for a in answers:
            store.set_new(a["name"]) if a["code"] is None else store.set_same(a["name"], a["code"])
    except (OSError, ValueError) as e:
        return jsonify({"error": f"could not save the answer: {e}"}), 500
    master.apply_decisions({a["name"]: a["code"] for a in answers}, strict=False)
    return _reread()


@app.post("/api/company_pair_answer")
def company_pair_answer():
    """{"a": name, "b": name, "same": true|false}: are these two spellings one company?"""
    body = request.get_json(silent=True) or {}
    a, b, same = body.get("a"), body.get("b"), body.get("same")
    if not isinstance(a, str) or not isinstance(b, str):
        return jsonify({"error": "a and b must be text"}), 400
    run = CURRENT.get("run")
    if run is None or not CURRENT.get("paths"):
        return jsonify({"error": "nothing is loaded; upload a file first"}), 400
    if not isinstance(same, bool) or not any(
            q.get("kind") == "pair" and {q["a"], q["b"]} == {a, b} for q in _company_questions(run)):
        return jsonify({"error": "that pair was not in the list of questions"}), 400
    try:
        DecisionStore(default_path(excel_run.IMPORTER_DIR)).set_pair(a, b, same=same)
    except (OSError, ValueError) as e:
        return jsonify({"error": f"could not save the answer: {e}"}), 500
    return _reread()


def _reread():
    """Read the uploaded files again so the answer just saved takes effect."""
    try:
        if CURRENT["paths"]:
            _read_visits(CURRENT["paths"])
        for key in BATCH["fams"]:
            _read_family(key)
    except Exception as e:
        return jsonify({"error": f"could not re-read the workbook(s): {e}"}), 400
    return jsonify(_batch_response())


@app.post("/api/flag")
def set_flag():
    """{"family", "flag", "include"}: one click to keep or leave out a flagged group (the JET-IT list, finance-list
    contacts, internal colleagues ...). The upload is not read again; nothing is deleted from the files."""
    body = request.get_json(silent=True) or {}
    key, flag, include = body.get("family"), body.get("flag"), body.get("include")
    f, adapter = BATCH["fams"].get(key), families.ADAPTERS.get(key)
    if f is None or f["full"] is None or not isinstance(include, bool) or flag not in {x[0] for x in adapter.flags}:
        return jsonify({"error": "no such group to include or leave out"}), 400
    f["left_out"].discard(flag) if include else f["left_out"].add(flag)
    f["result"] = adapter.apply(f["full"], f["left_out"])
    return jsonify(_batch_response())


@app.post("/api/clear")
def clear():
    with UPLOAD_LOCK:
        _forget()
        _clear_uploads()
    return jsonify({"ok": True})


# ---- dry run and commit: one pass over every family that has something to write, in a safe order ----

def _visits_dry_run(client) -> dict:
    result = CURRENT["result"]
    company = pipeline.preview_companies(client, result)
    p = pipeline.plan_counts(CURRENT["parsed"], result)
    names = {"would_create": "New companies", "created": "New companies", "already_present": "Companies already in the CRM"}
    lines = [(names.get(k, k.replace("_", " ").capitalize()), v) for k, v in company["disposition"].items()]
    lines += [("Contacts", f"{p['contacts']} ({p['contacts_with_phone']} with phone, {p['contacts_with_email']} with email)"),
              ("Visit meetings", p["meetings"]), ("Visit notes", p["notes"]),
              ("Tasks", f"{p['tasks']}" + (f" (due before {p['task_cutoff']} are marked done)" if p["task_cutoff"] else "")),
              ("Weekly reports", p["weeks"])]
    notes = [f"line {i['line']}{' / ' + i['column'] if i.get('column') else ''}: {i['reason']}"
             for i in (company.get("issues") or [])[:6]]
    return {"rows": families.rows(*lines), "notes": notes}


def _dry_run_family(key, client) -> dict:
    if key == "visits":
        return _visits_dry_run(client)
    return {"rows": families.ADAPTERS[key].dry_run(client, BATCH["fams"][key]["result"]), "notes": []}


@app.post("/api/dry_run")
def dry_run():
    """What a commit would write, per family. Writes nothing."""
    plan = [k for k in entry.COMMIT_ORDER if _family_ready(k)]
    if not plan:
        return jsonify({"error": "upload a workbook with something to import first"}), 400
    try:
        client = _client()
    except Exception as e:
        return jsonify({"error": f"could not log in: {e}"}), 400
    out = []
    for key in plan:
        view = {"key": key, "label": entry.FAMILIES[key][0]}
        try:
            out.append({**view, **_dry_run_family(key, client), "error": ""})
        except Exception as e:  # this family could not be checked; say so and go on
            out.append({**view, "rows": [], "notes": [], "error": f"{type(e).__name__}: {e}"})
    return jsonify({"families": out, "wrote": 0})


def _visits_commit(client, emit):
    opts = pipeline.Options(source_system=os.environ.get("IMPORTER_SOURCE_SYSTEM") or "visitlist")
    report = pipeline.run(client, CURRENT["parsed"], CURRENT["result"], opts, on_progress=emit)
    detail = {**asdict(report), "activities_skipped_no_date": len(report.activities_skipped_no_date)}
    disp = (detail.get("companies_report") or {}).get("disposition") or {}
    g = detail.get
    written = (disp.get("created", 0) + g("companies_updated", 0) + g("contacts_created", 0) + g("contacts_updated", 0)
               + g("activities_created", 0) + g("tasks_created", 0) + g("weeks_created", 0))
    present = disp.get("already_present", 0) + g("activities_idempotent", 0) + g("tasks_idempotent", 0) + g("weeks_idempotent", 0)
    no_seat = g("reps_without_seat", [])
    std = families.report(written, present, skipped=g("activities_skipped_no_date", 0), lines=[
        ("Companies created", disp.get("created", 0)), ("Companies enriched", g("companies_updated", 0)),
        ("Contacts created", g("contacts_created", 0)), ("Visits logged", g("activities_created", 0)),
        ("Visits already in the CRM", g("activities_idempotent", 0)), ("Tasks created", g("tasks_created", 0)),
        ("Tasks created as done", g("tasks_closed", 0)), ("Weekly reports created", g("weeks_created", 0)),
        ("Custom fields added", ", ".join(g("fields_created", [])) or None)],
        notes=[f"Reps without a CRM seat, named as author only: {', '.join(no_seat)}"] if no_seat else [])
    return std, detail


def _commit_family(key, client, emit):
    """(the standard report, the visits' own detail or None)."""
    if key == "visits":
        return _visits_commit(client, emit)
    return families.ADAPTERS[key].commit(client, BATCH["fams"][key]["result"], emit), None


@app.post("/api/commit")
def commit():
    plan = [k for k in entry.COMMIT_ORDER if _family_ready(k)]
    if not plan:
        return jsonify({"error": "upload a workbook with something to import first"}), 400
    try:
        client = _client()
    except Exception as e:
        return jsonify({"error": f"could not log in: {e}"}), 400

    if not COMMIT_LOCK.acquire(blocking=False):
        return jsonify({"error": "an import is already running: wait for it to finish before starting another"}), 409
    job_id = str(uuid.uuid4())
    JOBS[job_id] = {"status": "running", "step": "starting", "started_at": time.time(), "family": plan[0], "order": plan,
                    "families": {k: {"label": entry.FAMILIES[k][0], "status": "waiting"} for k in plan}}

    def work():
        job, detail = JOBS[job_id], {}
        try:
            for key in plan:
                state = job["families"][key]
                state["status"] = "running"
                job["family"] = key

                def on_progress(step, _state=state, **info):
                    _state.update(step=step, **info)
                    job.update(step=step, **info)

                try:
                    report, detail[key] = _commit_family(key, client, on_progress)
                except Exception as e:  # this family failed; the others still run
                    report = families.report(0, 0, failed=1, errors=[f"{type(e).__name__}: {e}"])
                state.update(report, label=state["label"])
            job.update(status="done", result={"families": [{"key": k, **job["families"][k]} for k in plan],
                                              "visits": detail.get("visits")})
        except Exception as e:
            job.update(status="error", error=str(e))
        finally:
            COMMIT_LOCK.release()

    threading.Thread(target=work, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.get("/api/commit/<job_id>")
def commit_status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "no such job"}), 404
    return jsonify(job)


# ---------------------------------------------------------------------------------------------
# Danger zone: rebuild the margince dev database and reseed the demo baseline.
#
# `make dev-fresh` is margince's own recipe for this (drop, create, db-init, migrate, boot), so the
# database name and grants always match what its dev stack expects. Every step writes to a log file,
# never to a pipe: `make dev` leaves the api/worker/vite running in the background, and they keep any
# inherited stdout/stderr pipe open, so capturing output would block until the timeout even though
# the stack came up fine. That was why the button always failed at "starting the dev stack".
# ---------------------------------------------------------------------------------------------

MARGINCE_DIR = os.path.abspath(os.path.join(ROOT, "..", "margince"))
RESET_LOG = os.path.join(UPLOAD_DIR, "reset_db.log")
# seed-dev only seeds USD/GBP/CHF; won CNY deals from the mail archive need this (local, representative value)
CNY_RATE_SQL = ("INSERT INTO fx_rate (from_currency, to_currency, rate, rate_date) VALUES ('CNY', 'EUR', 0.12, "
                "CURRENT_DATE) ON CONFLICT (from_currency, to_currency, rate_date) DO UPDATE SET rate = EXCLUDED.rate")
RESET_STEPS = [
    ("stopping the dev stack", ["make", "dev-stop"], 120),
    ("rebuilding the database and starting the dev stack", ["make", "dev-fresh"], 600),
    ("seeding demo data", ["make", "seed-dev"], 300),
    ("adding the CNY exchange rate", ["bash", "scripts/dev-psql.sh", "15432", "margince", "-v", "ON_ERROR_STOP=1",
                                      "-c", CNY_RATE_SQL], 60),
]
RESET_LOCK = threading.Lock()


def _run_logged(cmd: list[str], timeout: int, log) -> int | None:
    """Runs cmd with output appended to log; returns its exit code, or None on timeout."""
    log.write(f"\n$ {' '.join(cmd)}\n")
    log.flush()
    proc = subprocess.Popen(cmd, cwd=MARGINCE_DIR, stdin=subprocess.DEVNULL, stdout=log, stderr=subprocess.STDOUT,
                            start_new_session=True)  # the dev stack outlives this request (and this server)
    try:
        return proc.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        return None


def _log_tail(n: int = 1) -> list[str]:
    try:
        with open(RESET_LOG, errors="replace") as fh:
            lines = [l.rstrip() for l in fh if l.strip() and not l.startswith("$ ")]
    except OSError:
        return []
    return lines[-n:]


@app.post("/api/reset_db")
def reset_db():
    if not os.path.isdir(MARGINCE_DIR):
        return jsonify({"error": f"margince checkout not found at {MARGINCE_DIR}"}), 400
    if not RESET_LOCK.acquire(blocking=False):
        return jsonify({"error": "a database reset is already running"}), 409
    job_id = str(uuid.uuid4())
    JOBS[job_id] = {"status": "running", "step": RESET_STEPS[0][0], "started_at": time.time()}

    def work():
        try:
            os.makedirs(UPLOAD_DIR, exist_ok=True)
            with open(RESET_LOG, "w") as log:
                for label, cmd, timeout in RESET_STEPS:
                    JOBS[job_id].update(step=label)
                    code = _run_logged(cmd, timeout, log)
                    if code != 0:
                        why = f"timed out after {timeout}s" if code is None else (_log_tail() or ["failed"])[0]
                        JOBS[job_id].update(status="error", error=f"{label}: {why} (full log: {RESET_LOG})")
                        return
            _forget()
            MAIL.update(loaded=False, summary=None)
            JOBS[job_id].update(status="done", result={"message": "database reset to the seeded demo baseline"})
        except Exception as e:
            JOBS[job_id].update(status="error", error=str(e))
        finally:
            RESET_LOCK.release()

    threading.Thread(target=work, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.get("/api/job/<job_id>")
def job_status(job_id):
    return commit_status(job_id)


# ---------------------------------------------------------------------------------------------
# Email archive (extracted.json from mail/extract.py). Runs `python -m mail.cli` as a subprocess:
# the CLI is the tested path (parse -> preview -> commit, idempotent on Message-ID), and a long commit
# streams its progress on stdout, which is how the progress bar below follows it.
# ---------------------------------------------------------------------------------------------

MAIL_UPLOADED = os.path.join(UPLOAD_DIR, "mail")  # a copy of uploaded files
MAIL_CANONICAL = os.path.join(ROOT, "mail", "staging")  # "use the current extraction": with every reading and review
MAIL_FILES = os.path.join(ROOT, "mail", "staging", "files")  # attachment files the extractor stored
MAIL = {"loaded": False, "summary": None, "staging": MAIL_UPLOADED}
PROGRESS = re.compile(r"^\s*(emails)\s+(\d+)/(\d+)")
STAGES = ("vocabulary", "companies", "contacts", "products", "deals", "emails", "routing", "titles", "tasks and notes",
          "attachments")


def _mail_env() -> dict:
    cfg = _load_config()
    return {**os.environ, "MARGINCE_BASE_URL": cfg["base_url"], "MARGINCE_EMAIL": cfg["email"],
            "MARGINCE_PASSWORD": cfg["password"], "PYTHONUNBUFFERED": "1"}


def _cli(mode: str, *extra: str) -> subprocess.Popen:
    cmd = [sys.executable, "-u", "-m", "mail.cli", "--mode", mode, "--staging", MAIL["staging"], *extra]
    return subprocess.Popen(cmd, cwd=ROOT, env=_mail_env(), stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            text=True, bufsize=1)


def _json_objects(text: str) -> list:
    """cli.py prints one or two JSON objects back to back."""
    out, i, dec = [], 0, json.JSONDecoder()
    while True:
        i = text.find("{", i)
        if i < 0:
            return out
        try:
            obj, end = dec.raw_decode(text, i)
        except ValueError:
            i += 1
            continue
        out.append(obj)
        i = end


def _mail_summary(counts: dict, warnings: list, extracted: dict) -> dict:
    present = missing = 0
    for m in extracted["messages"]:
        for a in m["attachments"]:
            if a.get("stored_as") and not a.get("decoration"):
                if os.path.exists(os.path.join(MAIL_FILES, a["stored_as"])):
                    present += 1
                else:
                    missing += 1
    return {
        "loaded": True, "counts": counts, "warnings": warnings,
        "by_folder_type": extracted.get("summary", {}).get("by_folder_type", {}),
        "documents_unreadable": extracted.get("summary", {}).get("documents_unreadable"),
        "attachments_on_disk": present, "attachments_missing": missing,
        "company_rows": [{"name": c["name"], "type": c.get("customer_type"), "line": c.get("business_line"),
                          "domains": ", ".join(c.get("domains", [])), "messages": c.get("messages", 0),
                          "origin": c.get("origin")} for c in extracted["companies"]],
        "contact_rows": [{"full_name": c["full_name"], "company": c.get("company") or "", "title": c.get("title") or "",
                          "phones": ", ".join(p["phone"] for p in c["phones"]), "emails": ", ".join(c["emails"]),
                          "messages": c["messages"], "internal": c["internal"]} for c in extracted["contacts"]],
    }


@app.get("/api/mail/dataset")
def mail_dataset():
    return jsonify(MAIL["summary"] if MAIL["loaded"] else {"loaded": False})


@app.post("/api/mail/upload")
def mail_upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "no files in the request"}), 400
    extracted = enrichment = None
    for f in files:
        try:
            data = json.load(f.stream)
        except ValueError as e:
            return jsonify({"error": f"{f.filename} is not valid JSON: {e}"}), 400
        if isinstance(data, dict) and "messages" in data and "contacts" in data:
            extracted = data
        elif isinstance(data, dict):
            enrichment = data
        else:
            return jsonify({"error": f"{f.filename} is not an extracted.json or enrichment.json object"}), 400
    if extracted is None:
        return jsonify({"error": "none of the files is an extracted.json (it needs 'messages' and 'contacts')"}), 400

    os.makedirs(MAIL_UPLOADED, exist_ok=True)
    MAIL["staging"] = MAIL_UPLOADED
    with open(os.path.join(MAIL_UPLOADED, "extracted.json"), "w") as fh:
        json.dump(extracted, fh, ensure_ascii=False)
    ep = os.path.join(MAIL_UPLOADED, "enrichment.json")
    if enrichment is not None:
        with open(ep, "w") as fh:
            json.dump(enrichment, fh, ensure_ascii=False)
    elif os.path.exists(ep):
        os.remove(ep)
    link = os.path.join(MAIL_UPLOADED, "files")
    if os.path.islink(link):
        os.remove(link)
    if not os.path.exists(link):
        os.symlink(MAIL_FILES, link)

    proc = _cli("parse")
    out, err = proc.communicate()
    if proc.returncode != 0:
        return jsonify({"error": "could not build the import plan: " + (err.strip().splitlines() or ["?"])[-1]}), 400
    counts = (_json_objects(out) or [{}])[0]
    warnings = [l[len("WARNING: "):] for l in err.splitlines() if l.startswith("WARNING: ")]
    MAIL["summary"] = {**_mail_summary(counts, warnings, extracted), "has_enrichment": enrichment is not None,
                       "files": [f.filename for f in files]}
    MAIL["loaded"] = True
    return jsonify(MAIL["summary"])


@app.post("/api/mail/use_current")
def mail_use_current():
    """The extraction in mail/staging as it is, with every merged reading and review decision (no upload copy)."""
    path = os.path.join(MAIL_CANONICAL, "extracted.json")
    if not os.path.exists(path):
        return jsonify({"error": "mail/staging has no extracted.json yet: run python -m mail.extract"}), 400
    MAIL["staging"] = MAIL_CANONICAL
    proc = _cli("parse")
    out, err = proc.communicate()
    if proc.returncode != 0:
        return jsonify({"error": "could not build the import plan: " + (err.strip().splitlines() or ["?"])[-1]}), 400
    counts = (_json_objects(out) or [{}])[0]
    warnings = [l[len("WARNING: "):] for l in err.splitlines() if l.startswith("WARNING: ")]
    with open(path) as fh:
        extracted = json.load(fh)
    reading = counts.pop("reading", None)
    MAIL["summary"] = {**_mail_summary(counts, warnings, extracted), "has_enrichment": bool(reading),
                       "reading": reading, "files": ["mail/staging (current extraction)"]}
    MAIL["loaded"] = True
    return jsonify(MAIL["summary"])


@app.post("/api/mail/clear")
def mail_clear():
    MAIL.update(loaded=False, summary=None)
    return jsonify({"ok": True})


@app.post("/api/mail/preview")
def mail_preview():
    if not MAIL["loaded"]:
        return jsonify({"error": "upload an extracted.json first"}), 400
    proc = _cli("preview")
    out, err = proc.communicate()
    if proc.returncode != 0:
        return jsonify({"error": (err.strip().splitlines() or ["preview failed"])[-1]}), 400
    objs = _json_objects(out)
    return jsonify({"plan": objs[0] if objs else {}, "crm": objs[1] if len(objs) > 1 else {}})


@app.post("/api/mail/commit")
def mail_commit():
    if not MAIL["loaded"]:
        return jsonify({"error": "upload an extracted.json first"}), 400
    body = request.get_json(silent=True) or {}
    extra = ["--attachments", "documents" if body.get("attachments") else "none"]
    if str(body.get("limit_emails") or "").isdigit() and int(body["limit_emails"]) > 0:
        extra += ["--limit-emails", str(int(body["limit_emails"]))]
    job_id = str(uuid.uuid4())
    JOBS[job_id] = {"status": "running", "step": "starting", "started_at": time.time()}

    def work():
        try:
            proc = _cli("commit", *extra)
            err_lines: list[str] = []
            threading.Thread(target=lambda: err_lines.extend(proc.stderr), daemon=True).start()
            lines = []
            for line in proc.stdout:
                lines.append(line)
                stage, m = line.strip(), PROGRESS.match(line)
                if m:
                    JOBS[job_id].update(step="emails", done=int(m.group(2)), total=int(m.group(3)))
                elif stage in STAGES:
                    JOBS[job_id].update(step=stage, done=None, total=None)
            proc.wait()
            if proc.returncode != 0:
                tail = [l for l in err_lines if l.strip()][-1:] or ["the import stopped unexpectedly"]
                JOBS[job_id].update(status="error", error=tail[0].strip())
                return
            with open(os.path.join(MAIL["staging"], "report.json")) as fh:
                report = json.load(fh)
            JOBS[job_id].update(status="done", result={
                "created": report["created"], "existing": report["existing"], "approvals": report.get("approvals", {}),
                "failed_count": len(report["failed"]), "failed": report["failed"][:25]})
        except Exception as e:
            JOBS[job_id].update(status="error", error=str(e))

    threading.Thread(target=work, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.post("/api/mail/verify")
def mail_verify():
    """Fact check of the uploaded plan: against the original archive, then against the CRM. Reads only."""
    if not MAIL["loaded"]:
        return jsonify({"error": "upload an extracted.json first"}), 400
    job_id = str(uuid.uuid4())
    JOBS[job_id] = {"status": "running", "step": "reading the archive", "started_at": time.time()}

    def work():
        try:
            cache = os.path.join(ROOT, "mail", "staging", ".verify_cache.pickle")
            proc = subprocess.Popen([sys.executable, "-u", "-m", "mail.verify", "--staging", MAIL["staging"],
                                     "--cache", cache], cwd=ROOT, env=_mail_env(), stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE, text=True, bufsize=1)
            err_lines: list[str] = []
            threading.Thread(target=lambda: err_lines.extend(proc.stderr), daemon=True).start()
            for line in proc.stdout:
                m = PROGRESS.match(line)
                if m:
                    JOBS[job_id].update(step=f"comparing emails with the CRM {m.group(2)}/{m.group(3)}")
                elif line.startswith("  reading"):
                    JOBS[job_id].update(step=line.strip())
            proc.wait()
            path = os.path.join(MAIL["staging"], "verify_report.json")
            if proc.returncode not in (0, 1) or not os.path.exists(path):  # 1 = problems found, still a report
                tail = [l for l in err_lines if l.strip()][-1:] or ["the fact check stopped unexpectedly"]
                JOBS[job_id].update(status="error", error=tail[0].strip())
                return
            with open(path) as fh:
                report = json.load(fh)
            order = ["NOT FOUND", "MISSING", "CHANGED", "EXTRA", "DUPLICATE"]
            problems = sorted((f for f in report["findings"] if f["status"] in order),
                              key=lambda f: (f["check"], order.index(f["status"])))
            JOBS[job_id].update(status="done", result={
                "fields": report["fields"], "problems": problems[:40], "problem_count": len(problems),
                "csv": os.path.join(MAIL["staging"], "verify_report.csv"), "has_enrichment": MAIL["summary"]["has_enrichment"]})
        except Exception as e:
            JOBS[job_id].update(status="error", error=str(e))

    threading.Thread(target=work, daemon=True).start()
    return jsonify({"job_id": job_id})


# ---------------------------------------------------------------------------------------------
# Reading review: what the two readings found that needs a person (mail/reading.py, READING.md)
# ---------------------------------------------------------------------------------------------

READING_DIR = os.path.join(MAIL_CANONICAL, "reading")
ENRICHMENT_DIR = os.path.join(MAIL_CANONICAL, "enrichment")
SAFE_SLUG = re.compile(r"^[a-z0-9_-][a-z0-9_.-]*$")


def _reading():
    from mail import reading  # imported late: it loads the archive checker
    return reading


def _case(slug: str) -> tuple[dict, dict | None]:
    if not SAFE_SLUG.match(slug):
        raise FileNotFoundError(slug)
    with open(os.path.join(READING_DIR, slug, "case.json")) as fh:
        case = json.load(fh)
    path = os.path.join(ENRICHMENT_DIR, f"{slug}.json")
    merged = None
    if os.path.exists(path):
        with open(path) as fh:
            merged = json.load(fh)
    return case, merged


def _counts(merged: dict | None) -> dict:
    c = {"auto": 0, "needs_review": 0, "reviewed": 0, "rejected": 0}
    for it in (merged or {}).get("items", []):
        c["reviewed" if it.get("review") else it["state"]] += 1
    return c


def _snippet(case: dict, ev: dict) -> dict:
    """Where a quote stands: the mail's header and the text around it, or the attachment it is printed on."""
    if ev.get("message_id"):
        alias, m = next(((a, m) for a, m in case["messages"].items() if m["message_id"] == ev["message_id"]), (None, None))
        if not m:
            return {"source": "a mail outside this case"}
        text = f"Subject: {m['subject'] or ''}\n\n" + (m["text"] or "") + "\n\n" + (m["history"] or "")
        words = [re.escape(w) for w in (ev.get("quote") or "").split()]
        hit = re.search(r"\s*".join(words), text) if words else None
        out = {"source": f"{alias} · {(m['date'] or '')[:10]} · {m['from']} · {m['subject']}"}
        if hit:
            out.update(before=text[max(0, hit.start() - 220):hit.start()], quote=text[hit.start():hit.end()],
                       after=text[hit.end():hit.end() + 220])
        else:
            out.update(before="", quote=ev.get("quote"), after="", note="this quote is in the original mail, outside the part shown in the case file")
        return out
    doc_alias, doc = next(((a, d) for a, d in case["documents"].items() if d["sha256"] == ev.get("sha256")), (None, None))
    return {"source": f"{doc_alias} · {doc['filename']}" + (" · scan" if doc and doc["kind"] == "scan" else "") if doc else "an attachment",
            "quote": ev.get("quote"), "doc": doc_alias}


@app.get("/api/review/cases")
def review_cases():
    rows = []
    for slug in sorted(os.listdir(READING_DIR)) if os.path.isdir(READING_DIR) else []:
        try:
            case, merged = _case(slug)
        except (FileNotFoundError, ValueError):
            continue
        rows.append({"slug": slug, "case": case["case"], "mails": len(case["messages"]),
                     "readings": "".join(p.upper() for p in "ab" if os.path.exists(os.path.join(READING_DIR, slug, f"pass_{p}.json"))),
                     "merged": merged is not None, "stale": bool(merged) and merged["case_digest"] != case["digest"],
                     **_counts(merged)})
    return jsonify({"cases": rows})


@app.get("/api/review/case/<slug>")
def review_case(slug):
    try:
        case, merged = _case(slug)
    except (FileNotFoundError, ValueError):
        return jsonify({"error": "no such case"}), 404
    if not merged:
        return jsonify({"error": "this case has not been read and merged yet"}), 400
    items = []
    for it in merged["items"]:
        ev = []
        for side in ("a", "b"):
            for e in (it.get(side) or {}).get("evidence", []):
                if not any(x["quote"] == e["quote"] for x in ev):
                    ev.append({**e, **_snippet(case, e), "side": side})
        items.append({**it, "evidence_shown": ev})
    order = {"needs_review": 0, "auto": 1, "rejected": 2}
    items.sort(key=lambda it: (bool(it.get("review")), order[it["state"]], it["id"]))
    return jsonify({"case": case["case"], "slug": slug, "company": case["company"], "counts": _counts(merged),
                    "stale": merged["case_digest"] != case["digest"], "items": items,
                    "companies": case["known"]["companies"]})


@app.post("/api/review/case/<slug>/decide")
def review_decide(slug):
    if not SAFE_SLUG.match(slug):
        return jsonify({"error": "no such case"}), 404
    body = request.get_json(force=True)
    if not (body.get("by") or "").strip():
        return jsonify({"error": "enter your name first, so the CRM can say who checked it"}), 400
    try:
        item = _reading().save_review(os.path.join(ENRICHMENT_DIR, f"{slug}.json"), body["id"], body["choice"],
                                      body.get("value"), body["by"].strip())
    except (FileNotFoundError, StopIteration, ValueError, KeyError) as e:
        return jsonify({"error": f"could not save the decision: {e}"}), 400
    return jsonify({"ok": True, "review": item["review"]})


@app.get("/api/review/case/<slug>/file/<alias>")
def review_file(slug, alias):
    """A scan, so its numbers can be checked against the original."""
    try:
        case, _ = _case(slug)
        doc = case["documents"][alias]
    except (FileNotFoundError, KeyError, ValueError):
        return jsonify({"error": "no such attachment"}), 404
    files = os.path.realpath(os.path.join(MAIL_CANONICAL, "files"))
    path = os.path.realpath(os.path.join(ROOT, doc["file"]))
    if not path.startswith(files + os.sep) or not os.path.exists(path):
        return jsonify({"error": "file not found"}), 404
    return send_from_directory(os.path.dirname(path), os.path.basename(path))


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=int(os.environ.get("PORT", 5055)))
    app.run(host="127.0.0.1", port=ap.parse_args().port, debug=False, threaded=True)
