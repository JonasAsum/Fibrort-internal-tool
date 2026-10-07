"""Local web UI for both importers: visit-list workbooks (visits/) and the email
archive (mail/). Upload, review what was matched, preview, then commit - without
typing the CLI modes by hand. Single-user, local-only; state lives in process
memory (CURRENT, MAIL) and is lost on restart, same as a CLI run.

    cd importer && .venv/bin/python -m webui.server      # http://localhost:5055
"""
import glob
import json
import os
import re
import subprocess
import sys
import threading
import time
import uuid
from dataclasses import asdict

from flask import Flask, jsonify, request, send_from_directory

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)  # importer/
if ROOT not in sys.path:  # also allow `python webui/server.py`
    sys.path.insert(0, ROOT)

from common.client import ApiError, MargenceClient  # noqa: E402
from visits import pipeline  # noqa: E402
from visits.match import match  # noqa: E402
from visits.reader import parse_paths  # noqa: E402

app = Flask(__name__, static_folder="static")

CONFIG_PATH = os.path.join(HERE, ".webui_config.json")
UPLOAD_DIR = os.path.join(HERE, ".uploads")

DEFAULT_CONFIG = {"base_url": "http://localhost:8080", "email": "admin@demo.test", "password": "demo-password-123"}

CURRENT = {"files": [], "parsed": None, "result": None}
JOBS: dict[str, dict] = {}


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
    }


@app.get("/api/dataset")
def get_dataset():
    return jsonify(_dataset_summary())


@app.post("/api/upload")
def upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "no files in the request"}), 400

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    for stale in glob.glob(os.path.join(UPLOAD_DIR, "*.xlsx")):
        os.remove(stale)

    saved_paths, names = [], []
    for f in files:
        if not f.filename.lower().endswith(".xlsx"):
            return jsonify({"error": f"{f.filename} is not an .xlsx file"}), 400
        path = os.path.join(UPLOAD_DIR, f.filename)
        f.save(path)
        saved_paths.append(path)
        names.append(f.filename)

    try:
        parsed = parse_paths(saved_paths)
    except Exception as e:
        return jsonify({"error": f"could not parse the workbook(s): {e}"}), 400

    CURRENT["files"] = names
    CURRENT["parsed"] = parsed
    CURRENT["result"] = match(parsed.visits)
    return jsonify(_dataset_summary())


@app.post("/api/clear")
def clear():
    CURRENT.update(files=[], parsed=None, result=None)
    for stale in glob.glob(os.path.join(UPLOAD_DIR, "*.xlsx")):
        os.remove(stale)
    return jsonify({"ok": True})


@app.post("/api/preview")
def preview():
    if CURRENT["result"] is None:
        return jsonify({"error": "upload a workbook first"}), 400
    try:
        client = _client()
    except Exception as e:
        return jsonify({"error": f"could not log in: {e}"}), 400

    result = CURRENT["result"]
    try:
        company_report = pipeline.preview_companies(client, result)
    except ApiError as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"company": company_report, "plan": pipeline.plan_counts(CURRENT["parsed"], result)})


@app.post("/api/commit")
def commit():
    if CURRENT["result"] is None:
        return jsonify({"error": "upload a workbook first"}), 400
    try:
        client = _client()
    except Exception as e:
        return jsonify({"error": f"could not log in: {e}"}), 400

    job_id = str(uuid.uuid4())
    JOBS[job_id] = {"status": "running", "step": "starting", "started_at": time.time()}
    parsed, result = CURRENT["parsed"], CURRENT["result"]

    def on_progress(step, **info):
        JOBS[job_id].update(step=step, **info)

    def work():
        try:
            report = pipeline.run(client, parsed, result, on_progress=on_progress)
            JOBS[job_id].update(status="done", result={
                **asdict(report),
                "activities_skipped_no_date": len(report.activities_skipped_no_date),
            })
        except Exception as e:
            JOBS[job_id].update(status="error", error=str(e))

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
            CURRENT.update(files=[], parsed=None, result=None)
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
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", 5055)), debug=False, threaded=True)
