"""The Checks page: plain-language PASS / WARN / FAIL rows, each failing row saying how to fix it.
Reads only: it never creates a custom field, a record or a file outside a temporary one.

`environment(ctx)` answers "can this site work right now" (CRM, login, modules, folders, company
master, custom fields); `selftest.run()` answers "do the readers still read" and needs no CRM.
"""
import importlib
import os
import tempfile

from excel import accounts_margince, contacts_margince, receivables_margince

PASS, WARN, FAIL = "pass", "warn", "fail"
MODULES = (("flask", "the web server"), ("openpyxl", "reading .xlsx files"), ("xlrd", "reading old .xls files"),
           ("requests", "talking to the CRM"), ("rapidfuzz", "matching company names"),
           ("pypinyin", "matching Chinese company names"))
REINSTALL = "In a terminal: cd importer && .venv/bin/pip install -r requirements.txt, then restart the site."


def row(id_, title, status, detail, fix="") -> dict:
    return {"id": id_, "title": title, "status": status, "detail": detail, "fix": fix if status != PASS else ""}


def crm_reachable(ctx) -> dict:
    import requests
    url = ctx["load_config"]()["base_url"]
    try:
        r = requests.get(url, timeout=4)
        return row("crm", "CRM reachable", PASS, f"{url} answered (HTTP {r.status_code}).")
    except requests.RequestException as exc:
        return row("crm", "CRM reachable", FAIL, f"{url} did not answer ({type(exc).__name__}).",
                   "Start the CRM (cd ../margince && make dev), or open Settings (the gear) and correct the Server URL.")


def login(ctx):
    """(row, client or None)."""
    try:
        client = ctx["make_client"]()
        me = client.get("/v1/me")
        who = (me.get("user") or {}).get("display_name") or ctx["load_config"]()["email"]
        return row("login", "Login with the saved settings", PASS,
                   f"Signed in as {who} in {me.get('workspace_name') or 'the workspace'}."), client
    except Exception as exc:
        return row("login", "Login with the saved settings", FAIL, f"Could not sign in: {str(exc)[:160]}",
                   "Open Settings (the gear), check the email and password, then press Save & test."), None


def modules() -> dict:
    missing = []
    for name, why in MODULES:
        try:
            importlib.import_module(name)
        except ImportError:
            missing.append(f"{name} ({why})")
    if missing:
        return row("modules", "Required Python modules", FAIL, "Not installed: " + ", ".join(missing) + ".", REINSTALL)
    return row("modules", "Required Python modules", PASS, f"All {len(MODULES)} are installed.")


def _writable(folder: str) -> str:
    """'' when a file can be created in the folder, else why not."""
    try:
        os.makedirs(folder, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=folder, prefix=".check-"):
            pass
        return ""
    except OSError as exc:
        return f"{folder}: {exc.strerror or exc}"


def folders(ctx) -> dict:
    wanted = {"uploads": ctx["upload_dir"], "settings": os.path.dirname(ctx["config_path"]) or ".",
              "company answers": str(ctx["decisions_path"].parent)}
    bad = [f"{name} ({why})" for name, folder in wanted.items() if (why := _writable(folder))]
    if bad:
        return row("folders", "Upload and state folders writable", FAIL, "Cannot write to: " + "; ".join(bad) + ".",
                   "Fix the folder's permissions, or point IMPORTER_UPLOAD_DIR / IMPORTER_CONFIG / COMPANY_DECISIONS_FILE "
                   "at a folder you can write to, then restart the site.")
    return row("folders", "Upload and state folders writable", PASS, "Uploads, settings and company answers can be saved.")


def company_master(ctx) -> dict:
    try:
        master = ctx["master"]()
    except Exception as exc:
        return row("master", "Company master loaded", FAIL, f"Could not load it: {str(exc)[:160]}",
                   "Check excel/company_decisions.json is valid JSON (or remove it), and that the customer list files exist.")
    size = getattr(master, "size", 0)
    if not size:
        return row("master", "Company master loaded", WARN,
                   "No customer list was found, so every customer name will be asked about as a new company.",
                   "Put Customer_List_20260104.xlsx under mail/staging/files/ and the ERP customer lists in "
                   "'Excel Files from Sales Meetings' next to the importer folder, then restart the site.")
    if getattr(master, "origin", "files") == "crm":
        return row("master", "Company master loaded", PASS, f"{size} {'company' if size == 1 else 'companies'} known, read from the CRM "
                   f"({master.sources.get('lists', 0)} enriched from the customer lists; {master.sources.get('lists_only', 0)} more only in the "
                   "lists, not imported yet).")
    from excel import run as excel_run
    note = excel_run.master_note()
    return row("master", "Company master loaded", WARN if note else PASS,
               f"{size} {'company' if size == 1 else 'companies'} known ({master.sources.get('master', 0)} from the "
               f"customer list, {master.sources.get('erp', 0)} from the ERP lists)." + (f" Not the CRM: {note}." if note else ""),
               "Check the CRM address and login (gear icon), then upload again." if note else None)


def decisions(ctx) -> dict:
    from excel.decisions import DecisionStore
    try:
        n = len(DecisionStore(ctx["decisions_path"]).as_dict())
        return row("answers", "Saved company answers readable", PASS, f"{n} saved answers.")
    except ValueError as exc:
        return row("answers", "Saved company answers readable", FAIL, str(exc)[:200],
                   "Open the file named above and fix the JSON, or delete it (the questions are asked again).")


def needed_fields() -> list[tuple[str, str, str]]:
    """(CRM object, label, what needs it) for every custom field an import may use."""
    out = [("company", label, "receivables (AR number on the account)") for _, label, _ in receivables_margince.FIELDS]
    out += [("company", label, "opportunity data (competitor, forecast, application)") for attr, label, _ in
            accounts_margince.COMPANY_FIELDS if attr in ("application", "competitor", "forecast")]
    out += [("company", label, "contact lists") for _, label, *_ in contacts_margince.COMPANY_FIELDS]
    out += [("contact", label, "contact lists") for _, label, *_ in contacts_margince.CONTACT_FIELDS]
    out += [("lead", label, "leads") for attr, label, _ in accounts_margince.LEAD_FIELDS if attr in ("application", "details")]
    return out


def custom_fields(client) -> dict:
    if client is None:
        return row("fields", "CRM custom fields for AR, orders and opportunity data", WARN,
                   "Not checked: no login.", "Fix the login first.")
    have: dict[str, set] = {}
    missing = []
    try:
        for obj, label, why in needed_fields():
            if obj not in have:
                have[obj] = {f["label"].lower() for f in client.custom_fields(obj)}
            ok = any(c.lower() in have[obj] for c in (label, f"{label} ({accounts_margince.FIELD_SOURCE})",
                                                      f"{label} ({contacts_margince.SOURCE})",
                                                      f"{label}{receivables_margince.SUFFIX}"))
            if not ok:
                missing.append(f"{label} ({obj}, for {why})")
    except Exception as exc:
        return row("fields", "CRM custom fields for AR, orders and opportunity data", FAIL,
                   f"Could not read the CRM's custom fields: {str(exc)[:160]}",
                   "Check that the saved login belongs to a person who may read custom fields (not an agent passport).")
    if missing:
        return row("fields", "CRM custom fields for AR, orders and opportunity data", WARN,
                   f"{len(missing)} missing: " + "; ".join(missing[:6]) + (" ..." if len(missing) > 6 else "") +
                   ". They are created automatically at the first import that needs them.",
                   "Nothing to do, unless the saved login may not change the CRM's schema: then ask an admin to "
                   "create them, or import with an admin login once.")
    return row("fields", "CRM custom fields for AR, orders and opportunity data", PASS,
               f"All {len(needed_fields())} fields exist.")


def environment(ctx) -> list[dict]:
    """Every environment check, in the order a person should fix them."""
    reach = crm_reachable(ctx)
    out = [reach]
    client = None
    if reach["status"] == PASS:
        login_row, client = login(ctx)
    else:
        login_row = row("login", "Login with the saved settings", WARN, "Not tried: the CRM did not answer.",
                        "Fix the CRM connection first.")
    out.append(login_row)
    out += [modules(), folders(ctx), company_master(ctx), decisions(ctx), custom_fields(client)]
    if ctx.get("stale") is not None:
        out.append(row("code", "Site runs the current code", WARN if ctx["stale"] else PASS,
                       "The files on disk changed after the site started." if ctx["stale"] else "Up to date.",
                       "Restart the site: cd importer && .venv/bin/python -m webui.server"))
    return out


def selftest_rows(results: list[dict]) -> list[dict]:
    """The self-test results as check rows: one per family."""
    from .entry import FAMILIES
    names = {"visits": "Visit reports (01-03)", **{k: v[0] for k, v in FAMILIES.items()}}
    out = []
    for r in results:
        title = f"Reads a {names.get(r['family'], r['family'])} workbook"
        if r["ok"]:
            out.append(row(f"self-{r['family']}", title, PASS,
                           "Built a tiny workbook, the front door recognised it and the reader found "
                           + ", ".join(f"{v} {'visits' if k == 'found' else k}" for k, v in r["expected"].items()) + "."))
        elif r["error"]:
            out.append(row(f"self-{r['family']}", title, FAIL, f"The reader stopped: {r['error'][:160]}",
                           "This is a bug in the importer, not in your files: run the tests (pytest) and report the failure."))
        else:
            what = ("sent it to the wrong reader (" + str(r["detected_as"]) + ")" if r["detected_as"] != r["family"]
                    else f"expected {r['expected']} but found {r['got']}")
            out.append(row(f"self-{r['family']}", title, FAIL, f"The front door or reader {what}.",
                           "This is a bug in the importer, not in your files: run the tests (pytest) and report the failure."))
    return out
