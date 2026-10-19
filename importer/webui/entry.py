"""The front door: tells what an uploaded workbook is, by its content (the header rows of its
sheets), never only by its name. Reads the file, writes nothing.

A file is one of:
  import      a list this importer can read: family "visits" | "accounts" | "contacts" | "todos" |
              "orders" | "receivables" | "quotes" (folders 01-03, 04-05, 06, 07, 08, 09, 10 of the CRM folder review)
  kept_out    recognised, and the folder review says it does not belong in the CRM (invoice and
              packing lists, the formula workbook, bank cash flows, AP, fixed costs, Laepple 2023 ...)
  unknown     no list this importer knows; `reason` says what was looked for
  unreadable  the workbook could not be opened
"""
import os
import re
from collections import Counter, defaultdict

from excel import accounts, contacts, layouts, orders, quotes, receivables, todos
from excel.files import file_display_name
from excel.sheet import load

FOLDER_NAME = re.compile(r"^(0[1-9]) ")

# family -> (label, CRM target as the folder review says it, position in the commit order)
FAMILIES = {
    "accounts": ("Customer lists (04, 05)", "Accounts, opportunity data on the account, leads", 1),
    "contacts": ("Contact lists (06)", "Contacts on their accounts, tagged by where they came from", 2),
    "visits": ("Visit reports (01-03)", "Meetings, contacts, tasks and weekly notes", 3),
    "todos": ("To-dos and sales topics (07)", "Tasks (owner, due date, status)", 4),
    "orders": ("Sales orders (08)", "Read-only order history on each account", 5),
    "receivables": ("Receivables (09)", "Open and overdue receivables as a number on the account", 6),
    "quotes": ("Quote log (10)", "Deals with one offer per quote revision, on the account", 7),
}
COMMIT_ORDER = sorted(FAMILIES, key=lambda k: FAMILIES[k][2])

VISIT_LAYOUTS = {"mock": ("01", "Weekly visit report"), "weekly2026": ("01", "Weekly sales report"),
                 "lucas": ("02", "Weekly planning (Lucas)"), "detail": ("03", "Customer visits in detail")}
ACCOUNT_LAYOUTS = {"erp": ("04", "ERP customer list"), "master": ("05", "Customer list"),
                   "supporting": ("05", "Customer list (handover)"), "registration": ("05", "Customer registration"),
                   "effective": ("05", "Effective-customer analysis")}
CONTACT_LAYOUTS = {"list": ("06", "Contact list (finance, receivable and payable)"),
                   "sections": ("06", "Contact list (handover)")}
QUOTE_LAYOUTS = {"quotes": ("10", "Quote log")}
TODO_LAYOUTS = {"todo": ("07", "To-do list"), "topics": ("07", "Sales topics list")}
LAYOUT_TABLES = {"accounts": ACCOUNT_LAYOUTS, "contacts": CONTACT_LAYOUTS, "todos": TODO_LAYOUTS, "visits": VISIT_LAYOUTS,
                 "quotes": QUOTE_LAYOUTS}

# What the folder review says about the files that stay out. Keys are matched on the importers' own reason text.
KEPT_OUT = {
    "invoice_packing": ("Invoice / packing list", "08",
                        "invoice and packing lists do not belong in the CRM; orders live in the ERP"),
    "formula_workbook": ("Formula workbook", "08",
                         "the Actual_mit_Formeln calculation workbook is not CRM data; orders live in the ERP"),
    "stock_list": ("Inventory / stock list", "08", "stock lists are not CRM data; they live in the ERP"),
    "bank": ("Bank cash flows", "09", "bank cash flows stay in accounting"),
    "payable": ("Accounts payable (AP)", "09", "accounts payable stays in accounting"),
    "fixed": ("Fixed costs", "09", "fixed costs stay in accounting"),
    "summary": ("Finance overview", "09", "finance overviews and mappings stay in accounting"),
    "down_payment": ("Customer down payments", "09",
                     "down payments received in advance are a liability, not receivables; they stay in accounting"),
    "aging": ("Laepple 2023 ledger", "09", "the 2023 Laepple receivables and payables stay in accounting"),
}
RECEIVABLES_REASONS = (("aging", "aging"), ("bank cash flows", "bank"), ("accounts payable", "payable"),
                       ("fixed costs", "fixed"), ("finance summary", "summary"), ("down payments", "down_payment"))

# A sheet counts once, for the first family that claims it (a regional customer list also looks like a contact list).
SHEET_PRIORITY = ("quotes", "accounts", "contacts", "todos", "visits")


def folder_hint(relative_name: str) -> str:
    """The family folder ("06") named by a folder the browser reported, else ""."""
    for part in relative_name.replace("\\", "/").split("/")[:-1]:
        m = FOLDER_NAME.match(part.strip())
        if m:
            return m.group(1)
    return ""


def _sheet_family(sheet, path) -> dict:
    """{family: layout} for the families whose header this sheet has."""
    found = {}
    for key, fn in (("quotes", lambda: quotes.looks_like_quotes(sheet) and "quotes"),
                    ("orders", lambda: orders.looks_like_orders(sheet) and "orders"),
                    ("receivables", lambda: receivables.looks_like_receivables(sheet) and "receivables"),
                    ("accounts", lambda: accounts.detect(sheet)),
                    ("contacts", lambda: contacts.detect(sheet)),
                    ("todos", lambda: todos.detect(sheet)),
                    ("visits", lambda: layouts.detect(sheet, path))):
        try:
            layout = fn()
        except Exception:  # one odd sheet must not hide the file
            layout = None
        if layout:
            found[key] = layout
    return found


def _kept_out(kind: str) -> dict:
    what, folder, why = KEPT_OUT[kind]
    return {"kind": "kept_out", "family": None, "folder": folder, "what": what, "layout": "", "target": "Not imported",
            "reason": why, "outcome": f"Recognised — kept out of the CRM on purpose ({why})."}


def _receivables_kind(reason: str) -> str | None:
    low = reason.lower()
    return next((k for needle, k in RECEIVABLES_REASONS if needle in low), None)


def _import(family: str, folder: str, what: str, layout: str = "") -> dict:
    return {"kind": "import", "family": family, "folder": folder, "what": what, "layout": layout,
            "target": FAMILIES[family][1], "reason": "", "outcome": ""}


ORDERS_FILE = ("orders", "08", "ERP sales order list")
RECEIVABLES_FILE = ("receivables", "09", "AR / financial status (open invoices)")


def classify(path: str, hint: str = "") -> dict:
    """What the workbook at `path` is. `hint` is the folder number the browser reported (only used to
    say more in an 'unknown' answer; the content decides)."""
    name = os.path.basename(path)
    base = {"name": name, "display_name": file_display_name(name), "sheets": 0, "sheets_used": 0, "folder_hint": hint}
    try:
        sheets = load(path)
    except Exception as exc:
        return {**base, "kind": "unreadable", "family": None, "folder": "", "what": "Unreadable workbook",
                "layout": "", "target": "Nothing", "reason": f"could not be read ({type(exc).__name__})",
                "outcome": f"Could not be opened as an Excel workbook ({type(exc).__name__}). Nothing was read from it."}
    base["sheets"] = len(sheets)
    per_sheet = [_sheet_family(s, path) for s in sheets]

    # An order list or a receivables list is judged on the whole file: calculation workbooks hold a copy
    # of an order list, and finance workbooks hold bank sheets that look like a visit list.
    if any("orders" in f for f in per_sheet):
        kind, _ = orders.file_kind(path)
        if kind == "orders":
            return {**base, **_import(*ORDERS_FILE), "sheets_used": sum("orders" in f for f in per_sheet)}
        if kind in KEPT_OUT:
            return {**base, **_kept_out(kind)}
    if any("receivables" in f for f in per_sheet):
        kind, why = receivables.file_kind(path)
        if kind == "receivables":
            return {**base, **_import(*RECEIVABLES_FILE), "sheets_used": sum("receivables" in f for f in per_sheet)}
        if _receivables_kind(why):
            return {**base, **_kept_out(_receivables_kind(why))}

    tally, layout_of = Counter(), defaultdict(Counter)
    for f in per_sheet:
        for key in SHEET_PRIORITY:
            if key in f:
                tally[key] += 1
                layout_of[key][f[key]] += 1
                break

    def described(family):
        folder, what = LAYOUT_TABLES[family][layout_of[family].most_common(1)[0][0]]
        return {**base, **_import(family, folder, what, layout_of[family].most_common(1)[0][0]),
                "sheets_used": tally[family]}

    if tally and tally.most_common(1)[0][1] * 2 >= len(sheets):
        return described(tally.most_common(1)[0][0])
    # No family covers half the sheets: maybe a finance or order workbook with a few look-alike sheets.
    kind, _ = orders.file_kind(path)
    if kind in KEPT_OUT:
        return {**base, **_kept_out(kind)}
    kind, why = receivables.file_kind(path)
    if kind == "receivables":
        return {**base, **_import(*RECEIVABLES_FILE)}
    if _receivables_kind(why):
        return {**base, **_kept_out(_receivables_kind(why))}
    if tally:
        return described(tally.most_common(1)[0][0])
    where = f" It was in folder {hint}, but no matching header was found." if hint else ""
    looked = ("no sheet has the header columns of a visit report, customer list, contact list, to-do list, "
              "sales-order list or receivables list")
    return {**base, "kind": "unknown", "family": None, "folder": "", "what": "Not a list this importer knows",
            "layout": "", "target": "Nothing", "reason": looked,
            "outcome": f"Not recognised: {looked}.{where} Nothing is read from it."}
