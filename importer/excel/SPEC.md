# excel/ — reader for the real visit workbooks

Background: `../docs/excel-survey-2026-10-09.md` (layouts, quirks). Plan: `../docs/excel-import-plan-2026-10-09.md`.
Output is the existing `visits.reader.Parsed` (Visit + Week), so `visits.match` and `visits.pipeline`
are reused unchanged for matching and commit. Real data never goes into git (`excel/gold/`, `excel/out/`
and the Excel folder are gitignored); tests build tiny synthetic workbooks with openpyxl.

Run everything from `importer/` with `.venv/bin/python`. Tests: `.venv/bin/python -m pytest -q tests`.

## Modules

| file | job |
|---|---|
| `sheet.py` | `load(path) -> list[Sheet]`. xlsx via openpyxl `data_only=True` (cached formula values); `.xls` via xlrd. `Sheet.value(r,c)` (1-based, merged cells return the top-left value), `Sheet.is_merged_continuation(r,c)`, `Sheet.formula(r,c)` (xlsx only, second load without data_only, lazy), `max_row`, `max_col` capped at 40. |
| `headers.py` | `find_header(sheet) -> (row, {field: col})`: scan rows 1-30 × cols 1-40 for the first row whose labels map ≥ 4 fields including `date`. Labels normalized: lowercase, newlines→space, collapse spaces, strip `：:`. Synonyms below. |
| `dates.py` | `parse_date(value, ref: date) -> (iso or None, flag)` and `parse_due(value, visit: date) -> (iso or None, phrase, flag)`. Rules below. |
| `layouts.py` | `detect(sheet, path) -> "mock" | "weekly2026" | "lucas" | "detail" | None`. |
| `readers.py` | one function per layout returning `(list[Visit], Week|None, list[ReviewItem])`. `mock` delegates to `visits.reader.parse_sheet` unchanged. |
| `split.py` | `split_customer(cell) -> (name, address, sure: bool)`, `phone_text(value)`. |
| `companies.py` | master index + `match(name) -> Match(code, cn, en, score, method, candidates)`. |
| `run.py` | `read_paths(paths) -> RunResult(parsed, review, skipped_files, per_file_stats)`; dedupe; never aborts on one bad file. |
| `accounts.py`, `contacts.py`, `todos.py`, `orders.py`, `receivables.py` | the other kinds of list (folders 04-05, 06, 07, 08, 09): `read_paths(paths)` / `build_book(paths)` -> plain records + review items; `detect(sheet)` or `file_kind(path)` tells whether a file is that kind. Their sections are further down. |
| `accounts_margince.py`, `contacts_margince.py`, `todos_margince.py`, `orders_margince.py`, `receivables_margince.py` | the Margince mapping of each: `preview` / `plan` (reads the CRM, writes nothing) and `commit` (idempotent). `webui/families.py` wraps them for the site; `webui/entry.py` is the front door that picks the reader of a file. |
| `decisions.py` | the owner's "same company?" answers (`company_decisions.json`), shared by every kind of list. |
| `cli.py` | `python -m excel.cli --mode parse|dry-run|commit <globs>`; dry-run writes `excel/out/report.json` + `excel/out/review.json`; commit reuses `visits.cli` flow (`visits.match.match`, `visits.pipeline.run`). |

## Header synonyms (normalized label startswith / equals)

- date: `date`, `visiting date` (never `due date`)
- cw: `cw`
- sales (owner per row): `sales`
- plan: `plan`
- city: `location/city`, `location`, `city`
- customer: `customer (full address if new)`, `customer` (but NOT a label containing `new, exist`), `visiting customer` is NOT customer (Lucas column B = row label, ignore)
- potential: `potential`
- customer_type: label containing `new, exist`
- contact: `contact name`, `name`
- phone: `phone no.`, `phone`
- office: `office`
- wechat: `wechat`
- email: `email`
- industry: `industries`, `industry`
- purpose: `purpose_topics visit`, `purpose for visiting`, `activity`, `purpose`
- competitor: `competiton products`, `competitor`
- communication: `communication content`
- next_step: `next step / action`, `next step`
- follow_up: `follow up`
- due: `due date`, `deadline`
- status: `status`

## Dates

`parse_date(value, ref)` — `ref` = Monday of the sheet's CW (or file-name CW) when known, else None.
- datetime/date → that day. Number (xls serial) → `xlrd.xldate`.
- Text: `d/m/yyyy`, `m/d/yyyy` (ambiguous → pick the reading closest to `ref`; if both > 10 days from ref
  or no ref and ambiguous → flag `ambiguous`), `yyyy/m/d`, `yyyy-m-d`, `yyyy.m.d`, `d-mmm`, `dd/mmm`,
  `mmm d`, `13-May`, `26th,May`, `15th June 2026`, `June 15th 2026`, `1st June--5th June2026` (range → first day, flag `range`).
  Missing year → year of `ref`, else the workbook year from file name (`2026`/`2025` anywhere), else flag.
- Formula text left over (`=D9+1`) when no cached value: resolve `=<CELL>+N` recursively once via the caller (readers pass a resolver); else flag `formula`.
- `/`, `-`, empty → (None, "empty").

`parse_due(value, visit_day)`:
- datetime → iso. Text dates as above with ref=visit_day.
- Float/numeric text `M.DD` (`5.25`, `6.5`, `4.15`) → month.day; `5.1` is ambiguous (1st or 10th) → pick the candidate on/after visit_day nearest to it, flag `decimal`.
  Year = visit_day.year (+1 if that lands > 60 days before the visit).
- Anything else → (None, original phrase, "phrase") so `visits.when.resolve_due` can try it later.

## Layout readers

Common: owner = cell right of the `Owner：`/`Owner :` label in rows 1-8 (normalize to one of the reps:
Jason Zhao, Devin Guo, Mark Wang, Sean Wang, Ethan Lv, Jason Yi, Andrew Zhang, Lucas Guo — by
substring/first name, else folder `Weekly reports/<rep>/`, else the `sales` column). Week number = ISO
week of the visit date; sheet-name/file-name CW is only `ref` and a sanity check (|diff| > 1 → ReviewItem
`week_mismatch`). Year from the visit date.

**weekly2026** (header row ~8, column B start; variants: +communication, Andrew `plan`/`activity`, no cw):
- A visit starts on a row where `customer` (or `date` if customer empty) has a value that is not a merged
  continuation. It ends before the next start. Rows with nothing in customer/date are continuation rows.
- Collect over the block: every non-empty `next_step` cell → one `Note(text, highlight="green",
  deadline=<due iso or phrase from the same row, else the block's first due>)`; `purpose` and
  `communication` continuation cells are appended (newline-joined) to the visit's purpose / extra.
- Visit fields: date_raw/date_iso, city, customer_raw (= name part from `split_customer`), potential
  (A/B/C else ""), customer_type_raw, contact_raw, phone_raw (`phone_text`), office_phone, email,
  purpose. Everything else → `visit.extra` dict: wechat, industry, competitor, communication, status,
  address, sales, plan, cw_cell, source_cells (field → "D9").
- Skip sheets with no header (empty `CW07`, `Sheet1`) silently (counted in stats).

**lucas** (`Weekly Sales Planning and Reporting` file name; header row ~14; columns shifted; 4-row blocks):
same block logic; `follow_up` → extra; key plan: rows above the header whose column C is a weekday name
(`Monday`, `Tue` …) → `Week.days` with the text in column D; week = ISO week of the first visit date
(the `Timing` cell is stale in CW11-15 files — never trust it).

**detail** (`Customer Visits in Detail` in file name; one meeting per sheet): customer = first non-empty
cell right of the label `Customer name:`; participants = rows from the `Participants` label down to the
first blank row: name (first non-empty cell right of the label column) + role (next non-empty cell).
Item table: header row containing `ID` and `General topic`; each row with Date or ID starts an item; rows
below with only Details continue it. One Visit per sheet: date = first item date (else sheet-name `m.d`
with the year from the file name), purpose = General topic(s), each item's Details → Note(highlight=None);
items with `WV` (resubmission date) → Note(highlight="green", deadline=WV) assigned to `Wer`.
contact_raw = "" (participants mix both sides) → participants in extra + ReviewItem `participants`.

## Customer cell split (`split.py`)

newline → first line name, rest address. Else name = text up to and including the first of
`有限公司`, `股份有限公司`, `Co., Ltd.`/`Co.,Ltd`/`Ltd.`/`GmbH`; the rest (stripped of `,;，、 ()（）`) is the address.
Else if the cell contains an address marker (`省`, `市`, `区`, `路`, `号`, `No.`, `Road`, `District`) after
≥ 2 chars → split there, sure=False. Else name = whole cell, sure = len ≤ 40.

## Companies (`companies.py`)

Master = `mail/staging/files/*/*Customer_List_20260104*.xlsx` (header row 2: Customer Code, CN name,
EN name, Province, Region, Area, Previous / Current Supporting, Dealer, First Order Time…) plus the ERP
`.xls` lists `../Excel Files from Sales Meetings/**/*Customer List_*Region*.xls` (header row 1 or 2:
客户编码, 客户名称, 客户简称, 地区名称, 地址, 联系人, 电话, 手机, 专营业务员名称). Merge on customer code.
Normalize: lowercase, full-width→half-width, drop spaces/punctuation, drop legal suffixes (有限公司,
股份, 集团, co ltd, co., ltd, gmbh, inc), drop city prefix in brackets (`(上海)`).
Match order: exact normalized on CN/EN/简称 → 1.0 `exact`; unique substring (≥ 4 chars, either way) →
0.9 `contains`; rapidfuzz `token_set_ratio`/`partial_ratio` best unique ≥ 92 and ≥ 5 points above the
2nd → score×0.95/100 `fuzzy`; else 0 `none` with top-3 candidates. Auto-accept ≥ 0.9; else ReviewItem
`company`.

## Dedupe and stats (`run.py`)

1. Identical file bytes → keep one (prefer the `Weekly reports/<rep>/` path).
2. Visits keyed `(owner, date_iso, normalized customer)`; when the same key comes from several files the
   newest wins (higher CW in the file name, then mtime). Lucas files whose visits are all duplicates of
   an older file are reported as `stale_copy`.
3. Stats per family: files, sheets, visits found, auto-accepted, queued (by reason), skipped files (+ error).

ReviewItem = `{family, file, sheet, row, field, value, reason, candidates?}`. Reasons: `no_date`,
`ambiguous_date`, `range_date`, `week_mismatch`, `company`, `customer_split`, `participants`, `due_decimal`.
A visit is auto-accepted when it has a date, an owner, a customer and no ReviewItem other than `due_decimal`.

## Accuracy test

`excel/gold/gold.json` (gitignored, "needs Jonas check"): list of `{family, file, sheet, row, expected:
{date, owner, customer_name, city, contact, phone, email, potential, customer_type, purpose_start,
next_step_count, first_due}}`. `tests/test_excel_gold.py` (skips when the file is missing) runs `run.read_paths`
on the gold files and prints accuracy per family and per field; asserts ≥ 0.95 per family.
`purpose_start` = first 20 chars, whitespace-collapsed.

## Contact lists, folder 06 (`contacts.py`, `contacts_margince.py`)

Own module, not part of the visit `Parsed`. Owner's scope (`CRM_Folder_Review.xlsx`): customer contacts
belong in the CRM; the Chinese lists are finance-driven (receivables / payables) and the FIBRO handover
list is internal, so both are imported but tagged so they can be filtered or deleted.

API: `contacts.detect(sheet)` -> `"list" | "sections" | None`; `contacts.read_paths(paths, master=None)`
-> `ContactsResult(contacts, orgs, review, skipped_files, per_file_stats)`; `contacts.summarize(result)`
-> counts only. CRM: `contacts_margince.contacts_plan / contacts_preview / contacts_commit` (re-exported
by `excel_margince`; CLI `--family contacts`). Plain records first (`Entry`, `Org`), CRM mapping second.

Header found by label (EN/ZH, any column order, header row within the first 15 rows, needs a contact or
company column plus two more fields; a header with a date is a visit list). Two layouts:
- `list` (company column: 公司名称 / Company ...): one row per company and person. The unlabelled 应收/应付
  and 客户/供应商 columns are recognised by their values (`应收 预收 其他应收` = customer, `应付 预付
  其他应付` = supplier). A row without a company is skipped (`no_company`).
- `sections` (no company column): column A carries a team or supplier label down to the rows below. A row
  with only that label is a heading (`section`, skipped). Under a heading that names FIBRO the people are
  internal colleagues of the group company named in the heading; elsewhere the label is a supplier
  (`company_from_label`: a Chinese registered name wins, else a bracket that names a legal form or is not a
  service word, else the text before the bracket).

Fields: company, full name, emails, phones, position, remark, department (team label of an internal
colleague), address, kind (customer / supplier / internal), ledger type as written, file/sheet/row.
Phones: `T`/`Tel` labels stripped, fax dropped, `visits.pipeline.normalize_phone`; a number without
country code (a German one without `+`, a local number without area code) is not guessed: it stays out
and is reported (`phone_no_country`). A bad email is reported (`bad_email`), the person kept. Contact
data without a name is not a contact (`no_contact_name`). Placeholder cells ("无", "-") are empty.

Copies: byte-identical files are read once; the same person at the same company (across copies and rows)
is one record, the richest copy is the base, a name-only copy never wins and fills nothing, blanks in the
base are filled from the others, a differing position or address is reported. A blank copy adds no
contacts. Source id: `contact:<company key>:<person key>`.

Companies go through `excel.companies.Master` like the visits: >= 0.9 is filed under the master's name,
a score below with candidates is a `company` review item (the company and its people stay out until
answered via `excel/decisions.py`), a name the master does not know at all is a new account. Internal
colleagues are not matched; they sit on the account `<group name> (internal)` (relationship type
`other`), never on a customer of the group's name.

Tags in the CRM (picklist, so the owner can filter by it): company "Contact list origin", contact
"Contact origin" = `Finance list: customer | Finance list: supplier | Handover list: supplier |
Handover list: internal colleague`. Company "Ledger type" = 应收/应付 as written. Contact "Department" and
"Remark" as written. Relationship types `customer` / `supplier` / `other` (internal); lifecycle
`customer` for a receivables company that has none. Existing CRM values are never overwritten. Address
goes to `address.line1`.

Kept out by code (`SENSITIVE`, `CREDIT_CODE`, and no column for them is read): bank accounts, tax IDs,
VAT/IBAN/SWIFT, credit codes. Free text (position, remark, address) holding some is withheld and reported
as `withheld_sensitive` (without its value). Tests: `tests/test_excel_contacts.py`.

## To-do and sales topics, folder 07 (`todos.py`, `todos_margince.py`)

Owner's scope (`CRM_Folder_Review.xlsx`): tasks with owner, due date, status, priority belong in the CRM; the
JET-IT list "looks like an IT project list, not sales: check it". Real layouts (4 workbooks; the sheet is
formatted far below its data, so row counts mean nothing): `Projekt-OPL` sheet = two merged title rows
(`TOPIC - <title>`, `Project Responsible:`), header on row 3, data from row 4, plus a `Dropdown` sheet that
only lists the Status (Open / WIP / Done) and Prio (1-3) choices. Columns: No., Categorie, Date, Topics,
To Do / Next Steps, Responsible, Due Date, Status, Prio, Remarks, Documents or Websites. The one-on-one list
has no Categorie column; the IT list has no No./Date/Prio and rows with only a topic and a status; the blank
to-do template holds nothing but stray status and date cells. The sales-topics template is another shape:
`Name: XXX`, then Topic / Due date / Date provided (`2026-x-xx` = not yet).

Reader (`todos.read_paths(paths, master=None, include_it_list=True)`): header found by label in rows 1-30, EN / ZH / DE
synonyms (`SYNONYMS`), any column order; only mapped columns are read, so an unmapped column (bank, tax id) never
reaches a record, and a text cell that holds bank or tax details is withheld (`withheld_sensitive`, without its value).
Layouts: `todo` (owner / status / prio / to-do column) and `topics`.

| field | rule |
|---|---|
| subject | Topics cell, else the first line of the To Do cell |
| owner | Responsible cell, else the sheet's `Project Responsible:` / `Name:` label (placeholders XXX empty); `readers.normalize_owner` to the reps. A name that is not exactly one rep (`Chenley`, `Ethan, Mark`) stays unassigned, kept as written, note `owner_unknown` |
| status | `map_status`: done / cancelled close the task; open / WIP / waiting stay open (EN, ZH, DE words; empty = open; an unknown word = open + `unknown_status`). Topics: done once Date provided is a real date |
| prio | the Prio cell as written (1-3) |
| due | `dates.parse_due` against the row's list date, else the sheet's earliest date: a date is kept; `5.25` is 25 May (`due_decimal`); an ambiguous day/month text gets no date, the text stays and `ambiguous_due` is queued; `next week` / `CW 14` via `visits.when.resolve_due`; `TBC` or `2026-x-xx` give no date |
| list date | the Date column; the task's `occurred_at` |
| company | a name of the master list in topic / to-do / remarks with score >= 0.9, exactly one company; two = none (`company_ambiguous`); never guessed. The IT list is not linked |

Rows without topic and to-do text are skipped (`no_text`; the template's stray cells); a merged one-cell row is a
section heading (`section_heading`, becomes the category of the rows below). A recognised list with no task says
so (`no_tasks`: "blank template") and is not an error; the Dropdown sheet is not reported as an unread sheet.

Tags (task subject prefix `[needs-check: ...]`, a `Tag:` body line, `raw.tags`): the IT project list (file name has
the word `IT`, or a Responsible cell says JET-IT; one `it_project_list` review item; `include_it_list=False` leaves
it out, rows reported `it_list_excluded`) and a topics sheet that names no rep.

Source id: `todo:` / `topic:` + sha1[:12] of (kind, owner, folded topic, folded to-do text); identical rows in one
sheet get `#2`. Status, due date and row position are not in it: a re-export with a new status or one more row
adds only what is new, and an existing task is never changed. Copies merge on the id (richest copy, then the
newer file name date); byte-identical files are read once; a blank copy yields no task.

CRM (`todos_margince.py`, source_system `todolist`): `POST /v1/activities` kind `task` (`/tasks` has no source id).
The CRM task has assignee, due, done; it has no priority, so the body's first line is `[Prio 1] [Status WIP]
[Responsible Ethan Lv]`, then the to-do text, category, remarks, the link / note cell, `Due as written`, `Listed
on`, list title, source file / sheet / row. Assignee = the seat named like the rep, else none (the rep stays in the
body). A company link only when the CRM already holds the matched company (none is created). Done tasks are
closed right after creating. `preview` reads only; `commit` never stops on one refused task (`failed`).
Tests: `tests/test_excel_todos.py`.

## Family 08 — Sales Order Exports (`orders.py`, `orders_margince.py`)

Owner's rule (`CRM_Folder_Review.xlsx`): only the ERP sales-order lists, as read-only order history per account;
invoice / packing lists and the formula workbook are not CRM data; orders live in the ERP.

**What a file is** (`orders.classify`, by content, never by the file name; `read_orders(path).kind/.reason`):
`orders` (a sheet with a document-number header), `invoice_packing` (sheets with `Invoice #:` / `Packing List`),
`stock_list` (`Inventory`), `formula_workbook` (other sheets hold formulas that read an order sheet; the order sheet in
it is a copy), `unrecognised`. Reasons: `not CRM data: invoice/packing list` / `inventory/stock list` / `formula
workbook (calculations over a copy of the order list)`. Side sheets of an order file are listed in `excluded_sheets`.
In the real folder 9 files = 5 order lists + 2 invoice/packing workbooks (called `..._Order_List`) + 1 stock list + 1
formula workbook (`Actual_mit_Formeln.xlsx`).

**Columns** are found by label (EN / ZH / DE). A column naming a bank account or tax id (`银行`, `账号`, `纳税人`, `Bank`,
`IBAN`, `Tax ID` ...) is never read, not even into `extra` (`not read (bank account or tax id)` warning). Payloads carry
only whitelisted fields (never `extra`).

**One order per document** (`build_book(paths, master=None) -> OrdersBook`; plain `Order` / `OrderItem` records):
- Byte-identical files are read once. Sheets of one file that repeat a document (a "completed" or "open" cut of the
  total list) add each distinct line once; lines only one sheet holds are kept.
- Several exports hold a document: the NEWEST export's version wins in full (export stamp `导出时间 Output Time` above
  the header, else the day of its newest order, then mtime). Same stamp (a hash-prefixed copy of one export): the richer
  copy wins. A copy with no customer or no amount at all never replaces a complete one and never creates a record
  (`blank_order`); a zero-value order is an order. Documents only an older export holds stay (history).
- Key = document number; when two selling companies use one number they are two orders (id prefixed with the seller).
- Status words (`开立/Open`, `待发货/To Ship`, `已完成/Completed`, `Partially Received`, `Cancelled` ...) -> `status_key`.

**Currency** is never guessed and never converted. The Chinese entity (`Sales Organization` contains `有限公司`) states
local-currency totals (`本币含税金额`, `表体本币…`, `Amount Incl. Tax in Functional Currency`; the document total is
repeated on every line and counted once): they are `cny_net` / `cny_gross`. `currency = "CNY"` only when the transaction
amount equals them; an order in another currency shows the CNY value and says the own currency is not named. Other
entities (the US Inc.) state no currency: amounts are shown as "currency not stated in the export" (review item
`currency_not_stated`, not blocking) and are never summed.

**Customers** go through `Master.match` like the visits (>= 0.9 files under the master's name; `new` creates an account;
anything else is a `company` review item and its orders stay out until a person answers in the existing company
review). Orders without a date are asked about (`no_date`) and not written.

**CRM** (`orders_margince.py`, source_system `erp-orders`): the CRM has no order object, and a won deal would feed the
forecast and won-revenue figures a second time, so no deal is touched. Each order is ONE NOTE on its account
(`order:<document no>`, dated the order day): status, value, payment, planned shipping, product lines, a "read-only
copy of the ERP" line, `ERP export: <stamp>`. Each account also gets ONE overview note (`orders-summary:<company key>`)
that sets the orders per year (by order date, cancelled out, CNY totals, not-completed part) beside the account's open
deals (`GET /v1/deals?status=open`, per currency, not converted): the actual-vs-pipeline figure. Missing accounts are
created through the visit pipeline's company import (`AccountRun`: customer relationship + lifecycle).
`plan(client, book)` reads only; `commit(client, book, on_progress=None)` writes. A second run writes nothing; a later
export updates the notes whose text changed; an older export never overwrites a newer note (the `ERP export:` stamp);
an overview is only replaced by a run that knows at least the orders it counts. One refused note is an entry in
`errors`, the run goes on. Tests: `tests/test_excel_orders.py`.

## Family 09 — AR and Financial Status (`receivables.py`, `receivables_margince.py`)

Owner's rule: only open / overdue AR per customer as a read-only number on the account page. Bank cash flows, AP,
fixed costs and the 2023 Läpple data stay out (accounting).

**The one CRM sheet** is the invoice-level list (`AR` in the `…Financial_Status…` workbooks): a header with customer
name (EN and CN), `Sales invoice no.`, `Invoice Date`, `Due Date` and `Invoice Value Local Currency（RMB）`. Only those
columns are read (a remark, bank account or tax id beside them never reaches a record). Open = sum of the local
amounts; overdue = those with a due date before the as-of day (due today is not overdue). A row with an amount but no
invoice number is open, never overdue (`no_due_date`). `#N/A / No Invoice` placeholder rows and zero amounts are
skipped. The sheet's own two totals are checked against the sums (a difference is a warning; the real files agree).
Currency = named by the amount header (RMB -> CNY); a list that names none, or has no date, is not usable.
**As-of date** = the sheet's `Today` cell (the day its due days were counted from), else the workbook's Overview
`Date:`, else the date in the file name. FIBRO group companies in the list are flagged `group_company` (still a
receivable).

**Excluded by code, with a reason** (`sheet_reason`, shown per sheet and per file): bank cash flows (`Cash`, `EBA`,
`BOC`, `ABC`), accounts payable (`AP`), fixed costs, finance summaries / mappings (`Overview`, `EN mapping`, `AR-AP
offset`), customer down payments (`Received in advance`, `downpayment`: money received before delivery is a liability,
not a receivable), aging-bucket ledger snapshots (`within payment terms`, `days overdue`: the legacy Läpple AR/AP
files), hidden sheets (the 84 monthly Läpple ledger sheets). The real folder: 4 files = 2 with the AR list (the
20260228 and 20260410 snapshots) + 2 excluded (`Laepple_RT_downpayment_202602`, `Laepple_RT_APAR_20231231`).

**Snapshots**: each file is a complete list of what is open, so the newest as-of date replaces the older ones (they are
reported `superseded_by`); byte-identical copies are read once.

**CRM** (`receivables_margince.py`): the CRM's finance card is a mirror of an accounting connector and cannot be written,
so company custom fields are created once and reused: `AR open (CNY)`, `AR overdue (CNY)` (type currency, whole fen)
and `AR as-of date` (date); a same-named field of another type is left alone (`… (receivables)` is used). The as-of date
is part of the value: a newer snapshot replaces the figures, an older one never overwrites a newer one (`accounts_kept_
newer`), the same figures on a newer date move the date, and an account holding a figure from an older snapshot but
missing from the newest one has paid, so it goes to 0 on that date (not while any customer of the snapshot waits for a
company answer). Every write is read back (`errors`: "the CRM did not keep ..."). If the CRM refuses custom fields
(403 / 409 / 422 / 501) each account gets one updated note instead (`erp-receivables`, `ar:<key>`, `As of:` line; same
newer-wins rule; no zeroing). `plan(client, book)` reads only; `commit(client, book, on_progress=None)` writes; a
second run writes nothing. Accounts are created through the same company import as orders. Tests:
`tests/test_excel_receivables.py`.

**Wiring** (for the single entry point): `orders.file_kind(path)` / `receivables.file_kind(path)` classify a workbook;
`build_book(paths)` + `summarize(book)` give the counts and the review items (`book.review`, `ReviewItem` with
`reason == "company"` = a person must answer); `<module>_margince.plan(client, book)` is the preview and
`<module>_margince.commit(client, book, on_progress)` the write. A web upload should hand `build_book` the whole folder
at once (the overview note of an account is only replaced by a run that knows every order it already counts).

## Customer lists, folders 04 and 05 (`accounts.py`, `accounts_margince.py`)

Folder 04 (ERP customer exports by region) and folder 05 (customer lists) become **accounts**, with the
opportunity data the 05 analysis workbook holds, plus **leads**. `companies.py` stays the matcher the visit flow uses; this
module reads the same files properly instead of only taking names from them (three of the seven files gave no master
record there: they are lists of another shape). Tests: `tests/test_excel_accounts.py` (synthetic workbooks; `.xls` via
xlwt when installed, else skipped; real-folder tests assert counts only and skip without the folder).

| layout | file | read as |
|---|---|---|
| `erp` | 04 `Customer List_*Region*.xls` (header 客户编码 + 客户名称) | account skeleton: ERP code, name, short name, English name (助记码), province, city, industry, address, ERP sales rep. Only the whitelisted columns of `ERP_TABLE` are read |
| `master` | 05 `Customer_List_*.xlsx` (Customer Code + CN name) | region, product area (Service/Motion/Cutting), dealer, supporting sales, first order (YYYY or YYYY-MM; `2020.1` is January or October, so only the year is kept). Code `New` = no ERP code yet |
| `supporting` | 05 `Customer_List_Xulu*.xlsx` (Customer + FIBRO Existing Customer) | name, city, province, industry, product category, existing Y/N, U8 rep (-> ERP rep), supported by (-> supporting sales), a note for the active flag |
| `effective` | 05 `Regional Effective Customer Analysis ... .xls` | several lists on one sheet, each under a title row: *old customers* (customer, + application, competitor, expected annual volume), *new customers* (prospect, same columns), *competitors at the exhibition* (competitor accounts; the shared "contact established / who recommends FIBRO" text becomes a note), *machine-tool makers who attended* (leads, with the turntable they may need) |
| `registration` | 05 `Customer_registration*.xlsx` | one lead per row; the technical requirements go into the lead's "Details" |

**Never read** (04): 纳税人登记号 (tax ID), 开户银行, 银行账号, 所属银行名称, 法人 (legal representative), 传真 (fax),
tax rate, payment terms, currency, shipping columns. **Read since 16 Oct 2026** (owner decision): 联系人 (contact), 电话,
手机 and Email地址 of the ERP export become a contact on the account, matched by customer code, with every +86 number
(`excel.split.phone_e164`); a row with contact details but no person's name becomes one note on the company
(`tests/quality/decisions/accounts.md` D-ACC-01, D-ACC-14). Because a value can sit in the wrong column (a tax number was found in an English-name column of
the real list), any 15-20 character digit-and-capital string or 10+ digit number in a text field is dropped and logged as
`id_number_in_text_field` without its value. The test proves it with sentinel values on synthetic workbooks, against the
records, the review items and every CRM payload, and again with the real folder's bank and tax numbers.

**Merge**: rows are grouped on the ERP code (`companies._canon_code`: `00CR991038` = `CR991038`). A row with no code is
matched by name with `companies.Master` (the visit importer's master plus every coded row of this read, saved answers
applied): a confident hit joins that account and the account is filed under the master's spelling; a doubtful one is
queued (`company`, blocking: left out of the commit until a person answers in the company-decisions file); a new name is
its own account (`name:<normalized name>`). Per field the richest row wins (ties: ERP, master, handover, analysis lists);
a value that differs becomes a `conflict` review item naming the winner, never a silent overwrite. Other names written
for the account are kept as aliases to find it in the CRM. Review reasons that do not block: `conflict`, `name_differs`
(different Chinese names under one code), `same_name_other_code` (one company under two ERP codes in the ERP),
`do_not_use` (industry cell says "Do Not use": no lifecycle, a note), `no_code`, `no_name`, `code_cell_has_text`,
`forecast_not_a_number`, `unrecognised_section`, `id_number_in_text_field`.

**Lifecycle**: `customer` for an ERP row (not "Do Not use"), a master row with a first order, a handover row marked
existing, an "old customers" row; `prospect` for the "new customers" list; nothing otherwise. An exhibition attendee the master
knows becomes a note on that account instead of a lead.

**CRM mapping** (decision: opportunity data are **custom fields on the account, not deals**. A deal needs a pipeline stage and
a money amount; the lists give units per year, and an invented stage or price would distort the forecast. If real deals
are wanted later they can be created from these fields):

| list | CRM |
|---|---|
| name, address, province, city, industry | company `display_name` (Chinese, else English), `legal_name` (the other language), `address.line1/city/region`, `industry` |
| ERP code, sales region, product area, dealer, supporting sales, ERP rep, first order | company custom fields "ERP customer code", "Sales region", "Product area", "Dealer", "Supporting sales", "ERP sales rep", "First order" (text); owner = the rep's seat for a **new** company only |
| application, competitor, expected annual volume | custom fields "Application", "Competitor" (text), "Expected annual volume" (number) |
| customer / prospect / competitor | `lifecycle`, `relationship_types` |
| contact name + phone (analysis lists) | a contact employed at the company (a phone needs a country prefix to be stored) |
| remarks | a `note` activity linked to the company, `source_id = note:<account>:<hash of text>` |
| attendees, registration | a lead (`source_id = lead:<normalized name>`), custom fields Phone, Location, Industry, Application, Details; skipped when the CRM already holds that name as a company |

**Idempotency and safety**: nothing is overwritten. A company is found by ERP code, else by exactly one name match
(Chinese or English, normalized; a company that already carries another ERP code is another customer; two matches = left
alone and listed). Blanks are filled; a different value is counted as a conflict and kept as it is. Notes and leads carry
content-based source ids; contacts are found by name at the company. A second run writes nothing. One failing record is
listed in the report and the run goes on. Preview reads the CRM only.

**Wiring**: `accounts.file_kind(path)` -> `("accounts", layouts)` or `("unrecognised", why)`;
`accounts.read_paths(paths, master=None) -> AccountsResult` (`.accounts`, `.leads`, `.review` as `ReviewItem`, `.skipped_files`,
`.per_file_stats`, `accepted_accounts()`, `accepted_leads()`; a `reason == "company"` item needs a person's answer);
`accounts.summarize(result)` and `accounts.file_report(result)` give counts only; `accounts_margince.plan_counts(result)`,
`accounts_margince.preview(client, result)`, `accounts_margince.commit(client, result, on_progress=None)` return dicts.
Hand `read_paths` all files of both folders at once (the merge works across files).
