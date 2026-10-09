# excel/ — reader for the real visit workbooks

Background: `../excel-survey-2026-10-09.md` (layouts, quirks). Plan: `../excel-import-plan-2026-10-09.md`.
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
