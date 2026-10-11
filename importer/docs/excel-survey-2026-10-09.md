# Excel survey: what should go into the CRM (2026-10-09)

This survey only read files. No workbook, CRM record or git state was changed. The per-file profiles
(sheets, header row/column, header labels, row counts, merges, fills, validation lists,
formulas, hidden rows/cols, date formats, Chinese text) and the importer baseline are in
`excel-survey-2026-10-09.json`. In that file, header labels that occur in only one file, or that
look like company names or document IDs, are replaced with `<redacted>`. No cell values are copied.

**Scope:** `~/Mac Documents`, `~/Documents`, `~/Downloads`, `~/Desktop`, `~/Library/CloudStorage`
(two Google Drive mounts. There is no OneDrive/SharePoint sync folder). `node_modules`, `.venv`,
`venv`, `.git`, `build` and `dist` were skipped. One stray `.venv311/site-packages` (numpy test CSVs) was
picked up and filed under *unrelated*.

**Found:** 362 files (251 xlsx/xls, 111 csv). All of them were profiled, and 2 failed to open (an ERP
`Sales_Order_List__List_0325.xlsx` with a broken named style, and an eBay template).

**Method:** the header row is the first row with at least 3 text cells. Files were then grouped by
path plus header vocabulary. The baseline ran `importer/visits/reader.py` (`parse_sheet`, the same
code `cli.py --mode parse` uses, with no network call). For every row that has a customer cell, each
of 7 fields (date, city, customer, contact, phone, email, purpose) was compared with what a
header-name lookup reads from the same row.

## Families

| # | Family | Files (unique) | CRM-relevant | Rows (approx.) |
|---|---|---|---|---|
| 1 | Weekly Sales Reporting 2026 (one sheet per CW) | 29 (17 unique, 8 rep workbooks are byte-identical copies in `Excel Files from Sales Meetings/` and `Weekly reports/<rep>/`, 1 copy in mail staging) | **yes: visits** | 445 sheets, ~137 distinct rep×CW sheets, 636 visit rows in the newest file per rep |
| 2 | Weekly Sales Planning and Reporting (Lucas Guo, older template) | 9 | **yes: visits + key plan** | 38 visit rows |
| 3 | Mock visit list (importer test data) | 5 (4 + 1 web-UI upload copy) | test only | 309 visits |
| 4 | Customer Visits in Detail (meeting minutes per customer) | 3 (2 identical Devin copies + Downloads `…-20260918.xlsx`) | **yes: meeting notes/tasks** | 11 sheets, 1 meeting per sheet |
| 5 | ERP (U8) customer master by region, Chinese headers | 5 (+1 AR aging file with similar headers) | **yes: companies/contacts** | North 216 + North-west 74 + South 144 |
| 6 | Customer lists / registrations (hand-made) | 5 | **yes: companies** | `Customer_List_20260104` 1 416 rows, Xulu 30, Mark's "Effective Customer" 47, 1 registration form |
| 7 | Contact lists | 4 (2×2 duplicates) | **yes: contacts** | 25 + 14 |
| 8 | One-on-one to-do / sales-topic lists | 4 | maybe: tasks | 6-12 rows each |
| 9 | ERP sales-order exports (Sales Order List, OIT overview) | 10 | later: deals/orders | ~3 000 rows |
| 10 | AR / financial status per customer | 4 | later: account health | ~8 500 rows |
| 11 | Sales planning / forecast | 6 | no (not per customer) | |
| 12 | Quotations (CCMT exhibition, freight) | 12 | no | |
| 13 | Importer artifacts (CSV dumps, ground truth, verify reports) | 9 | no | |
| 14 | Other Fibro mail attachments (balance reports, invoices, customs, stock, payroll, HR) | 91 | no | |
| 15 | Log Book Heiko Asum (Google Drive) | 2 | no (other business) | |
| 16 | Unrelated (SEO exports, car tracking, numpy tests, Tongji vocab, personal) | 163 | no | |

### 1. Weekly Sales Reporting 2026: the main target

Example: `Excel Files from Sales Meetings/Weekly reports/Jason Zhao/Template for Weekly Sales Reporting 2026_CW21_Jason Zhao.xlsx`.
Reps: Jason Zhao, Devin Guo, Mark Wang, Sean Wang, Ethan Lv, Jason Yi, Andrew Zhang, plus the blank template.

Canonical header, **row 8, starting at column B**:
`Sales | CW | Date | Location/City | Customer (full address if new) | Potential (A, B, C) | Customer (new, exist, dealer) | Contact Name | Phone No. | WeChat | Email | Industries | Purpose_Topics Visit | Competiton Products | Next Step / Action | Due Date | Status`

Variants:
- **1a, standard (17 cols, B-R):** Devin, Mark, Sean, Ethan, Jason Yi, template.
- **1b, + `Communication content` (18 cols, B-S):** this extra column before Next Step comes from Jason Zhao's own copy and shifts Next Step, Due Date and Status one column right.
- **1c, Andrew:** `Plan | Date | Location/City | Customer | Potential | Customer type | … | activity (instead of Purpose_Topics Visit) | …`. There is no Sales or CW column, and Date is a text range such as "15th June 2026".

Rows 1-7 above the header hold only a title ("Weekly Sales Avtivity - Plan" in B1, merged B1:S2), `Owner：` in E4 with the rep in F4, and "Customer" / "Summary" headings in E6/I6. The key plan and summary blocks of the mock layout are empty here.

Quirks (counts are across every sheet of the family):
- **Header not in column A:** `Date` sits in D8 (C8 for Andrew). `reader._header_row` only checks column 1, so every sheet fails.
- **Sheet names:** `CW05`, `CW05 ` and `CW22 ` (trailing spaces), `27` with no prefix (Mark CW27 file), `W23`…`W30` (Andrew), and empty `Sheet1` / `Sheet3`. Jason Zhao's `CW07` and `CW08` are blank, with no header.
- **Year:** reader takes it from `_YYYY.xlsx` at the end of the file name. Real names end in `_CW21_Jason Zhao.xlsx` or `-Devin.xlsx`, so it falls back to today's year. That is correct only by luck in 2026.
- **Date column:** 985 real dates (format `[$-409]d/mmm;@`), 171 text dates `dd/m/yyyy` / `d/m/yyyy` (Sean), 28 formulas `=D9+1` (cached values are needed, so load with `data_only=True`), and ~15 ordinal-text dates ("15th June 2026", ranges such as "1st June--5th June2026"), plus `yyyy/m/dd` and `dd/mmm`.
- **CW column:** "CW27" text in 641 cells, a real date in 354 (Jason Zhao keeps the date there), bare numbers in ~100, formulas in 24, text `m/dd/yyyy` in 6. It cannot be trusted as a week number.
- **Visit block shape:**
  - Jason Zhao: B:P merged over 4 rows, one Next-Step line per row in Q. Some sheets have 95-117 merged ranges.
  - Devin: the Purpose text continues in N on the rows below, which carry no customer.
  - Mark: one flat row per visit, with no merges.
  - The block ends at the next row that has a customer.
- **Customer cell:** name plus full address in one cell (up to ~110 chars, Chinese and English mixed).
- **Phone No.:** stored as a number (loses formatting, risks float rendering). WeChat is often a phone number.
- **Due Date:** 379 real dates, 39 text `dd/mm/yyyy`, and **decimal numbers that mean month.day** (`5.25`, `6.5` = 25 May, 5 Jun), ~40 cells. There is also free text ("next week", "May.2026").
- **Status:** free text (Done / Open / On going / In progress…) mixed with numbers 0/0.5/0.6/0.8/1, which are progress fractions.
- **Potential:** A/B/C in 1 060 cells, "/" in 6.
- **Fill colours:** yellow `FFFF00` is used only on header cells and indexed white elsewhere. The green/yellow action highlighting that reader.py expects in column K does not exist here.
- **Hidden rows:** 2 in Devin's files.
- **Chinese text:** present in most reps' sheets (customer, purpose, communication).
- **Workbook size:** Mark's files are ~15 MB because of one 14.5 MB PNG. Sean's sheets have `max_column` = 10 248 (stray formatting, 2.5 MB of sheet XML each). Use bounded column scans.
- **Duplicates across versions:** each new file repeats every earlier CW sheet. 56 of 146 rep×CW sheets appear in more than one file, and 7 of those changed between versions (Jason Zhao CW05/06/10/11/12, Devin CW20, Mark CW27). Import only the newest file per rep, or dedupe by rep+CW and keep the newest content. The 8 byte-identical root vs `Weekly reports/` copies must not double-import.

### 2. Weekly Sales Planning and Reporting (Lucas Guo)

The sheet is always named `W01`. `Owner：` is in C4 (Lucas Guo), `Timing：` in H4 with the CW in I4. "Major Content in Office" sits in B6, the key plan for Monday-Fri in C7:C11 with text in D, and "Item Summary" in H6.

Header **row 14, column B**: `Visiting Customer | Date | Location/City | Customer | Potential (A, B, C) | Customer (new, exist, dealer) | Name | Phone | Office | Email | Purpose for Visiting | Next Step / Action | follow up | deadline`.

Each visit is a 4-row block (row 15, 19, 23 …), and the dates are real dates. This is the closest to the mock layout: the same columns, but shifted 2 columns right, with an extra "follow up" column and the header 6 rows lower.

Quirks:
- Every sheet is named `W01`, so the CW comes from the file name only.
- **The Timing cell says CW10 in the CW11-CW21 files.** The CW11, CW12, CW13, CW14 and CW15 files contain the *CW10 visits* (dates 2-6 Mar). They are stale copies, so dedupe by content, not by file name.

### 3. Mock visit list

Header `Date` is in A8, Owner/Timing/Summary are in row 1, the key plan sits in rows 2-7, and notes are in K with green `1E8070` / yellow fills. Sheet names: `CW 18`, `Week 17`, `CW14`, `W18`. The year is in the file name (`_2025.xlsx`).

### 4. Customer Visits in Detail

Example: `Excel Files from Sales Meetings/Devin Guo - 2026_Template for Customer Visits in Detail-202605 的副本.xlsx` (a copy of Jason Yi's template). A newer one is `~/Downloads/2026_Template for Customer Visits in Detail-20260918.xlsx` with a single sheet, `Tabelle1`.

Layout, one meeting per sheet, with sheet names that are dates (`5.8-9`, `5.11-15`, `5.18`):
- B2 holds the title.
- B3 "Customer name:" with the value in D3.
- B5 "Participants", with names in D5:D7 and roles in F5:F7.
- G9 holds the type legend (E = Entscheidung, I = Information, A = Abstimmung).
- The item table header is in **row 11, B-J**: `ID | Date | Typ (E/I/A) | from (Person) | General topic | Details | WV | Wer | Status`. Usually there is 1 item with Details text (up to 650 chars) continuing in G on the rows below.

The header labels are German/English (WV = Wiedervorlage, Wer = owner). Each sheet has `max_column` 152, which is formatting only. This is a key-value form plus an item table, not a list.

### 5. ERP customer master by region (U8)

Example: `Excel Files from Sales Meetings/Weekly reports/Lucas Guo/20260125_Customer List_South Region.xls`.

Header in row 1 (row 2 in Mark's copy, where row 1 is a title). The 30 Chinese columns include 序号, 客户编码 (customer code), 客户名称, 客户简称, 纳税人登记号 (tax ID), 客户分类, 地区名称, 所属行业, 地址, 开户银行, 银行账号, 法人, Email地址, 联系人, 电话, 传真, 手机, 专营业务员名称 (sales rep), 发货地址 and 省直辖市.

Quirks:
- `.xls` (BIFF, WPS-saved), so `xlrd` is needed (not in the importer venv).
- Devin's "（已联络）" (contacted) copy adds a `Purpose_Topics Visit` column: 24 of 216 rows have visit notes.
- Rows are colour-coded by background (6 palette indexes), which presumably marks contact status. This is not documented.
- It holds bank account / tax IDs, which the import should skip.

### 6 and 7. Customer and contact lists

- `Customer_List_20260104.xlsx` (mail staging), header row 2: `Customer Code | CN name | EN name | Province | Region | Area | Previous / Current Supporting | Dealer | First Order Time from 2020`, 1 416 rows. This is the best master list for company dedupe and EN↔CN name matching.
- `Customer_List_Xulu_251225_1.xlsx`: `序号 Item | 客户 Customer | Company English Name | City | Province | 行业 Industry segment | FIBRO Tables Category | FIBRO Existing Customer (Y/N) | Active Status… (Y/N) | have been supported by… | U8 Asigned to…`.
- Mark's "Regional Effective Customer Analysis….xls", header row 2, bilingual: `序号 no. | 客户名称 customer name | English name | 地址 adress | 联系人 Contact Name | 电话 phone | 主要应用型号 Main application models | 竞争对手 competitor | 预测全年采购数量 Expected usage quantity per year`.
- Contact lists: `Handover_Contacts_List_Monica` (`Contact Person | E-mail Addr. | Phone No. | Position | Remark`, header starts in column B) and `客户供应商联系方式` (`序号 | 公司名称 | 地址 | 联系人 | 电话`, with 应收 / 客户 section rows inside the table).

### 8-10. Later candidates

- **To-do lists:** `No. | (Categorie) | Date | Topics | To Do / Next Steps | Responsible | Due Date | Status | Prio | Remarks | Documents or Websites (Hyperlink)`, header row 3, with a `Dropdown` sheet that feeds the Status/Prio lists.
- **Sales Order List / OIT:** bilingual ERP export headers, more than 36 columns, with Customer, Salesperson, Document Date and amounts.
- **Financial Status:** about 16 sheets (Overview, AR, AP, …), AR header row 7 with a pivot block below it.

## Importer baseline (`visits/reader.py`, parse only)

| Family | Files run | Visit rows present | Extracted | Fields correct | Accuracy |
|---|---|---|---|---|---|
| Mock visit list | Chen Hao, Zhao Lei | 152 | 152 | 1 064 / 1 064 | **100 %** |
| Weekly Sales Reporting 2026 (1a/1b/1c) | Jason Zhao CW21, Mark CW27 + CW38, Sean CW38, Devin CW30, Ethan, Jason Yi, Andrew | 751 | 0 | 0 / 5 257 | **0 %** |
| Lucas planning | all 9 | 38 | 0 | 0 / 266 | **0 %** |
| Customer Visits in Detail | 2 | 11 meetings | 0 | n/a | **0 %** |
| ERP / customer / contact lists | not runnable | — | — | — | no importer exists |

The failures are hard failures, not partial ones:
- Every real sheet raises `no 'Date' header row found in the first 20 rows`, because `Date` is not in column A.
- Visits in Detail raises `sheet name carries no CW/Week number`.
- `parse_paths` stops on the first bad workbook, so in a batch run one real file aborts the whole run.

## What blocks 95 %+ accuracy (by impact)

1. **Fixed column positions.** reader.py hard-codes A=Date … J=Purpose, K=notes, L=deadline. The real templates start at column B (or C) and differ per rep (+Communication content, +WeChat/Industries/Competitor/Status, Andrew's Plan/activity, Lucas's Office/follow up). Detect the header row by scanning all columns for `Date`, then map columns by label rather than by position.
2. **Week and year identity.** Sheet names include `27`, `W23`, `W01`, and names with trailing spaces. The year is missing from file names, the CW column holds dates, numbers or text, and Lucas files have a stale Timing cell. Derive the week from the visit date where one exists, and use the sheet name or file name only as a fallback.
3. **Cumulative and duplicate workbooks.** Each weekly file repeats all earlier weeks, root and per-rep folders hold identical copies, 7 rep-weeks changed between versions, and the Lucas CW11-15 files are CW10 content. Dedupe on rep + visit date + customer, keeping the newest file.
4. **Date and due-date parsing.** Values arrive as real dates, `dd/m/yyyy` text, `=D9+1` formulas (need `data_only=True`), "15th June 2026" and ranges, and due dates written as `5.25` decimals (month.day) or free text.
5. **Visit block shapes.** Visits are 4-row merged blocks (Jason Zhao, Lucas), continuation rows for Purpose (Devin) or flat rows (Mark). Action lines live in the "Next Step / Action" column, not in K with green/yellow fills.
6. **Customer cell = name + address.** This must be split before company matching. Customer_List_20260104 (1 416 CN/EN names with codes) is the obvious reference for matching.
7. **Other layouts.** Visits in Detail (key-value form plus item table, one meeting per sheet) and the `.xls` ERP lists (need `xlrd`, Chinese headers) each need their own reader.
