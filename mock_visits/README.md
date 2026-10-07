# Mock visit lists (fictional data)

Purpose: develop and test the visit-list reader without access to the real Excel files.
All companies, people, phone numbers and e-mails are invented (phones use an unallocated `0000` block,
e-mails use `example.com`). Any resemblance to real firms is coincidence.

## Contents
- `excel/` one workbook per owner, one sheet per calendar week (same layout as the printed list)
- `ground_truth/` the expected parser output (what a perfect reader should return)
- `generate_mock_visits.py` regenerate or scale up: `python3 generate_mock_visits.py --owners 10 --weeks 26 --year 2025 --seed 3 --out .`

Default run: 4 owners, 56 sheets, 309 visits, 763 note lines, 112 customers.

## Sheet layout (per week)
- Row 1: `Owner :` name, `Timing :` CW number, `Summary`
- Row 2: `Item` / `Key Plan`; rows 3-7: Monday, Tue, Wed, Thu, Fri (sic) with the day's plan; Summary text at the right
- Row 9: table header, data from row 10
- Columns: Date | Location/ City | Customer | Potential (A, B, C) | Customer (new, exist, dealer) | Name | Phone | Office | Email | Purpose for Visiting | Next Step / Action | Dead line and follow up
- Columns A-J are merged vertically per visit. Columns K-L hold one note per row.
- K is filled teal (green in ground truth) for action lines, yellow for waiting lines. The meaning is a guess from the printed sample.

## Mess that is injected on purpose
- 4 date styles (`13-May`, `26th,May`, `May 13`, `13/05`), no year in the sheet (year is in the file name)
- 4 sheet-name styles (`CW20`, `CW 20`, `Week 20`, `W20`)
- Customer written in up to 5 ways (`Hengrui Automation`, `Hengrui`, `Hengrui Auto.`, `HENGRUI`, `Hengrui Co.`)
- Contact written as `Mr.Zhang`, `Zhang Wei`, `Zhangwei`, `Ms. Li`
- Phones: plain, with spaces, with dashes, with `+86`; many empty
- Customer type variants (`New`, `exist`, `Exist-motion customer`, ...), exhibitions, internal meeting rows
- Typos in notes, cancelled visits, follow-up visits that match an earlier note "Arrange visit in CW n"
- Public-holiday gaps (2025), customers shared between two owners, parent/child companies mentioned in notes
- Potential rating drifts over time for some customers

## Ground truth files
- `visits.csv` one row per visit: raw cell values, normalized values (`date_iso`, `phone_e164`, `customer_type`), `customer_id`, `contact_id`, `status` (done, cancelled, internal, exhibition), Excel row range
- `notes.csv` one row per note line: text, `highlight` (green, yellow, none), `deadline_followup`, Excel row
- `weeks.csv` one row per sheet: owner, CW, Monday date, key plan per weekday, summary lines
- `customers.csv` / `contacts.csv` the true entities with aliases, for testing the matching step
