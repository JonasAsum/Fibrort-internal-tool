# Plan: every Excel file into Margince, at 95 %+ accuracy (aiming for 100 %)

Based on `excel-survey-2026-10-09.md`. This plan assumes money is no constraint and that the work runs
in Claude Code wherever it can.

## 1. Where we stand

- **Files.** 362 spreadsheets found, about 60 of them CRM-relevant. They fall into 7 families: weekly sales
  reports (3 variants), Lucas's planning template, Customer Visits in Detail, ERP customer master (.xls),
  customer and contact lists, to-do lists, and order/AR exports.
- **Today's importer.** It reads the mock layout at 100 % and every real file at **0 %**. It is hard-wired
  to column positions, finds the week only through the sheet name and takes the year from the file name.
- **What already exists to build on:**
  - the mail importer's *reading layer*: two independent Claude readings, every fact quoted and checked
    against the source by code, disagreements sent to a review tab;
  - the web UI with its review tab;
  - idempotent `source_id` writes;
  - `mail.verify`, which reads the CRM back and fact-checks it;
  - Margince's MCP/REST agent surface, where every write is logged against a passport.

**Rule for the whole project:** a number such as "95 %" only means something when it is measured against a
hand-checked gold set. So the gold set is built first.

## 2. Target architecture: deterministic where possible, Claude where needed, people only on the residue

```
 files ─► 0 inventory+dedupe ─► 1 cell graph ─► 2 layout profile ─► 3 deterministic extract
                                                     ▲ (Claude proposes,          │
                                                     │  human approves once        ▼
                                                     │  per template)       4 normalise values
                                                                                   │
           5 Claude reader A ─┐                                                    ▼
           5 Claude reader B ─┼─► 6 merge + cell-citation check ─► 7 entity resolution
           5 vision reader C ─┘                                                    │
                                                                                   ▼
                       10 read-back verify ◄─ 9 commit (MCP/REST) ◄─ 8 review queue (web UI)
```

| Stage | What it does | Who / what |
|---|---|---|
| 0 Inventory | Content hash per file, sheet and row. Keep the newest file per rep×CW, drop byte-identical copies, flag stale copies (Lucas CW11-15 hold CW10 data). | Python |
| 1 Cell graph | Each sheet becomes JSON with value, cached formula value (`data_only`), number format, merge range, fill and hidden flag per cell, column scans bounded (Sean's 10 248 columns). `.xls` files are read through xlrd or converted with LibreOffice. | Python |
| 2 Layout profile | One YAML per template family: header row (found anywhere), column→field map by label, how a visit block is shaped (merged block, continuation rows, flat), where the week and the rep come from. **Claude drafts a profile once per new header fingerprint, Jonas approves it, and deterministic code applies it.** An unknown layout stops the run with a draft profile; it is never guessed. | Claude + human + code |
| 3 Extract | Applies the profile and keeps the source cell address on every value. | code |
| 4 Normalise | **Dates:** real dates, `dd/m/yyyy`, `=D9+1`, "15th June 2026", ranges, and `5.25` decimals read as month.day. The decimal reading is only accepted near the visit date. The week comes from the date. **Phones:** libphonenumber to +86 and WeChat IDs. **Potential and status:** Status free text or 0..1 becomes open / in-progress / done. | code + tests |
| 5 Readers | Fields that need understanding: split the customer cell into name and address; contact names such as 王总 or "Mr. Li (purchasing)"; turn each Next Step line into a task with a due date and assignee; competitor products; industry. **A** and **B** are independent Claude passes (different models or prompts, structured outputs) and must cite cell addresses. **C** is a vision pass over a rendered PNG of the sheet, which catches merged and colour meaning. | Claude (API batch or headless `claude -p --json-schema`) |
| 6 Merge | Code checks that every cited cell contains the quoted text and that every number appears in its source cell. A and B agree → auto. They disagree → review. A failed citation → rejected. This is the same rule set as `mail/READING.md`. | code |
| 7 Entities | Companies are matched against `Customer_List_20260104` (1 416 CN/EN names with codes) plus the ERP master (customer code, tax ID): exact, then alias, then pinyin/EN↔CN, then Claude adjudication with evidence. Optional: Firecrawl looks up the official Chinese legal name and address (Qichacha/Tianyancha/company site). Contacts are keyed on company + surname + phone. | code + Claude + Firecrawl |
| 8 Review | A new **Excel review** tab in the existing web UI. It shows the cropped sheet image beside the extracted record and highlights the source cells. Actions: Accept A / Accept B / Edit / Reject. Decisions survive re-runs. | Jonas / assistant |
| 9 Commit | Uses the existing idempotent pipeline (`source_system`, `source_id = file:sheet:row`), or Margince MCP with a passport so that every write is audited. | code |
| 10 Verify | Reads every record back from the CRM and diffs it field by field against the approved extraction, like `mail.verify`. | code |

**How this reaches 100 %:** the deterministic parts are exact. The parts Claude reads are double-read and
checked against their cited cells. Whatever is left (disagreements, low confidence, invariant failures) is
small enough that a person can look at every item.

**Invariants code checks before commit:**
- the visit date falls inside the sheet's CW (±3 days);
- the due date is on or after the visit date;
- Potential is A, B or C;
- the phone parses;
- the company resolves to exactly one record;
- one visit is not imported twice across file versions;
- every header and row of every sheet is accounted for, either imported or explicitly skipped with a reason.

## 3. Phased plan, each phase with its verify step

| Phase | Deliverable | Verify (done when) |
|---|---|---|
| **P0 Gold set** (1-2 days) | A Claude-assisted hand labelling of about 150 visit rows covering every rep and variant, 3 Visits-in-Detail sheets, 100 ERP rows and 50 list rows, stored in `excel_gold/` | Jonas has signed off every row |
| **P1 Inventory + cell graph** | `excel/inventory.py` and `excel/cellgraph.py`, plus a dedupe report | 362 files are inventoried and the duplicate counts match the survey (8 identical pairs, 7 changed rep-weeks, Lucas stale copies) |
| **P2 Layout profiles** | YAML profiles for 1a, 1b, 1c, Lucas, Visits-in-Detail, ERP master, Customer_List and the contact lists, and a generic reader that applies them | Every structural field (row boundaries, dates, columns) is 100 % against the gold set |
| **P3 Normalisers** | dates, phones, status, potential, customer-cell split, with property tests. The mock generator is extended to produce every real quirk (column offset, M.DD decimals, merged blocks, `27` sheet names) | Normaliser tests are green; the synthetic quirk corpus is 100 % |
| **P4 Reading layer** | readers A, B and C with structured outputs and cell citations, plus the merge and citation check | Auto-accepted fields are ≥ 99 % correct on the gold set, and the review queue is under 10 % of fields |
| **P5 Entity resolution** | company and contact matching against the master lists, plus Firecrawl enrichment | Companies ≥ 99 % against the gold set; zero silent merges (every ambiguous case is queued) |
| **P6 Review UI + commit + verify** | Excel review tab, commit, read-back verify | Read-back equals the approved extraction at 100 % |
| **P7 Full run** | All families imported into a fresh `make dev` CRM, then into the real one | Accuracy report per family ≥ 95 % auto, 100 % after review |
| **P8 Keep it clean** | Automatic intake plus a fix at the source (see §5) | New weekly files are imported with no manual step |

## 4. Running it in Claude Code

- **Package it as a skill:** `/import-excel <glob>`, built with `skill-creator`. It runs P1 to P3
  deterministically, fans out readers A/B/C, merges, and opens the review tab.
- **Subagents:** `layout-mapper` (drafts profiles), `reader-a`, `reader-b`, `vision-reader`, and `verifier`
  (read-only; the existing `verifier` agent type fits).
- **Workflow tool:** fans out across about 150 sheets in parallel, and resumable runs re-read only the sheets
  that changed.
- **Headless runs:** `claude -p --output-format json --json-schema …` or the Agent SDK (`output_format`
  json_schema, Pydantic models) for scheduled runs. Use the Message Batches API with prompt caching (the
  layout profile and schema are cached) for the bulk reads.
- **Hooks:** a PreToolUse hook refuses `--mode commit` unless the latest verify report for the batch is
  green and every review item is resolved.
- **Schedule:** a nightly routine (`/schedule`) picks up new files from the mail staging folder and the
  sales-meeting folder, runs the pipeline and posts "N visits imported, M waiting for review" to this
  project.
- **Margince MCP:** let the import agent write through Margince's governed agent surface, so every
  imported record is audited back to the passport and the source cell.

## 5. Ideas beyond a better importer

1. **A third, visual reader.** Render each sheet to an image with LibreOffice and have Claude read it with
   vision. It is independent of openpyxl's view of merged cells and colours, and three-way agreement is
   very strong evidence.
2. **Ask the source.** After each import, generate a short "this is what I read from your CW21 report"
   summary per rep (email or WeChat) and treat their corrections as ground truth. That is the only route
   to a real 100 %. Each message would need Jonas's approval before it is sent.
3. **Fix the template.** Lock the 2026 template so the header cannot be edited and Date is a real date
   cell. Add drop-downs for Potential, Customer type and Status, one row per visit, a separate address
   column, and a hidden `template_version` cell. Distribute it with the Claude for Excel add-in so reps
   get help filling it in.
4. **Skip Excel entirely.** Offer a visit form in Margince, a WeChat mini-program, or a voice note to
   Claude after each visit that becomes the CRM record directly. Margince's pitch is "a CRM that fills
   itself".
5. **Auto-intake.** If the reports move to OneDrive or SharePoint, Microsoft Graph delta query plus
   webhooks can import a file minutes after a rep saves it. Today they arrive by mail, which the mail
   importer already stages.
6. **Strategy tournament.** Use the `arena` skill to have several extraction strategies compete on the
   gold set (pure deterministic, pure LLM, hybrid, vision-first) and keep the winner per field.
7. **Adversarial tests.** Let Claude generate nasty-but-realistic variants (new rep, shifted columns,
   Chinese headers, ranges) to harden the layout detector before reps invent them.
8. **Company enrichment.** Firecrawl looks up each company's legal Chinese name, address and industry,
   which improves matching and fills empty CRM fields. Every enriched value is source-linked, as Margince
   requires.
9. **Accuracy dashboard.** A small artifact page showing per-family, per-field accuracy against the gold
   set and the review backlog, updated on each run.

**Cost guess:** about 1 000 visit rows × 3 readers plus profiles and adjudication comes to a few hundred
thousand tokens, so tens of dollars. Time, not money, is the constraint.

## 6. Skills to use, step by step

| Step | Skill (installed) | Why |
|---|---|---|
| Nail the decisions | `agent-skills:interview-me`, `brainstorming`, `agent-skills:idea-refine` | Settle scope (which families in v1), field mapping and review ownership |
| Written spec | `agent-skills:spec` (spec-driven-development) | `SPEC.md` holds the CRM field map per family, invariants and accuracy targets |
| Quality bar | `agent-skills:constraints` | `CONSTRAINTS.md`: ≥ 95 % auto, 100 % after review, no silent merges, no bank or tax IDs |
| Task breakdown | `agent-skills:plan` (planning-and-task-breakdown) | Turns §3 into ordered tasks with acceptance criteria |
| Reading Excel | `anthropic-skills:xlsx` | openpyxl/pandas patterns, formula recalculation, rendering |
| Claude calls | `claude-api` | Structured outputs, Batches, caching, model choice for readers A/B/C |
| Fan-out | `workflow-authoring` (+ Workflow tool) | Parallel per-sheet reading, resumable runs |
| Build | `agent-skills:build`, `agent-skills:test` (TDD, Prove-It), `agent-skills:doubt-driven-development`, `agent-skills:debugging-and-error-recovery` | Small verified increments; every quirk becomes a failing test first |
| Compare strategies | `arena` | Tournament over extraction strategies on the gold set |
| Research | `anthropic-skills:deep-research` + Firecrawl | Vendors and tools, company enrichment |
| Review | `agent-skills:review`, `code-review`, `security-review` | PII (bank accounts, tax IDs, phones) and idempotency |
| Ship | `agent-skills:ship` | Go/no-go before the run against the real CRM |
| Package | `skill-creator`, `update-config` (hooks) | `/import-excel` skill and commit-guard hook |
| Automate | `schedule`, `loop` | Nightly intake routine |
| Show results | `artifact-design`, `dataviz` | Accuracy dashboard |
| Find more | `find-skills` | Look for spreadsheet or ETL skills not yet installed |

## 7. Decisions needed from Jonas

1. **v1 scope.** Weekly reports + Lucas + Visits in Detail + customer and contact lists? Or orders and AR
   too (they would become deals and account health)?
2. **Rep confirmations.** May reps be asked to confirm their imported data (idea 2)?
3. **Future intake.** Change the template now (idea 3) or move reps off Excel (idea 4)?
4. **Target CRM.** Local `make dev` first, then which production instance?
5. **Who reviews.** Who works the review queue: you, or an assistant in Shanghai?
