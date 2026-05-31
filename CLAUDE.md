# CLAUDE.md — Financial Dashboard Agent (סטודיו מרים בסקון / SMB)

## Purpose
A deterministic, re-runnable engine that ingests the monthly Fireberry exports and renders a
frozen-style, **RTL, ₪** Excel dashboard at three zooms (monthly / quarterly / annual).
**The code computes every number; the generator builds every layout.** The agent orchestrates and
adapts — it never invents a number or designs a layout ad-hoc.

## Input file map (Hebrew filename → role)
| File | Role / key columns | Location |
|---|---|---|
| `לקוחות.xlsx` | leads — שם · תאריך הגעה · מקור הגעה · סוג לקוח | `inputs/current/` |
| `פרויקטים.xlsx` | projects/deals — שם · לקוח · שכ"ט · שכ"ט-אם-הופסק · תאריך התחלה · סטטוס | `inputs/current/` |
| `שתפים.xlsx` | partnership commissions — שם · סטטוס · תאריך שבוצע · שכ"ט · שכ"ט-אם-בוטל | `inputs/current/` |
| `הכנסות.xlsx` | cash receipts — תאריך · פרויקט · שת"פ · סכום (ברוטו) · חשבונית | `inputs/current/` |
| `משימות.xlsx` | tasks — מבצע · סטטוס · עודכן בתאריך | `inputs/current/` |
| `הוצאות קבועות.xlsx` | fixed expenses — קטגוריה · תיאור · סכום · סוג | **root** |
Keep these Hebrew names exactly — they are re-exported under the same names each period.

## Run
```
C:\Python314\python.exe scripts\run.py 2026-Q1    # quarterly (Jan–Mar + summary)
C:\Python314\python.exe scripts\run.py 2026-04    # monthly (April)
C:\Python314\python.exe scripts\run.py 2026       # annual
C:\Python314\python.exe scripts\verify_outputs.py # rebuild + correctness/structure report
```
Deps: `pandas`, `openpyxl` (required); `python-docx` (only to read a *filled* quarterly Word and to
regenerate the `.docx` guide). Each run writes its own period folder:
`outputs/<period>/` → `<name>.xlsx` + `.html` preview (e.g. `outputs/2026-Q1/2026-Q1_רבעוני.xlsx`).
**xlsx + html only — no PDF anywhere.** **No file sits directly under `outputs/`.** `verify_outputs.py`
writes `verification_report.md` to the project root; `scripts/tests/run_tests.py` is the pass/fail gate
(numbers · parity · reconciliation · expense-sheet · determinism · structure · business-logic).
Per-month input folders (`inputs/YYYY-MM/`) are the archive.

## Monthly run flow (`run.py`)
1. Locate inputs (`inputs/current/` + root expenses).
2. **Expense gate (human-in-the-loop):** today expenses = `הוצאות קבועות.xlsx` only ("no one-offs").
   When one-off expenses begin, the agent pauses here for confirmation before computing.
3. **Validate** (graded — below); stop on blocking.
4. **Compute** all metrics (deterministic).
5. **Targets:** if `inputs/meetings/<quarter>.docx` exists, pull its targets table; else skip.
6. **Render** a fresh, high-end workbook from `generator.py` + `style_config.py` (NOT a static template),
   all numbers sourced through `viewmodel.py` (the single parity source the HTML shares):
   a leading **`סקירה` (overview)** sheet — KPI cards (profit colored by sign) → the two income pictures
   side by side → trend/breakdown charts **with visible data labels** (incl. a monthly profit-trend line)
   → targets attainment → **month-by-month + summary analytics matrices** (expenses/profit, leads-by-source,
   tasks-by-performer, attribution — like the income pictures: ינואר · פברואר · מרץ · סיכום) → cohort
   (+ time-to-close caveat) → missing report — plus row-level, period-scoped **detail sheets**
   (`פרויקטים · לידים · הכנסות · שתפים · משימות · הוצאות`, frozen header + autofilter) and a hidden
   `נתונים` chart source. The `הוצאות` sheet replicates each fixed expense per month (row sum = the headline).
   A neutral-gray KPI Δ vs the previous period is shown only when a prior *reporting* period exists.
   The header carries the **studio logo** (HTML: on a white chip in the brand corner, logo-only; Excel:
   an image in the `סקירה` header) and the KPI row sits with clear **breathing room** below the tabs.
7. **Write** outputs into `outputs/<period>/` with fixed names (**xlsx + html only — no PDF**); the HTML is
   a self-contained, deterministic **tabbed interactive BI dashboard** (inline CSS + inline JS for behavior
   + inline SVG charts) built from the same data — **JS drives only interaction/display, never numbers** —
   so every figure matches the Excel by construction. Emit the missing-items report. On quarterly runs,
   also drop a blank next-quarter Word into `inputs/meetings/`.
8. **Publish & back up (both gated — never fail the run, same pattern as the old optional steps):**
   upload the period's `xlsx`+`html` to Google Drive (Shared-Drive `ניתוחים וישיבות הנהלה` →
   `דשבורדים/<period>/`, find-or-create) via the **GWS CLI** as `studio@smb-arch.com`; then
   `git add -A && commit && push` the whole project to the private **GitHub** backup
   (`studio139/smb-dashboard`). On missing tool / no auth / offline each prints a notice and continues.
   *(The GWS OAuth app is now in **Production** — the token no longer expires every 7 days.)* See `publish.py`.

## Determinism
Every number originates in `scripts/metrics.py`. The agent never estimates or hand-places a value.
Consistency = generator + style frozen → each period looks identical yet adapts to the data.

## Open schema (incl. employees / מבצע)
Structure (which dimensions, which metrics) and business logic (statuses, VAT, the two income
pictures) are **fixed**. The values inside each dimension — מקור הגעה, סוג לקוח, קטגוריה, סטטוס,
and **employee (מבצע)** — are derived from the data every run; a new value becomes a new
row/column automatically, in the existing style, without changing structure.

## Two income pictures (never merge)
- **Picture A — pipeline ("money closed"):** project fees + partnership fees.
- **Picture B — cash ("money in"):** receipts; net = gross ÷ 1.18 when `חשבונית=כן`, else gross
  (empty / `-` = not invoiced).
They stay in **separate blocks** — never combined into one figure.

## Business logic (fixed)
- Projects: `בהשהייה` = active (full fee); only `הופסק` → reduced fee (`שכ"ט-אם-הופסק`).
- Partnerships: `בוצע` + `נסגר - טרם סיום תשלום` = closed (full fee); `בוטל` → excluded (recorded).
- Tasks: count `סטטוס=הושלם` per `מבצע` per month (by `עודכן בתאריך`). The export has no task-id
  column, so identical (מבצע, status, day) rows are legitimate separate tasks — counted, not deduped.
- Conversion: cohort-anchored to the lead's **arrival** month; last ~3 months flagged
  `עוד מבשיל`; the headline `אחוז המרה` uses **mature cohorts only**. Companion: time-to-close.
- Join leads↔projects by name (`שם לקוח ↔ לקוח`); an attribution/reconciliation count is emitted.
- Ranges: leads from 9.25 (Sep–Dec 2025 = lookback for attribution + mature cohorts, **not** report
  months); projects/income/partnerships/tasks from 1.26; report period = 2026.

## Validation (graded — `scripts/validator.py`)
BLOCKING (stop + explain): whole file missing · required column missing · empty core data ·
impossible total (net>gross, conversion>100%). NON-BLOCKING (continue + record → `לא צוין`):
single missing field · unmatched join rows · single outlier. Every run ends with a missing-items report.

## Targets & timeline (`scripts/targets.py`)
Source = the targets **table** in the quarterly Word (`inputs/meetings/<quarter>.docx`); metric names
match dashboard metrics. Insights/initiatives are human text — never parsed for numbers. Q1 quarterly +
April monthly = no targets yet. At Q1 close a blank Q2 template is emitted; once filled, Q2 targets
show as target-vs-actual from month 5.

## Modules (`scripts/`)
`loader` · `metrics` · `validator` · `targets` · `viewmodel` (the single parity source — month+summary
blocks, derived per-month attribution, expense replication, detail/chart specs — consumed identically by
both renderers) · `generator` (Excel overview + detail sheets) · `preview` (self-contained **tabbed interactive**
RTL HTML dashboard; inline JS for behavior only) · `svg` (dependency-free inline bar/line/donut charts) ·
`run` · `publish` (gated Drive auto-publish via GWS + GitHub auto-backup) · `style_config` ·
`setup_folders` (one-time reorg) · `verify_outputs` (rebuild + correctness/structure report) ·
`tests/` (`run_tests.py` — dependency-free PASS/FAIL gate; `baseline.json` freezes the known figures).
Helpers: `screenshot.ps1` (sheet→png visual QA), `build_guide.py` (regenerates `מדריך-שימוש.docx`).
Sources of truth: `references/build-brief.md`, `references/spec-dashboard-smb.md`. Frozen visual
spec: `references/style.md` (mirrored by `scripts/style_config.py`).
