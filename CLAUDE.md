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
C:\Python314\python.exe scripts\run.py 2026-04    # monthly (April)
C:\Python314\python.exe scripts\run.py 2026-03    # monthly + Q1 (closing month — auto roll-up)
C:\Python314\python.exe scripts\run.py 2026-Q1    # just the quarterly (explicit, on demand)
C:\Python314\python.exe scripts\run.py 2026       # just the annual (explicit, on demand)
C:\Python314\python.exe scripts\verify_outputs.py # rebuild + correctness/structure report
```
Deps: `pandas`, `openpyxl` (required); `python-docx` (only to read a *filled* quarterly Word and to
regenerate the `.docx` guide). Outputs are organized **year → type → period**, identical locally and on
Drive: `outputs/<year>/<type>/<period>/` → `<name>.xlsx` + `.html` (type ∈ `חודשי`/`רבעוני`/`שנתי`; e.g.
`outputs/2026/רבעוני/2026-Q1/2026-Q1_רבעוני.xlsx`). **xlsx + html only — no PDF anywhere.** **No loose file
under a type folder.** `verify_outputs.py` writes `verification_report.md` to the project root;
`scripts/tests/run_tests.py` is the pass/fail gate
(numbers · parity · reconciliation · expense-sheet · determinism · structure · business-logic).
Per-month input folders (`inputs/YYYY-MM/`) are the archive.

### What to generate each month (output rule)
Run **the month** — a closing month **auto-emits its roll-ups**, so there's nothing extra to remember:

| Month run (`run.py <YYYY-MM>`) | Produces |
|---|---|
| Jan · Feb · Apr · May · Jul · Aug · Oct · Nov | monthly |
| **Mar · Jun · Sep** (quarter close) | monthly **+ quarterly** |
| **Dec** (quarter + year close) | monthly **+ quarterly (Q4) + annual** |

So `run.py 2026-03` → March monthly **+ Q1**; `run.py 2026-12` → Dec monthly **+ Q4 + 2026 annual**. A
quarter/annual argument (`2026-Q1`, `2026`) builds only itself (re-generate one on demand). `run.py` also
takes several arguments at once (batch, e.g. `run.py 2026-01 2026-02`).

## Monthly run flow (`run.py`)
1. Locate inputs (`inputs/current/` + root expenses).
2. **Expense gate (human-in-the-loop):** today expenses = `הוצאות קבועות.xlsx` only ("no one-offs").
   When one-off expenses begin, the agent pauses here for confirmation before computing.
3. **Validate** (graded — below); stop on blocking.
4. **Compute** all metrics (deterministic).
5. **Targets:** read forward the *previous* quarter-closing month's targets Word from the outputs tree
   (`outputs/<year>/רבעוני/<prev-Q>/<prev-Q>_יעדים.docx`; cross-year aware — Jan–Mar read the prior year's
   Q4); a missing or still-blank doc → skip.
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
   The header carries the **studio logo in the top-LEFT corner** (the secondary side in RTL — HTML: on a
   white chip, logo-only; Excel: an image anchored at the left-edge column), with the issue date on the
   right; the KPI row sits with clear **breathing room** below the tabs (HTML: `main{padding-block}`, kept
   separate from `.wrap`'s `padding-inline` so it isn't overridden). The two income pictures are sized so
   each fills ~half the width with a clean central gutter — title/labels on one line, nothing truncated.
   The single-month income chart is a **horizontal** net/gross/value composition (labels at the bar ends,
   axis headroom) so no number collides with the title.
7. **Write** outputs into `outputs/<year>/<type>/<period>/` with fixed names (**xlsx + html only — no PDF**); the HTML is
   a self-contained, deterministic **tabbed interactive BI dashboard** (inline CSS + inline JS for behavior
   + inline SVG charts) built from the same data — **JS drives only interaction/display, never numbers** —
   so every figure matches the Excel by construction. Emit the missing-items report. On quarterly runs,
   also emit a blank **next-quarter targets Word** INTO that quarter's output folder
   (`outputs/<year>/רבעוני/<period>/<period>_יעדים.docx`) — created only if absent, so a doc the studio
   has filled is never overwritten (the folder is rebuilt every run; the targets Word is preserved).
8. **Publish & back up (both gated — never fail the run, same pattern as the old optional steps):**
   upload the period's `xlsx`+`html` to Google Drive (Shared-Drive `ניתוחים וישיבות הנהלה` →
   `דשבורדים/<year>/<type>/<period>/`, find-or-create each level) via the **GWS CLI** as `studio@smb-arch.com`;
   on a quarter close the **targets Word** is mirrored to that same Drive folder **without overwriting** an
   existing one. Then `git add -A && commit && push` the whole project to the private **GitHub** backup
   (`studio139/smb-dashboard`) — `outputs/` is git-ignored *except* the targets Word (the only `.docx` there),
   so a filled targets doc is backed up to **both** Drive and GitHub. On missing tool / no auth / offline each
   prints a notice and continues.
   *(The GWS OAuth app is now in **Production** — the token no longer expires every 7 days.)* See `publish.py`.

## Determinism
Every number originates in `scripts/metrics.py`. The agent never estimates or hand-places a value.
Consistency = generator + style frozen → each period looks identical yet adapts to the data.

## Visual / output changes — standing screenshot loop (MANDATORY)
Whenever you change anything **visual or output-facing** — the HTML dashboard, the Excel layout, charts,
the logo, or any styling — you MUST work in a **screenshot-iterate loop** before declaring the work done:
**render → take a screenshot → review it honestly and list every flaw → fix → re-render and re-screenshot →
repeat** until a fresh screenshot shows no remaining flaws and the result genuinely looks polished
(**at least 3 iterations**, with no fixed upper limit — the stopping condition is that it *genuinely looks
good*, not a number of passes). **Show the before → after screenshots.** Tooling: `scripts/screenshot.ps1`
(Excel sheet→PNG) and headless Edge/Chrome for the HTML (copy the Hebrew `.html` to an ASCII path first).

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
- Ranges: leads **from the start of 2025 (or earlier — whatever exists) = lookback** for attribution +
  mature-cohort conversion, **not** report-period leads; projects/income/partnerships/tasks from 1.26;
  report period = 2026. A wider leads history **raises attribution and may shift the settled-conversion
  rate** (more mature cohorts feed the headline) — expected and correct, **not** an error.

## Validation (graded — `scripts/validator.py`)
BLOCKING (stop + explain): whole file missing · required column missing · empty core data ·
impossible total (net>gross, conversion>100%). NON-BLOCKING (continue + record → `לא צוין`):
single missing field · unmatched join rows · single outlier. Every run ends with a missing-items report.

## Targets & timeline (`scripts/targets.py`)
The targets **table** lives in a Word that is now an **output** of the quarter-closing month, in that
quarter's folder (`outputs/<year>/רבעוני/<period>/<period>_יעדים.docx`) — one per quarter, per year. The
agent **emits it blank** at quarter close (only if absent — a filled doc is never overwritten); the studio
fills the **next** quarter's targets after the quarterly meeting (fill the *local* copy — that's the one read
forward). Each report **reads forward** the *previous* quarter-closing doc (`prev_quarter_close`): a report in
quarter Q reads Q-1's doc — **cross-year aware**, so Jan–Mar of year Y read `…/<Y-1>/רבעוני/<Y-1>-Q4/`. One
doc serves all three months of its target quarter (that single mapping is the carry-forward). Metric names
match dashboard metrics; insights/initiatives are human text — never parsed for numbers. A missing or
still-blank doc → no targets (no crash). Q1 quarterly + April monthly = no targets yet; once the Q1 doc is
filled, Q2's reports (Apr–Jun) show target-vs-actual. The Word template structure is unchanged.

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
