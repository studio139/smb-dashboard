# -*- coding: utf-8 -*-
"""Entry point — the monthly/quarterly/annual run.

  C:\\Python314\\python.exe scripts\\run.py 2026-04        # April monthly
  C:\\Python314\\python.exe scripts\\run.py 2026-03        # March monthly + Q1 (auto roll-up)
  C:\\Python314\\python.exe scripts\\run.py 2026-12        # Dec monthly + Q4 + annual (auto)
  C:\\Python314\\python.exe scripts\\run.py 2026-Q1        # just the quarterly (explicit)
  C:\\Python314\\python.exe scripts\\run.py 2026-01 2026-02 # batch (each arg processed in turn)

A MONTH argument auto-emits its roll-ups (see expand_periods): a quarter-closing month
(Mar/Jun/Sep) also produces that quarter; December also produces the annual. Outputs are
organized outputs/<year>/<type>/<period>/ (type = חודשי/רבעוני/שנתי), mirrored on Drive.

Flow: locate the month's input folder -> expense gate (dormant) -> validate -> for each
period: compute -> targets (if a filled Word exists) -> render xlsx + html -> publish to
Drive; then back up the project to GitHub once. Per-month folders (inputs/YYYY-MM/) ARE
the archive — nothing is copied out. Drive publish + GitHub backup are gated (never block).
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from scripts import loader, metrics, validator, generator, preview, targets, publish  # noqa: E402
from scripts import style_config as sc  # noqa: E402

OUT = os.path.join(ROOT, "outputs")
INPUTS = os.path.join(ROOT, "inputs")
TEMPLATES = os.path.join(ROOT, "templates")

# Reporting window starts 1.26 (projects/income/partnerships/tasks); 9–12.25 leads are
# lookback only. A KPI Δ vs "previous period" is shown only when that prior period falls
# inside the reporting window — so 2026-Q1 (prev = Q4-2025 lookback) shows no Δ.
REPORT_START = "2026-01"


def out_name(zoom, year, quarter, months):
    if zoom == "quarterly":
        return f"{year}-Q{quarter}_רבעוני"
    if zoom == "annual":
        return f"{year}_שנתי"
    return f"{months[0]}_חודשי"


# outputs/ (and Drive דשבורדים/) are organized year → type → period.
TYPE_FOLDER = {"monthly": sc.TYPE_MONTHLY, "quarterly": sc.TYPE_QUARTERLY, "annual": sc.TYPE_ANNUAL}


def period_leaf(zoom, year, quarter, months):
    """The period's own folder name (mirrors the run argument)."""
    if zoom == "quarterly":
        return f"{year}-Q{quarter}"
    if zoom == "annual":
        return f"{year}"
    return months[0]


def period_segments(zoom, year, quarter, months):
    """Path segments under outputs/ and under דשבורדים/ on Drive: [year, type, period]."""
    return [str(year), TYPE_FOLDER[zoom], period_leaf(zoom, year, quarter, months)]


def output_target(rep):
    """(period_dir, xlsx_path, html_path, segments) for a computed report — the single
    source of truth for where a report's files live (reused by verify_outputs.py)."""
    segs = period_segments(rep["zoom"], rep["year"], rep["quarter"], rep["months"])
    name = out_name(rep["zoom"], rep["year"], rep["quarter"], rep["months"])
    period_dir = os.path.join(OUT, *segs)
    return period_dir, os.path.join(period_dir, name + ".xlsx"), os.path.join(period_dir, name + ".html"), segs


def expand_periods(arg):
    """A closing MONTH auto-emits its roll-ups: month 3/6/9/12 → + that quarter; month 12 →
    + the annual. Quarter / annual args produce only themselves."""
    zoom, months, year, quarter = metrics.parse_period(arg)
    if zoom != "monthly":
        return [arg]
    out = [arg]
    m = int(months[0].split("-")[1])
    if m % 3 == 0:
        out.append(f"{year}-Q{m // 3}")
    if m == 12:
        out.append(f"{year}")
    return out


def _prev_period(zoom, year, quarter, months):
    """(previous-period string, its last month) for the KPI Δ comparison."""
    if zoom == "quarterly":
        q, y = quarter - 1, year
        if q == 0:
            q, y = 4, year - 1
        return f"{y}-Q{q}", f"{y:04d}-{q * 3:02d}"
    if zoom == "annual":
        return f"{year - 1}", f"{year - 1:04d}-12"
    pm = metrics.add_months(months[0], -1)
    return pm, pm


def _prev_kpis(data, rep):
    """Previous-period card totals (from the canonical compute) — or None when the
    prior period precedes the reporting window."""
    prev_str, prev_last = _prev_period(rep["zoom"], rep["year"], rep["quarter"], rep["months"])
    if prev_last < REPORT_START:
        return None
    pr = metrics.compute(data, prev_str)
    return {
        "net": pr["cash"]["net_total"],
        "profit": pr["profit"]["total"],
        "deals": pr["pipeline"]["deals_total"],
        "pipe_value": pr["pipeline"]["value_total"],
        "leads": pr["leads"]["total"],
        "conversion": pr["cohort"]["headline_rate"],
        "tasks": pr["tasks"]["total"],
    }


def _build(period, data, base_missing, log):
    """Compute, validate, render, and Drive-publish ONE report. Returns 0, or 2 on a
    blocking-totals error. Inputs are passed in (loaded once by main)."""
    rep = metrics.compute(data, period)
    rep["missing"] = base_missing + rep["missing"]

    block2 = validator.validate_report(rep)
    if block2:
        print(f"BLOCKING (totals) — {period} stopped:")
        for b in block2:
            print("  ! " + b)
        return 2

    # targets: read forward the PREVIOUS quarter-close doc from the outputs tree (cross-year
    # aware: Q1 of year Y reads Y-1 Q4). None when that doc is missing or still the blank template.
    rep["targets"] = targets.find_targets(OUT, period)

    # KPI Δ vs previous period (None when no prior reporting period exists)
    rep["prev"] = _prev_kpis(data, rep)

    period_dir, xlsx, htmlp, segs = output_target(rep)
    os.makedirs(period_dir, exist_ok=True)
    generator.build_workbook(rep, xlsx, data)
    preview.build_html(rep, data, htmlp)
    log.append("wrote " + os.path.relpath(xlsx, ROOT))
    log.append("wrote " + os.path.relpath(htmlp, ROOT))
    # auto-publish this period's files to Drive (gated — never blocks the run)
    log.append(publish.to_drive([xlsx, htmlp], segs))

    # On a quarter close, emit the NEXT quarter's blank targets Word INTO this quarterly output
    # folder (only if absent — a filled doc is preserved across regeneration), and mirror it to
    # the same Drive folder without overwriting an existing one. Both gated; never block the run.
    if rep["zoom"] == "quarterly":
        emitted = targets.ensure_targets_doc(TEMPLATES, period_dir, rep["year"], rep["quarter"])
        tpath = os.path.join(period_dir, targets.targets_doc_name(rep["year"], rep["quarter"]))
        if os.path.isfile(tpath):
            log.append("targets Word: " + os.path.relpath(tpath, ROOT) +
                       (" (emitted blank)" if emitted else " (kept existing)"))
            log.append(publish.to_drive_preserve([tpath], segs))

    # ascii-safe key numbers (Hebrew-free) so the console is readable everywhere
    print(f"[{period}] leads={rep['leads']['total']} deals={rep['pipeline']['deals_total']} "
          f"pipeline={rep['pipeline']['value_total']:.0f} "
          f"net={rep['cash']['net_total']:.0f} gross={rep['cash']['gross_total']:.0f} "
          f"profit={rep['profit']['total']:.0f} tasks={rep['tasks']['total']} "
          f"conv={'' if rep['cohort']['headline_rate'] is None else round(rep['cohort']['headline_rate']*100,1)} "
          f"attrib={rep['attribution']['matched']}/{rep['attribution']['total']} "
          f"missing={len(rep['missing'])}")
    return 0


def main(arg):
    """Produce the report(s) for `arg`; a closing month auto-emits its roll-ups (see
    expand_periods). Inputs are loaded once; the project is backed up once at the end."""
    log = []
    os.makedirs(OUT, exist_ok=True)
    periods = expand_periods(arg)
    extra = ("  ->  " + ", ".join(periods)) if len(periods) > 1 else ""
    print(f"== run {arg} =={extra}")

    input_dir = loader.latest_input_dir()
    if input_dir is None:
        print("BLOCKING — no inputs/YYYY-MM snapshot folder found")
        return 2
    log.append("inputs: " + os.path.relpath(input_dir, ROOT))
    data, report_load = loader.load_all(input_dir)
    # human-in-the-loop expense gate (dormant — fixed expenses only, no one-offs yet)
    print("expense gate: one-off expenses = none; expenses = הוצאות קבועות.xlsx")

    blocking, missing = validator.validate_inputs(data, report_load)
    if blocking:
        print("BLOCKING — run stopped:")
        for b in blocking:
            print("  ! " + b)
        return 2

    rc = 0
    for period in periods:
        rc = max(rc, _build(period, data, missing, log))

    # auto-backup the whole project to GitHub once, after all reports (gated — never blocks)
    log.append(publish.to_github(ROOT, arg))
    for line in log:
        print("  " + line)
    print("OK" if rc == 0 else "DONE (with blocking totals)")
    return rc


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: run.py <period> [<period> ...]   e.g. 2026-04 | 2026-03 (->+Q1) | 2026-Q1 | 2026")
        sys.exit(1)
    code = 0
    for a in sys.argv[1:]:
        code = max(code, main(a))
    sys.exit(code)
