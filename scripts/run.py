# -*- coding: utf-8 -*-
"""Entry point — the monthly/quarterly/annual run.

  C:\\Python314\\python.exe scripts\\run.py 2026-Q1
  C:\\Python314\\python.exe scripts\\run.py 2026-04
  C:\\Python314\\python.exe scripts\\run.py 2026

Flow: locate the month's input folder -> expense gate (dormant) -> validate ->
compute -> targets (if a filled Word exists) -> render xlsx + html -> write output.
Per-month folders (inputs/YYYY-MM/) ARE the archive — nothing is copied out.
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from scripts import loader, metrics, validator, generator, preview, targets, publish  # noqa: E402

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


def period_folder(zoom, year, quarter, months):
    """Subfolder under outputs/ for this period (mirrors the run argument)."""
    if zoom == "quarterly":
        return f"{year}-Q{quarter}"
    if zoom == "annual":
        return f"{year}"
    return months[0]


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


def main(period):
    log = []
    os.makedirs(OUT, exist_ok=True)
    print(f"== run {period} ==")

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
        # also surface as ascii codes for safety
        return 2

    rep = metrics.compute(data, period)
    rep["missing"] = missing + rep["missing"]

    block2 = validator.validate_report(rep)
    if block2:
        print("BLOCKING (totals) — run stopped:")
        for b in block2:
            print("  ! " + b)
        return 2

    # targets: carry-forward the latest FILLED meeting Word (none yet on first build)
    rep["targets"] = targets.find_targets(INPUTS, period)
    if rep["zoom"] == "quarterly":
        emitted = targets.ensure_blank_next_quarter(TEMPLATES, INPUTS, rep["year"], rep["quarter"])
        if emitted:
            log.append("emitted blank meeting template: " + emitted)

    # KPI Δ vs previous period (None when no prior reporting period exists)
    rep["prev"] = _prev_kpis(data, rep)

    name = out_name(rep["zoom"], rep["year"], rep["quarter"], rep["months"])
    folder = period_folder(rep["zoom"], rep["year"], rep["quarter"], rep["months"])
    period_dir = os.path.join(OUT, folder)
    os.makedirs(period_dir, exist_ok=True)
    xlsx = os.path.join(period_dir, name + ".xlsx")
    htmlp = os.path.join(period_dir, name + ".html")
    generator.build_workbook(rep, xlsx, data)
    preview.build_html(rep, data, htmlp)
    log.append("wrote " + os.path.relpath(xlsx, ROOT))
    log.append("wrote " + os.path.relpath(htmlp, ROOT))

    # auto-publish to Drive + auto-backup the project to GitHub.
    # Both are gated (mirroring the optional-feature pattern): on any problem they record a
    # notice and the run continues — the report is never blocked by network/auth issues.
    log.append(publish.to_drive([xlsx, htmlp], folder))
    log.append(publish.to_github(ROOT, folder))

    # ascii-safe key numbers (Hebrew-free) so the console is readable everywhere
    print(f"leads={rep['leads']['total']} deals={rep['pipeline']['deals_total']} "
          f"pipeline={rep['pipeline']['value_total']:.0f} "
          f"net={rep['cash']['net_total']:.0f} gross={rep['cash']['gross_total']:.0f} "
          f"profit={rep['profit']['total']:.0f} tasks={rep['tasks']['total']} "
          f"conv={'' if rep['cohort']['headline_rate'] is None else round(rep['cohort']['headline_rate']*100,1)} "
          f"attrib={rep['attribution']['matched']}/{rep['attribution']['total']}")
    for line in log:
        print("  " + line)
    print(f"missing-items: {len(rep['missing'])}")
    print("OK")
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: run.py <period>  e.g. 2026-Q1 | 2026-04 | 2026")
        sys.exit(1)
    sys.exit(main(sys.argv[1]))
