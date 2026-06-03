# -*- coding: utf-8 -*-
"""Targets MECHANISM — the deterministic forward-read resolver and the computed-values mapping.
The numbers stay frozen (metrics untouched); these lock the new plumbing: a report in quarter Q
reads quarter Q-1's doc, and Q1 of year Y steps back into the PRIOR year's Q4 (the cross-year
requirement). actual_for keeps mapping each metric to its canonical computed value, never a
re-derived one."""
import harness
from harness import check, get_ctx, EPS

from scripts import targets
from scripts import style_config as sc


def test_prev_quarter_close_crossyear(ctx=None):
    """A report in quarter Q reads quarter Q-1's doc; Q1 steps back into the prior YEAR's Q4."""
    cases = {
        "2027-01": (2026, 4),   # Jan 2027 (Q1) → prior-year Q4  ← the cross-year requirement
        "2027-02": (2026, 4),
        "2027-03": (2026, 4),
        "2027-Q1": (2026, 4),   # the quarterly report resolves the same way
        "2026-04": (2026, 1),   # April (Q2) → Q1 of the same year
        "2026-06": (2026, 1),
        "2026-07": (2026, 2),   # July (Q3) → Q2
        "2026-Q1": (2025, 4),   # Q1 report → prior-year Q4 (no doc yet → no targets)
        "2026-Q4": (2026, 3),
    }
    for period, want in cases.items():
        got = targets.prev_quarter_close(period)
        check(got == want, "prev_quarter_close(%s)=%r != %r" % (period, got, want))
    check(targets.prev_quarter_close("2026") is None, "annual must resolve to None")


def test_targets_doc_path_crossyear(ctx=None):
    """Jan-2027 resolves its targets file to outputs/2026/רבעוני/2026-Q4/2026-Q4_יעדים.docx."""
    y, q = targets.prev_quarter_close("2027-01")
    path = targets.targets_doc_path(harness.OUTPUTS, y, q)
    tail = path.replace("\\", "/").split("/")[-4:]
    want = ["2026", sc.TYPE_QUARTERLY, "2026-Q4", "2026-Q4_%s.docx" % targets.TARGETS_SUFFIX]
    check(tail == want, "cross-year targets path tail %r != %r" % (tail, want))


def test_actual_for_computed_values(ctx=None):
    """actual_for maps each metric key to (frozen label, the canonical computed value) — the
    comparison block's actuals come straight from metrics.compute (baseline), never recomputed."""
    ctx = ctx or get_ctx()
    rep = ctx.reports[harness.APR]
    base = ctx.baseline[harness.APR]
    av = targets.actual_for(rep)
    # labels come from the frozen style config
    check(av["leads"][0] == sc.M_LEADS, "leads label != M_LEADS")
    check(av["profit"][0] == sc.M_PROFIT, "profit label != M_PROFIT")
    # values are exactly the computed figures (baseline), not re-derived here
    check(av["leads"][1] == base["leads"], "leads actual %r != baseline %r" % (av["leads"][1], base["leads"]))
    check(av["deals"][1] == base["deals"], "deals actual %r != baseline %r" % (av["deals"][1], base["deals"]))
    check(av["tasks"][1] == base["tasks"], "tasks actual %r != baseline %r" % (av["tasks"][1], base["tasks"]))
    check(abs(av["net"][1] - base["net"]) <= EPS, "net actual %r != baseline %r" % (av["net"][1], base["net"]))
    check(abs(av["profit"][1] - base["profit"]) <= EPS,
          "profit actual %r != baseline %r" % (av["profit"][1], base["profit"]))
