# -*- coding: utf-8 -*-
"""Monthly columns sum to the summary column — income A/B, expenses/profit,
attribution, source, performer. Ratio rows (conversion / pct) are exempt."""
import harness
from harness import check, get_ctx, load_ws, find_label_cell, EPS
from scripts import style_config as sc
from scripts import viewmodel as vm


def _blocks(rep, data):
    a, b = vm.income_blocks(rep)
    return [("income_a", a), ("income_b", b),
            ("expenses", vm.expenses_block(rep)),
            ("attribution", vm.attribution_block(rep, data)),
            ("source", vm.source_matrix(rep)),
            ("performer", vm.performer_matrix(rep))]


def _check_viewmodel(ctx, p):
    rep = ctx.reports[p]
    months = rep["months"]
    for name, block in _blocks(rep, ctx.data):
        for row in block.rows:
            if row.fmt == "pct" or row.by_month is None:   # ratios / summary-only exempt
                continue
            s = sum((row.by_month.get(m) or 0) for m in months)
            check(abs(s - row.summary) <= EPS,
                  "%s %s/%s: months=%r != summary %r" % (p, name, row.label, s, row.summary))


def test_reconciliation_q1(ctx=None):
    _check_viewmodel(ctx or get_ctx(), harness.Q1)


def test_reconciliation_april(ctx=None):
    # n==1: the single month column IS the summary (no double counting)
    _check_viewmodel(ctx or get_ctx(), harness.APR)


def test_reconciliation_xlsx_income(ctx=None):
    """Independent pass over the rendered xlsx income blocks (Q1, multi-month)."""
    ctx = ctx or get_ctx()
    rep = ctx.reports[harness.Q1]
    n_vcols = len(vm.value_columns(rep))        # 3 months + sum
    ws = load_ws(ctx.paths[harness.Q1]["xlsx"], sc.T_OVERVIEW)
    for label in (sc.M_DEALS, sc.M_PIPEVAL, sc.M_NET, sc.M_GROSS, sc.M_PAYMENTS):
        r, c = find_label_cell(ws, label)
        month_sum = sum(ws.cell(r, c + 1 + j).value or 0 for j in range(n_vcols - 1))
        summary = ws.cell(r, c + n_vcols).value
        check(abs(month_sum - summary) <= EPS,
              "xlsx %s: months=%r != summary %r" % (label, month_sum, summary))
