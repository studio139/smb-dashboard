# -*- coding: utf-8 -*-
"""metrics == xlsx cells == HTML — the same number in all three artifacts."""
import harness
from harness import (check, get_ctx, figures, load_ws, matrix_summary_value,
                     kpi_value, html_text, strip_gen, html_fig, EPS)
from scripts import style_config as sc
from scripts import viewmodel as vm

# (figure key, the overview matrix row label whose summary cell holds it)
INCOME_KEYS = [("net", sc.M_NET), ("gross", sc.M_GROSS),
               ("deals", sc.M_DEALS), ("pipeline_value", sc.M_PIPEVAL)]


def _check(ctx, p):
    rep = ctx.reports[p]
    fig = figures(rep)
    ws = load_ws(ctx.paths[p]["xlsx"], sc.T_OVERVIEW)
    n_vcols = len(vm.value_columns(rep))
    htxt = strip_gen(html_text(ctx.paths[p]["html"]))

    # xlsx income-matrix summary cells == metrics
    for figkey, label in INCOME_KEYS:
        x = matrix_summary_value(ws, label, n_vcols)
        check(x is not None, "%s xlsx: row %r not found" % (p, label))
        check(abs(x - fig[figkey]) <= EPS, "%s xlsx %s: %r != %r" % (p, figkey, x, fig[figkey]))

    # xlsx KPI cards == metrics (leads, tasks)
    check(abs(kpi_value(ws, sc.K_LEADS) - fig["leads"]) <= EPS, "%s xlsx leads card mismatch" % p)
    check(abs(kpi_value(ws, sc.K_TASKS) - fig["tasks"]) <= EPS, "%s xlsx tasks card mismatch" % p)

    # HTML == metrics (exact, via data-fig anchors)
    expect = [("net", vm.fmt_money(fig["net"])), ("deals", vm.fmt_int(fig["deals"])),
              ("leads", vm.fmt_int(fig["leads"])), ("tasks", vm.fmt_int(fig["tasks"])),
              ("profit", vm.fmt_money(fig["profit"]))]
    for figkey, s in expect:
        h = html_fig(htxt, figkey)
        check(h == s, "%s html %s: %r != %r" % (p, figkey, h, s))


def test_parity_q1(ctx=None):
    _check(ctx or get_ctx(), harness.Q1)


def test_parity_april(ctx=None):
    _check(ctx or get_ctx(), harness.APR)
