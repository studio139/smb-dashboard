# -*- coding: utf-8 -*-
"""Key figures unchanged — metrics.compute matches the frozen baseline exactly."""
import harness
from harness import check, figures, get_ctx, EPS

INT_KEYS = {"leads", "deals", "payments", "tasks", "attr_matched", "attr_total"}
FLOAT_KEYS = {"pipeline_value", "net", "gross", "expenses_base", "expenses_period",
              "profit", "conversion"}
SUM_KEYS = ("leads", "deals", "payments", "tasks", "net", "gross",
            "pipeline_value", "profit", "attr_matched", "attr_total")


def _assert_period(ctx, p):
    got = figures(ctx.reports[p])
    base = ctx.baseline[p]
    for k in INT_KEYS:
        check(got[k] == base[k], "%s.%s: %r != baseline %r" % (p, k, got[k], base[k]))
    for k in FLOAT_KEYS:
        a, b = got[k], base[k]
        if a is None or b is None:
            check(a == b, "%s.%s: None mismatch %r/%r" % (p, k, a, b))
        else:
            check(abs(a - b) <= EPS, "%s.%s: %r != baseline %r" % (p, k, a, b))


def test_numbers_q1(ctx=None):
    _assert_period(ctx or get_ctx(), harness.Q1)


def test_numbers_april(ctx=None):
    _assert_period(ctx or get_ctx(), harness.APR)


def test_numbers_monthly_sum_quarter(ctx=None):
    """Jan+Feb+Mar == Q1 for every additive figure (catches a corrupted baseline too)."""
    ctx = ctx or get_ctx()
    q1 = figures(ctx.reports[harness.Q1])
    months = [figures(ctx.reports[m]) for m in harness.MONTHS3]
    for k in SUM_KEYS:
        s = sum(m[k] for m in months)
        check(abs(s - q1[k]) <= EPS, "sum(%s) months=%r != Q1 %r" % (k, s, q1[k]))
