# -*- coding: utf-8 -*-
"""No-regression on business logic (mirrors verify_outputs.py checks) — exercising the
real metrics code paths, never a re-implementation."""
import pandas as pd

import harness
from harness import check, get_ctx
from scripts import metrics


def test_bl_vat_rule(ctx=None):
    """net = gross/1.18 iff invoice=כן, else gross. (== verify check 4)"""
    ctx = ctx or get_ctx()
    for _, r in ctx.data["income"].iterrows():
        m = metrics.mk(r["date"])
        if m is None:
            continue
        g = float(r.get("gross") or 0)
        issued = bool(r.get("invoice")) and metrics.nfc(r["invoice"]).strip() == metrics._YES
        net = metrics._cash({"income": pd.DataFrame([r])}, [m])["net_total"]
        want = g / metrics.VAT if issued else g
        check(abs(net - want) <= 1e-6, "VAT: net %r != %r (issued=%s)" % (net, want, issued))


def test_bl_project_fee_exclusive(ctx=None):
    """הופסק -> שכ\"ט-אם-הופסק ONLY (replaces full fee). (== verify check 5)"""
    ctx = ctx or get_ctx()
    for _, r in ctx.data["projects"].iterrows():
        if r.get("status") and metrics._STOPPED in metrics.nfc(r["status"]):
            m = metrics.mk(r["start"])
            if m is None:
                continue
            out = metrics._pipeline({"projects": pd.DataFrame([r]), "partnerships": None}, [m], [])
            check(abs(out["value_total"] - float(r.get("fee_stopped") or 0)) <= 1e-6,
                  "stopped project fee not exclusive")


def test_bl_partnership_cancel_exclusive(ctx=None):
    """בוטל -> שכ\"ט-אם-בוטל ONLY (recorded, replaces full fee). (== verify check 5)"""
    ctx = ctx or get_ctx()
    for _, r in ctx.data["partnerships"].iterrows():
        st = metrics.nfc(r["status"]) if r.get("status") else ""
        if metrics._P_CANCEL in st:
            m = metrics.mk(r["executed"])
            if m is None:
                continue
            out = metrics._pipeline({"projects": None, "partnerships": pd.DataFrame([r])}, [m], [])
            check(abs(out["value_total"] - float(r.get("fee_cancelled") or 0)) <= 1e-6,
                  "cancelled partnership fee not exclusive")


def test_bl_tasks_done_only(ctx=None):
    """Tasks counted = הושלם per performer. (== verify check 6)"""
    ctx = ctx or get_ctx()
    T = ctx.reports[harness.Q1]["tasks"]
    check(T["total"] == 551 and len(T["performers"]) > 0, "tasks done-only / performers")


def test_bl_cohort_anchored(ctx=None):
    """Conversion cohort anchored to arrival; recent months flagged maturing. (== check 3)"""
    ctx = ctx or get_ctx()
    rows = ctx.reports[harness.Q1]["cohort"]["rows"]
    check(rows and any(r["maturing"] for r in rows), "cohort not anchored / none maturing")


def test_bl_attribution_not_missing(ctx=None):
    """Attribution non-matches are NOT recorded as data-quality misses. (== check 8)"""
    ctx = ctx or get_ctx()
    miss = " | ".join(ctx.reports[harness.Q1]["missing"])
    check("ללא ליד" not in miss and "תואם" not in miss, "attribution leaked into missing report")
