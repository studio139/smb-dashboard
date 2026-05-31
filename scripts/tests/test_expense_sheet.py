# -*- coding: utf-8 -*-
"""The הוצאות detail sheet reconciles to the expenses headline:
sum of amount cells in the period == expenses.total * n_months."""
import harness
from harness import check, get_ctx, load_ws, find_header_col, _nfc
from scripts import style_config as sc


def _check(ctx, p, n):
    rep = ctx.reports[p]
    ws = load_ws(ctx.paths[p]["xlsx"], sc.DT_EXPENSES)
    hdr = 3
    amt_col = find_header_col(ws, hdr, "סכום")
    check(amt_col is not None, "%s: expenses 'סכום' header not found" % p)
    nodata = _nfc(sc.S_NODATA)
    total, rows = 0.0, 0
    for row in ws.iter_rows(min_row=hdr + 1):
        first = row[0].value
        if first is None or _nfc(first) == nodata:
            continue
        rows += 1
        cell = row[amt_col - 1]
        if isinstance(cell.value, (int, float)):
            total += cell.value
    want = rep["expenses"]["total"] * n
    check(abs(total - want) <= 1e-4, "%s expenses sum %r != total*n %r" % (p, total, want))
    n_lines = len(ctx.data["expenses"])
    check(rows == n_lines * n, "%s expenses rows %d != %d (lines*n)" % (p, rows, n_lines * n))


def test_expense_sheet_q1(ctx=None):
    _check(ctx or get_ctx(), harness.Q1, 3)


def test_expense_sheet_april(ctx=None):
    _check(ctx or get_ctx(), harness.APR, 1)
