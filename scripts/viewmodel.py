# -*- coding: utf-8 -*-
"""Presentation view-model — the SINGLE producer of every displayed number.

`metrics.py` computes the canonical report; this module reshapes that report (plus
the loaded `data` frames) into plain-Python *blocks* that BOTH renderers consume
identically:
  * `generator.py` (Excel)  -> `_matrix_block` / `_detail_sheet`
  * `preview.py`   (HTML)   -> `block_table` / detail `<details>` tables

Because both renderers iterate the same objects and choose only a *format kind* and a
*semantic color token*, the numbers in the xlsx and the html are identical BY
CONSTRUCTION — neither side computes a value of its own.

NOTHING here invents a number: every summary cell comes straight from the canonical
report; the only derived series (attribution per month) is bucketed from the report's
own inputs using the FROZEN `metrics` helpers and guarded by a reconciliation check
against the canonical totals. Sorting is centralized (compound key) so the order is
deterministic and identical across both renderers.
"""
from collections import namedtuple

import pandas as pd

from scripts import style_config as sc
from scripts import metrics as M

# A single data row of a monthly+summary matrix.
#   by_month: {"YYYY-MM": number|None} (or None when the row has no monthly split)
#   summary:  number|None  (printed verbatim in the "סיכום" column — never recomputed)
#   fmt:      "money" | "int" | "pct" | "days"
#   color:    semantic token "good"|"bad"|"warn"|"primary"|None  (NOT a hex)
Row = namedtuple("Row", "label by_month summary fmt color")

# A block = a colored title + an ordered list of Rows.
#   color: semantic header token "pic_a"|"pic_b"|"primary"
Block = namedtuple("Block", "key title color rows")


# --------------------------------------------------------------- canonical formatters
# The ONE source of display strings for the HTML layer (and the test harness). The
# Excel layer uses style_config number formats for cells; these mirror them exactly so
# a value formatted here matches what Excel shows (zero -> em dash).
def fmt_money(v):
    return "—" if v is None or round(v) == 0 else "₪ {:,.0f}".format(v)


def fmt_int(v):
    return "—" if v is None or round(v) == 0 else "{:,.0f}".format(v)


def fmt_pct(v):
    return "—" if v is None else "{:.1f}%".format(v * 100)


def fmt_days(v):
    return "—" if v is None else "{:,.0f} ימים".format(v)


FMT = {"money": fmt_money, "int": fmt_int, "pct": fmt_pct, "days": fmt_days}


def _sorted_keys(mapping):
    """Value-ranked, name tiebroken — deterministic and identical in both renderers."""
    return sorted(mapping, key=lambda k: (-float(mapping[k] or 0), str(k)))


def _txt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    return str(v)


# --------------------------------------------------------------- columns
def value_columns(report):
    """The month/summary column rule — the single home (was duplicated in both layers).

    n==1  -> one month column, no summary
    n<=6  -> one column per month + a "סיכום" summary
    n==12 -> summary only (months live in the charts; keeps width bounded)
    """
    months = report["months"]
    n = len(months)
    if n == 1:
        return [("month", months[0])]
    if n <= 6:
        return [("month", m) for m in months] + [("sum", None)]
    return [("sum", None)]


# --------------------------------------------------------------- income pictures
def income_blocks(report):
    """The two income pictures (never merged): A = pipeline, B = cash."""
    P, C = report["pipeline"], report["cash"]
    a = Block("income_a", sc.T_PIC_A, "pic_a", [
        Row(sc.M_DEALS, P["deals_by_month"], P["deals_total"], "int", None),
        Row(sc.M_PIPEVAL, P["value_by_month"], P["value_total"], "money", None),
    ])
    b = Block("income_b", sc.T_PIC_B, "pic_b", [
        Row(sc.M_NET, C["net_by_month"], C["net_total"], "money", None),
        Row(sc.M_GROSS, C["gross_by_month"], C["gross_total"], "money", None),
        Row(sc.M_PAYMENTS, C["payments_by_month"], C["payments_total"], "int", None),
    ])
    return a, b


# --------------------------------------------------------------- expenses & profit
def profit_row(report):
    """Profit is already per-month in the report; colored by sign (good/bad)."""
    Pr = report["profit"]
    color = "good" if Pr["total"] >= 0 else "bad"
    return Row(sc.M_PROFIT, Pr["by_month"], Pr["total"], "money", color)


def expenses_block(report):
    """Expenses (monthly-replication model) + profit, as a monthly+summary matrix.

    Fixed expenses repeat each month -> every month column shows the same value and the
    summary = value * n. Profit uses the real per-month profit. Both reconcile: the
    monthly columns sum to the summary column.
    """
    E = report["expenses"]
    months = report["months"]
    n = len(months)

    def repl(value):
        return {m: value for m in months}

    rows = [
        Row(sc.M_EXP, repl(E["total"]), E["total"] * n, "money", None),
        Row("  קבועות", repl(E["fixed"]), E["fixed"] * n, "money", None),
        Row("  משתנות", repl(E["variable"]), E["variable"] * n, "money", None),
    ]
    for k in _sorted_keys(E["by_category"]):
        v = E["by_category"][k]
        rows.append(Row("    " + str(k), repl(v), v * n, "money", None))
    rows.append(profit_row(report))
    return Block("expenses", sc.T_EXP, "primary", rows)


# --------------------------------------------------------------- attribution (derived)
def attribution_block(report, data):
    """Deals-to-leads attribution, per month + summary.

    The report stores only the period totals, so the per-month split is bucketed from
    the same inputs metrics used (projects started in-period, matched by client name
    against lead names) via the FROZEN helpers `metrics.mk`/`metrics.norm_name`. The
    monthly columns therefore sum to the canonical totals BY CONSTRUCTION — asserted
    below so any future drift in metrics fails loudly instead of silently desyncing.
    """
    A = report["attribution"]
    months = report["months"]
    proj = (data or {}).get("projects")
    leads = (data or {}).get("leads")

    if proj is None or not len(proj):
        # No per-month split available — show canonical summary only.
        rows = [
            Row("עסקאות שמקורן בליד מתויג", None, A["matched"], "int", None),
            Row("סך עסקאות שנסגרו", None, A["total"], "int", None),
            Row("שיעור שמקורו בליד", None, A["pct"], "pct", None),
        ]
        return Block("attribution", sc.T_ATTR, "primary", rows)

    lead_names = set()
    if leads is not None and len(leads):
        lead_names = {M.norm_name(nm) for nm in leads["name"].dropna()}

    matched = {m: 0 for m in months}
    total = {m: 0 for m in months}
    for _, r in proj.iterrows():
        m = M.mk(r.get("start"))
        if m not in total:                       # start month outside the period
            continue
        total[m] += 1
        nm = M.norm_name(r.get("client"))
        if nm and nm in lead_names:
            matched[m] += 1

    # Reconciliation guard — the derived split must equal the canonical totals.
    if sum(total.values()) != A["total"] or sum(matched.values()) != A["matched"]:
        raise ValueError(
            "attribution per-month derivation diverged from metrics: "
            "matched {}/{} total {}/{}".format(
                sum(matched.values()), A["matched"], sum(total.values()), A["total"]))

    pct = {m: (matched[m] / total[m] if total[m] else None) for m in months}
    rows = [
        Row("עסקאות שמקורן בליד מתויג", matched, A["matched"], "int", None),
        Row("סך עסקאות שנסגרו", total, A["total"], "int", None),
        Row("שיעור שמקורו בליד", pct, A["pct"], "pct", None),   # ratio: summary != sum
    ]
    return Block("attribution", sc.T_ATTR, "primary", rows)


# --------------------------------------------------------------- open-schema breakdowns
def source_matrix(report):
    """Leads by source, monthly + summary (+ a total row)."""
    L = report["leads"]
    months = report["months"]
    by_source = L["by_source"]
    totals = {s: sum(by_source[s].get(m, 0) for m in months) for s in by_source}
    rows = [Row(str(s), by_source[s], totals[s], "int", None) for s in _sorted_keys(totals)]
    rows.append(Row(sc.S_TOTAL, L["by_month"], L["total"], "int", "primary"))
    return Block("source", sc.T_LEADS, "primary", rows)


def performer_matrix(report):
    """Tasks by performer (מבצע), monthly + summary (+ a total row)."""
    T = report["tasks"]
    months = report["months"]
    bpm = T["by_performer_by_month"]
    totals = {p: sum(bpm[p].get(m, 0) for m in months) for p in bpm}
    rows = [Row(str(p), bpm[p], totals[p], "int", None) for p in _sorted_keys(totals)]
    rows.append(Row(sc.S_TOTAL, T["by_month"], T["total"], "int", "primary"))
    return Block("performer", sc.T_TASKS, "primary", rows)


def analytics_blocks(report, data):
    """The monthly+summary analytics matrices, in display order."""
    return [
        expenses_block(report),
        attribution_block(report, data),
        source_matrix(report),
        performer_matrix(report),
    ]


# --------------------------------------------------------------- detail tables
# Shared row-level detail-sheet structure (Excel sheets AND HTML collapsible tables).
# cols = (key, header, kind, width); width is honored by Excel, ignored by HTML.
DETAIL_SPECS = [
    {"name": sc.DT_PROJECTS, "datecol": "start", "done_only": False, "cols": [
        ("name", "שם הפרויקט", "text", 26), ("client", "לקוח", "text", 20),
        ("city", "עיר", "text", 12), ("ptype", "סוג", "text", 14),
        ("status", "סטטוס", "text", 18), ("start", "תאריך התחלה", "date", 14),
        ("fee", "שכ״ט (₪)", "money", 14), ("fee_stopped", "שכ״ט אם הופסק (₪)", "money", 16),
    ]},
    {"name": sc.DT_LEADS, "datecol": "arrival", "done_only": False, "cols": [
        ("name", "שם", "text", 26), ("arrival", "תאריך הגעה", "date", 14),
        ("source", "מקור הגעה", "text", 22), ("client_type", "סוג לקוח", "text", 18),
    ]},
    {"name": sc.DT_INCOME, "datecol": "date", "done_only": False, "cols": [
        ("date", "תאריך", "date", 14), ("project", "פרויקט", "text", 26),
        ("partner", "שת״פ", "text", 18), ("gross", "סכום ברוטו (₪)", "money", 16),
        ("invoice", "חשבונית", "text", 12),
    ]},
    {"name": sc.DT_PART, "datecol": "executed", "done_only": False, "cols": [
        ("name", "שם", "text", 26), ("status", "סטטוס", "text", 22),
        ("executed", "תאריך שבוצע", "date", 14), ("fee", "שכ״ט (₪)", "money", 14),
        ("fee_cancelled", "שכ״ט אם בוטל (₪)", "money", 16),
    ]},
    {"name": sc.DT_TASKS, "datecol": "updated", "done_only": True, "cols": [
        ("performer", "מבצע", "text", 22), ("status", "סטטוס", "text", 16),
        ("updated", "עודכן בתאריך", "date", 14),
    ]},
]

# The new expenses detail sheet (monthly-replication: each expense row × each month).
EXPENSE_SPEC = {"name": sc.DT_EXPENSES, "cols": [
    ("month_lbl", "חודש", "text", 16), ("category", "קטגוריה", "text", 22),
    ("desc", "תיאור", "text", 30), ("etype", "סוג", "text", 14),
    ("amount", "סכום (₪)", "money", 16),
]}


def filter_detail_rows(df, months, datecol, done_only=False):
    """Period-scoped rows for a detail dimension — the single filter both layers use."""
    out = []
    if df is None or not len(df):
        return out
    for _, row in df.iterrows():
        if M.mk(row.get(datecol)) not in months:
            continue
        if done_only and not (row.get("status") and M._DONE in M.nfc(row["status"])):
            continue
        out.append(row)
    return out


def detail_tables(data, months):
    """For each of the five dimensions: its spec + the period-filtered rows.

    Returns [{"name", "cols", "rows"}] where rows are pandas Series (or dicts) the
    renderers paint by key. Identical filtering -> Excel sheets and HTML tables match.
    """
    out = []
    for spec in DETAIL_SPECS:
        df = (data or {}).get(_DETAIL_DATA_KEY[spec["name"]])
        rows = filter_detail_rows(df, months, spec["datecol"], spec["done_only"])
        out.append({"name": spec["name"], "cols": spec["cols"], "rows": rows})
    return out


def expense_detail_rows(data, months):
    """Monthly-replication expansion of the fixed-expenses file: each row × each month.

    Grouped by month (the period-scoping column). Invariant: the sum of `amount` over
    the returned rows == report["expenses"]["total"] * len(months) == the overview
    expenses headline.
    """
    df = (data or {}).get("expenses")
    rows = []
    if df is None or not len(df):
        return rows
    for m in months:
        for _, r in df.iterrows():
            amt = r.get("amount")
            amt = None if (amt is None or (isinstance(amt, float) and pd.isna(amt))) else float(amt)
            rows.append({
                "month": m, "month_lbl": sc.heb_month(m),
                "category": _txt(r.get("category")), "desc": _txt(r.get("desc")),
                "etype": _txt(r.get("etype")), "amount": amt,
            })
    return rows


# Maps a detail sheet's display name -> the loader data key.
_DETAIL_DATA_KEY = {
    sc.DT_PROJECTS: "projects", sc.DT_LEADS: "leads", sc.DT_INCOME: "income",
    sc.DT_PART: "partnerships", sc.DT_TASKS: "tasks",
}


# --------------------------------------------------------------- chart series
def chart_series(report):
    """Every number the charts plot — ONE source for the Excel data sheet and the HTML
    inline SVG, with breakdowns pre-sorted (same order in both)."""
    months = report["months"]
    out = {
        "months": months,
        "month_labels": [sc.heb_month(m) for m in months],
        "net": [report["cash"]["net_by_month"][m] for m in months],
        "gross": [report["cash"]["gross_by_month"][m] for m in months],
        "value": [report["pipeline"]["value_by_month"][m] for m in months],
        "leads": [report["leads"]["by_month"][m] for m in months],
        "deals": [report["pipeline"]["deals_by_month"][m] for m in months],
        "tasks": [report["tasks"]["by_month"][m] for m in months],
        "profit": [report["profit"]["by_month"][m] for m in months],
    }
    L = report["leads"]["by_source"]
    src_tot = {k: sum(L[k].get(m, 0) for m in months) for k in L}
    out["src"] = [(str(k), src_tot[k]) for k in _sorted_keys(src_tot)]
    T = report["tasks"]["by_performer_by_month"]
    perf_tot = {k: sum(T[k].get(m, 0) for m in months) for k in T}
    out["perf"] = [(str(k), perf_tot[k]) for k in _sorted_keys(perf_tot)]
    cat = report["expenses"]["by_category"]
    out["cat"] = [(str(k), cat[k]) for k in _sorted_keys(cat)]
    return out
