# -*- coding: utf-8 -*-
"""Deterministic metric computation. Every number in the dashboard originates here.

Two income pictures are kept strictly separate:
  Picture A (pipeline / "money closed") = project fees + partnership fees
  Picture B (cash / "money in")         = receipts; net = gross/1.18 when חשבונית=כן
"""
import re
import unicodedata

import pandas as pd

VAT = 1.18
MATURING_MONTHS = 3  # most-recent arrival months flagged "still maturing"

# status / value tokens (substring matches — open-schema tolerant)
_STOPPED = "הופסק"               # הופסק (project -> reduced fee)
_P_CLOSED = ("בוצע", "נסגר")  # בוצע / נסגר (partnership closed)
_P_CANCEL = "בוטל"                    # בוטל (partnership cancelled -> excluded)
_DONE = "הושלם"                  # הושלם (task completed)
_YES = "כן"                                     # כן (invoice issued)


def nfc(s):
    return unicodedata.normalize("NFC", str(s))


def norm_name(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    return re.sub(r"\s+", " ", nfc(s)).strip()


def mk(ts):
    return None if pd.isna(ts) else f"{ts.year:04d}-{ts.month:02d}"


def add_months(ym, n):
    y, m = map(int, ym.split("-"))
    idx = (y * 12 + (m - 1)) + n
    return f"{idx // 12:04d}-{idx % 12 + 1:02d}"


def parse_period(period):
    p = period.strip().upper()
    if "-Q" in p:
        y, q = p.split("-Q")
        s = (int(q) - 1) * 3 + 1
        return "quarterly", [f"{int(y):04d}-{m:02d}" for m in range(s, s + 3)], int(y), int(q)
    parts = p.split("-")
    if len(parts) == 2:
        return "monthly", [f"{int(parts[0]):04d}-{int(parts[1]):02d}"], int(parts[0]), None
    return "annual", [f"{int(parts[0]):04d}-{m:02d}" for m in range(1, 13)], int(parts[0]), None


def _sum_over(by_month, months):
    return sum(by_month.get(m, 0) for m in months)


def _latest_month(data):
    latest = None
    for key, col in (("leads", "arrival"), ("projects", "start"),
                     ("partnerships", "executed"), ("income", "date"), ("tasks", "updated")):
        df = data.get(key)
        if df is None or col not in df:
            continue
        s = df[col].dropna()
        if len(s):
            mx = mk(s.max())
            if mx and (latest is None or mx > latest):
                latest = mx
    return latest


# ----------------------------------------------------------------- blocks
def _leads(data, months):
    df = data.get("leads")
    out = {"by_month": {}, "by_source": {}, "by_type": {}, "total": 0}
    if df is None or not len(df):
        return out
    df = df.assign(_m=df["arrival"].map(mk))
    inp = df[df["_m"].isin(months)]
    out["total"] = int(len(inp))
    for m in months:
        out["by_month"][m] = int((inp["_m"] == m).sum())
    for src, g in inp.groupby(inp["source"].fillna("לא צוין")):
        out["by_source"][src] = {m: int((g["_m"] == m).sum()) for m in months}
    for t, g in inp.groupby(inp["client_type"].fillna("לא צוין")):
        out["by_type"][t] = {m: int((g["_m"] == m).sum()) for m in months}
    return out


def _pipeline(data, months, missing):
    proj = data.get("projects")
    part = data.get("partnerships")
    deals = {m: 0 for m in months}
    value = {m: 0.0 for m in months}
    if proj is not None and len(proj):
        for _, r in proj.iterrows():
            m = mk(r["start"])
            if m not in deals:
                continue
            # Every started project is a deal. הופסק → use ONLY שכ"ט-אם-הופסק
            # (it replaces the full fee — never both; empty ⇒ 0). Otherwise full fee.
            stopped = bool(r.get("status")) and _STOPPED in nfc(r["status"])
            fee = r.get("fee_stopped") if stopped else r.get("fee")
            deals[m] += 1
            value[m] += float(fee or 0)
    if part is not None and len(part):
        for _, r in part.iterrows():
            m = mk(r["executed"])
            if m not in deals:
                continue
            st = nfc(r["status"]) if r.get("status") else ""
            cancelled = _P_CANCEL in st
            closed = any(tok in st for tok in _P_CLOSED)
            if not (closed or cancelled):
                continue
            # בוטל → use ONLY שכ"ט-אם-בוטל (replaces full fee). Otherwise full fee.
            fee = r.get("fee_cancelled") if cancelled else r.get("fee")
            deals[m] += 1
            value[m] += float(fee or 0)
    return {"deals_by_month": deals, "value_by_month": value,
            "deals_total": _sum_over(deals, months), "value_total": _sum_over(value, months)}


def _cash(data, months):
    df = data.get("income")
    gross = {m: 0.0 for m in months}
    net = {m: 0.0 for m in months}
    pay = {m: 0 for m in months}
    if df is not None and len(df):
        for _, r in df.iterrows():
            m = mk(r["date"])
            if m not in gross:
                continue
            g = float(r.get("gross") or 0)
            issued = r.get("invoice") and nfc(r["invoice"]).strip() == _YES
            gross[m] += g
            net[m] += g / VAT if issued else g
            pay[m] += 1
    return {"gross_by_month": gross, "net_by_month": net, "payments_by_month": pay,
            "gross_total": _sum_over(gross, months), "net_total": _sum_over(net, months),
            "payments_total": _sum_over(pay, months)}


def _expenses(data):
    df = data.get("expenses")
    out = {"total": 0.0, "by_category": {}, "fixed": 0.0, "variable": 0.0}
    if df is None or not len(df):
        return out
    for _, r in df.iterrows():
        amt = float(r.get("amount") or 0)
        cat = r.get("category") or "לא צוין"
        out["total"] += amt
        out["by_category"][cat] = out["by_category"].get(cat, 0.0) + amt
        et = nfc(r.get("etype")) if r.get("etype") else ""
        if "משתנ" in et:           # משתנה
            out["variable"] += amt
        else:
            out["fixed"] += amt
    return out


def _tasks(data, months):
    df = data.get("tasks")
    out = {"by_performer_by_month": {}, "by_month": {m: 0 for m in months},
           "total": 0, "avg": 0.0, "performers": []}
    if df is None or not len(df):
        return out
    done = df[df["status"].map(lambda s: bool(s) and _DONE in nfc(s))].copy()
    done["_m"] = done["updated"].map(mk)
    done = done[done["_m"].isin(months)]
    out["total"] = int(len(done))
    for m in months:
        out["by_month"][m] = int((done["_m"] == m).sum())
    performers = sorted(p for p in done["performer"].dropna().unique())
    out["performers"] = performers
    for p in performers:
        g = done[done["performer"] == p]
        out["by_performer_by_month"][p] = {m: int((g["_m"] == m).sum()) for m in months}
    out["avg"] = round(out["total"] / len(months), 1) if months else 0.0
    return out


def _cohort(data, today_month, missing):
    leads = data.get("leads")
    proj = data.get("projects")
    out = {"rows": [], "headline_rate": None, "headline_n": 0, "headline_conv": 0,
           "ttc_median": None}
    if leads is None or not len(leads):
        return out
    # map normalized client name -> earliest project start
    proj_starts = {}
    if proj is not None and len(proj):
        for _, r in proj.iterrows():
            nm = norm_name(r.get("client"))
            if not nm or pd.isna(r.get("start")):
                continue
            if nm not in proj_starts or r["start"] < proj_starts[nm]:
                proj_starts[nm] = r["start"]
    cutoff = add_months(today_month, -MATURING_MONTHS) if today_month else None
    by_month = {}
    ttc_days = []
    for _, r in leads.iterrows():
        am = mk(r.get("arrival"))
        if not am:
            continue
        rec = by_month.setdefault(am, {"n": 0, "c": 0})
        rec["n"] += 1
        nm = norm_name(r.get("name"))
        start = proj_starts.get(nm)
        if start is not None and start >= r["arrival"]:
            rec["c"] += 1
            ttc_days.append((start - r["arrival"]).days)
    hl_n = hl_c = 0
    for am in sorted(by_month):
        rec = by_month[am]
        maturing = bool(cutoff) and am > cutoff
        rate = rec["c"] / rec["n"] if rec["n"] else 0.0
        out["rows"].append({"month": am, "n_leads": rec["n"], "n_conv": rec["c"],
                            "rate": rate, "maturing": maturing})
        if not maturing:
            hl_n += rec["n"]
            hl_c += rec["c"]
    out["headline_n"] = hl_n
    out["headline_conv"] = hl_c
    out["headline_rate"] = (hl_c / hl_n) if hl_n else None
    if ttc_days:
        out["ttc_median"] = float(pd.Series(ttc_days).median())
    if out["headline_rate"] is None:
        missing.append("אחוז המרה: אין קוהורטות בשלות")  # no mature cohorts
    return out


def _attribution(data, months, missing):
    leads = data.get("leads")
    proj = data.get("projects")
    out = {"matched": 0, "total": 0, "pct": None, "unmatched": []}
    if proj is None or not len(proj):
        return out
    lead_names = set()
    if leads is not None and len(leads):
        lead_names = {norm_name(n) for n in leads["name"].dropna()}
    started = proj[proj["start"].map(lambda t: mk(t) in months)]
    out["total"] = int(len(started))
    for _, r in started.iterrows():
        nm = norm_name(r.get("client"))
        if nm and nm in lead_names:
            out["matched"] += 1
        else:
            out["unmatched"].append(str(r.get("name") or r.get("client") or "?"))
    out["pct"] = (out["matched"] / out["total"]) if out["total"] else None
    # Attribution is INFORMATIONAL only — a non-match means the client arrived before
    # the lookback window or via referral/repeat/direct, NOT a data problem. The
    # client field is always present; never recorded as a missing/data-quality item.
    return out


def compute(data, period):
    zoom, months, year, quarter = parse_period(period)
    missing = []
    profit_net = _cash(data, months)
    exp = _expenses(data)
    profit = {m: profit_net["net_by_month"][m] - exp["total"] for m in months}
    today = _latest_month(data) or months[-1]
    return {
        "period": period, "zoom": zoom, "months": months, "year": year, "quarter": quarter,
        "today_month": today,
        "leads": _leads(data, months),
        "pipeline": _pipeline(data, months, missing),
        "cash": profit_net,
        "expenses": exp,
        "profit": {"by_month": profit, "total": sum(profit.values())},
        "tasks": _tasks(data, months),
        "cohort": _cohort(data, today, missing),
        "attribution": _attribution(data, months, missing),
        "targets": None,
        "missing": missing,
    }
