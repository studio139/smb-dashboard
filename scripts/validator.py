# -*- coding: utf-8 -*-
"""Graded validation.

BLOCKING (stop and explain): a whole file missing, a required column missing,
core data empty, or an impossible total.
NON-BLOCKING (continue and record): a single missing field (-> "לא צוין"),
unmatched join rows, a single outlier. These feed the missing-items report.
"""
import pandas as pd

# canonical column that MUST resolve for each file to produce a reliable report
REQUIRED_COLS = {
    "leads": ["name", "arrival", "source"],
    "projects": ["name", "fee", "start", "status"],
    "partnerships": ["status", "executed", "fee"],
    "income": ["date", "gross", "invoice"],
    "tasks": ["performer", "status", "updated"],
    "expenses": ["category", "amount"],
}
# files whose emptiness is fatal (others degrade gracefully)
CORE_NONEMPTY = ["leads", "projects", "income", "expenses"]

OUTLIER_HI = 5_000_000   # ₪ — a single fee/amount above this is flagged for review


def validate_inputs(data, report):
    blocking, missing = [], []
    for key, cols in REQUIRED_COLS.items():
        info = report.get(key, {})
        if not info.get("found"):
            blocking.append(f"קובץ חסר: {key}")          # missing file
            continue
        absent = [c for c in cols if c in info.get("missing_cols", [])]
        if absent:
            blocking.append(f"עמודות חסרות ב-{key}: {', '.join(absent)}")
        df = data.get(key)
        if (df is None or not len(df)) and key in CORE_NONEMPTY:
            blocking.append(f"דאטה ריק: {key}")            # empty core data

    # non-blocking single-field gaps + outliers
    _note_gaps(data.get("leads"), "source", "ליד ללא מקור הגעה", missing)
    _note_gaps(data.get("leads"), "client_type", "ליד ללא סוג לקוח", missing)
    _note_outliers(data.get("projects"), "fee", "פרויקט", missing)
    _note_outliers(data.get("income"), "gross", "הכנסה", missing)
    return blocking, missing


def _note_gaps(df, col, label, missing):
    if df is None or col not in df:
        return
    n = int(df[col].isna().sum())
    if n:
        missing.append(f"{label}: {n}")


def _note_outliers(df, col, label, missing):
    if df is None or col not in df:
        return
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    hi = int((s > OUTLIER_HI).sum())
    if hi:
        missing.append(f"{label} — ערכים חריגים לבדיקה: {hi}")  # outliers to verify


def validate_report(report):
    """Impossible-total checks after computation -> blocking."""
    blocking = []
    cash = report["cash"]
    if cash["net_total"] - cash["gross_total"] > 1e-6:
        blocking.append("סכום בלתי-אפשרי: נטו > ברוטו")  # net>gross
    hr = report["cohort"]["headline_rate"]
    if hr is not None and (hr < 0 or hr > 1.0001):
        blocking.append("אחוז המרה מחוץ לטווח")  # conversion out of range
    for key in ("deals_total",):
        if report["pipeline"][key] < 0:
            blocking.append("ספירה שלילית")
    return blocking
