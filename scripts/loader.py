# -*- coding: utf-8 -*-
"""Input loader — identifies the Hebrew export files and returns canonical frames.

Robust to Fireberry export artifacts:
  * column headers normalized (trailing quotes / spaces stripped, blank cols dropped)
  * dates accepted as Excel serials, datetimes, or text
  * amounts accepted as numbers or text ("3,000 (ממוצע)", "הלוואה - 40,000")
  * "-" / empty treated as a no-link sentinel, never as a value

The loader searches inputs/current/ first, then the project root, so it works
whether or not the one-time reorg has run.
"""
import os
import re
import unicodedata
from datetime import datetime, timedelta

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SENTINELS = {"", "-", "–", "—", "none", "nan", "null"}
_EXCEL_EPOCH = datetime(1899, 12, 30)


def nfc(s):
    return unicodedata.normalize("NFC", str(s))


def norm_header(h):
    """Strip whitespace and trailing gershayim/quote artifacts from a header."""
    s = nfc(h).strip()
    s = s.strip('"\'“”″ ')
    return s.strip()


def is_sentinel(v):
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return nfc(v).strip().lower() in SENTINELS


def to_text(v):
    return None if is_sentinel(v) else nfc(v).strip()


def to_number(v):
    """Coerce numbers or numeric-bearing text to float; None if impossible."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = nfc(v)
    m = re.findall(r"-?\d[\d,]*\.?\d*", s)
    if not m:
        return None
    try:
        return float(m[0].replace(",", ""))
    except ValueError:
        return None


def to_date(v):
    """Coerce Excel serial / datetime / text to a pandas Timestamp (or NaT)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return pd.NaT
    if isinstance(v, datetime):
        return pd.Timestamp(v)
    if isinstance(v, (int, float)):
        if 20000 < float(v) < 90000:                 # plausible Excel serial
            return pd.Timestamp(_EXCEL_EPOCH + timedelta(days=float(v)))
        return pd.NaT
    s = nfc(v).strip()
    if is_sentinel(s):
        return pd.NaT
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def latest_input_dir():
    """Latest inputs/YYYY-MM snapshot folder (or None)."""
    base = os.path.join(ROOT, "inputs")
    if not os.path.isdir(base):
        return None
    months = sorted(n for n in os.listdir(base)
                    if re.fullmatch(r"\d{4}-\d{2}", n)
                    and os.path.isdir(os.path.join(base, n)))
    return os.path.join(base, months[-1]) if months else None


def _find_file(input_dir, *patterns):
    """First xlsx whose NFC name contains any pattern: month snapshot, then root."""
    search_dirs = [d for d in (input_dir, ROOT) if d]
    for d in search_dirs:
        if not os.path.isdir(d):
            continue
        for name in sorted(os.listdir(d)):
            if not name.lower().endswith(".xlsx") or name.startswith("~$"):
                continue
            n = nfc(name)
            if any(p in n for p in patterns):
                return os.path.join(d, name)
    return None


def _read_raw(path):
    df = pd.read_excel(path, sheet_name=0, header=0, dtype=object, engine="openpyxl")
    df.columns = [norm_header(c) for c in df.columns]
    # drop fully-blank / unnamed columns
    keep = [c for c in df.columns if c and not c.lower().startswith("unnamed")]
    return df[keep]


def _pick(cols, *pats):
    """First column whose normalized name contains ALL given substrings."""
    for c in cols:
        n = nfc(c)
        if all(p in n for p in pats):
            return c
    return None


# --------------------------------------------------------------- per-file loaders
def _frame(df, mapping):
    """Build a canonical frame; record which canonical keys were not found."""
    cols = list(df.columns)
    out = pd.DataFrame(index=df.index)
    missing = []
    for key, picker in mapping.items():
        src = picker(cols)
        if src is None:
            out[key] = None
            missing.append(key)
        else:
            out[key] = df[src]
    return out, missing


def load_all(input_dir=None):
    """Return (data, report). data[name] = canonical DataFrame; reads the given
    inputs/YYYY-MM snapshot (default: latest) plus הוצאות קבועות.xlsx from root."""
    if input_dir is None:
        input_dir = latest_input_dir()
    data, report = {}, {}

    specs = {
        "leads": (("לקוח",), {                       # לקוח
            "name":   lambda c: _pick(c, "שם"),               # שם
            "arrival": lambda c: _pick(c, "הגע") or _pick(c, "תאריך"),  # הגע / תאריך
            "source": lambda c: _pick(c, "מקור"),   # מקור
            "client_type": lambda c: _pick(c, "סוג"),    # סוג
        }),
        "projects": (("פרויקט",), {        # פרויקט
            "name":   lambda c: _pick(c, "שם"),
            "client": lambda c: _pick(c, "לקוח"),
            "city":   lambda c: _pick(c, "עיר"),
            "fee_stopped": lambda c: _pick(c, "הופסק"),  # הופסק (check first)
            "fee":    lambda c: _pick(c, "שכט") or _pick(c, "שכר") or _pick(c, "טרח"),
            "start":  lambda c: _pick(c, "התחל") or _pick(c, "תאריך"),
            "status": lambda c: _pick(c, "סטטוס"),
            "ptype":  lambda c: _pick(c, "סוג"),
        }),
        "partnerships": (("שתפ", "שותף"), {  # שתפ / שותף
            "name":   lambda c: _pick(c, "שם"),
            "status": lambda c: _pick(c, "סטטוס"),
            "executed": lambda c: _pick(c, "תאריך"),
            "fee_cancelled": lambda c: _pick(c, "בוטל"),     # בוטל (check first)
            "fee":    lambda c: _pick(c, "כלל") or _pick(c, "שכר") or _pick(c, "שכט"),
        }),
        "income": (("הכנס",), {                     # הכנס
            "date":   lambda c: _pick(c, "תאריך"),
            "project": lambda c: _pick(c, "פרויקט"),
            "partner": lambda c: _pick(c, "שת"),
            "gross":  lambda c: _pick(c, "סכום"),    # סכום
            "invoice": lambda c: _pick(c, "חשבונית"),  # חשבונית
        }),
        "tasks": (("משימ",), {                      # משימ
            "performer": lambda c: _pick(c, "מבצע"),  # מבצע
            "status": lambda c: _pick(c, "סטטוס"),
            "updated": lambda c: _pick(c, "תאריך") or _pick(c, "עודכן"),
        }),
        "expenses": (("הוצאות", "קבועות"), {  # הוצאות / קבועות
            "category": lambda c: _pick(c, "קטגור"),  # קטגוריה
            "desc":   lambda c: _pick(c, "תיאור"),     # תיאור
            "amount": lambda c: _pick(c, "סכום"),           # סכום
            "etype":  lambda c: _pick(c, "סוג"),                 # סוג
        }),
    }

    for key, (pats, mapping) in specs.items():
        path = _find_file(input_dir, *pats)
        if path is None:
            data[key] = None
            report[key] = {"path": None, "rows": 0, "missing_cols": list(mapping.keys()), "found": False}
            continue
        raw = _read_raw(path)
        frame, missing = _frame(raw, mapping)
        frame = _coerce(key, frame)
        data[key] = frame
        report[key] = {
            "path": os.path.relpath(path, ROOT),
            "rows": int(len(frame)),
            "missing_cols": missing,
            "found": True,
        }
    return data, report


def _coerce(key, df):
    """Type-coerce canonical columns by name convention."""
    date_cols = {"arrival", "start", "executed", "date", "updated"}
    num_cols = {"fee", "fee_stopped", "fee_cancelled", "gross", "amount"}
    text_cols = {"name", "client", "city", "source", "client_type", "status",
                 "ptype", "project", "partner", "invoice", "performer",
                 "category", "desc", "etype"}
    for c in df.columns:
        if c in date_cols:
            df[c] = df[c].map(to_date)
        elif c in num_cols:
            df[c] = df[c].map(to_number)
        elif c in text_cols:
            df[c] = df[c].map(to_text)
    # drop rows that are entirely empty
    return df.dropna(how="all").reset_index(drop=True)
