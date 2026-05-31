# -*- coding: utf-8 -*-
"""Quarterly targets — read the targets TABLE from the filled Word, if one exists.

The Word's metric names are matched (by substring) to dashboard metrics. Insights /
initiatives are human text and are never parsed for numbers. python-docx is imported
lazily so the engine runs even when it is not installed and no Word is present.
"""
import os
import re
import shutil
import unicodedata

from scripts import style_config as sc

# Word metric label  ->  dashboard metric key. Matching is substring-based (tolerant).
LABEL_MAP = [
    (("פני", "ליד"), "leads"),                 # פניות / לידים
    (("עסק", "נסגר"), "deals"),          # עסקאות שנסגרו
    (("מחזור", "שווי", "פייפ"), "pipe_value"),  # מחזור/שווי/פייפליין
    (("נטו",), "net"),                                       # נטו
    (("ברוטו",), "gross"),                         # ברוטו
    (("רווח",), "profit"),                              # רווח
    (("המר", "המרה"), "conversion"),     # המרה
    (("תקרת", "הוצא"), "exp_ceiling"),  # תקרת הוצאות
    (("משימ",), "tasks"),                               # משימות
]


def _nfc(s):
    return unicodedata.normalize("NFC", str(s))


def _match_label(text):
    n = _nfc(text)
    for pats, key in LABEL_MAP:
        if any(p in n for p in pats):
            return key
    return None


def _num(text):
    m = re.findall(r"-?\d[\d,]*\.?\d*", _nfc(text))
    if not m:
        return None
    try:
        return float(m[0].replace(",", ""))
    except ValueError:
        return None


def _meeting_docs(inputs_root):
    """All meeting Word files across inputs/YYYY-MM/, newest month folder first."""
    found = []
    if not os.path.isdir(inputs_root):
        return found
    for mon in sorted((n for n in os.listdir(inputs_root)
                       if re.fullmatch(r"\d{4}-\d{2}", n)), reverse=True):
        d = os.path.join(inputs_root, mon)
        if not os.path.isdir(d):
            continue
        for name in sorted(os.listdir(d)):
            low = name.lower()
            if low.endswith(".docx") and not name.startswith("~$") \
               and ("meeting" in low or "quarter" in low or "רבעון" in name):
                found.append(os.path.join(d, name))
    return found


def find_targets(inputs_root, period):
    """Carry-forward: return {key: value} from the most recent FILLED meeting Word
    across the per-month folders, or None. A blank template yields no numbers -> None.
    The quarterly meeting Word lives in the quarter-closing month folder and carries
    the next quarter's targets; the latest filled one applies until a newer replaces it."""
    for path in _meeting_docs(inputs_root):
        t = read_targets(path)
        if t:
            return t
    return None


def read_targets(path):
    """Return {metric_key: value} from the Word targets table, or None."""
    if not path or not os.path.isfile(path):
        return None
    try:
        import docx  # python-docx, lazy
    except ImportError:
        return None
    doc = docx.Document(path)
    targets = {}
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text for c in row.cells]
            if len(cells) < 2:
                continue
            key = _match_label(cells[0])
            val = _num(cells[1])
            if key and val is not None:
                targets[key] = val
    return targets or None


def ensure_blank_next_quarter(templates_dir, inputs_root, year, quarter):
    """On a quarter-closing run, drop a blank meeting Word into the closing month's
    folder (e.g. inputs/2026-03/ for Q1). The studio fills its targets table after the
    quarter meeting; those become the NEXT quarter's targets (carry-forward)."""
    if quarter is None:
        return None
    closing_month = quarter * 3                       # Q1->3, Q2->6, Q3->9, Q4->12
    folder = os.path.join(inputs_root, f"{year:04d}-{closing_month:02d}")
    dest = os.path.join(folder, f"meeting-{year}-Q{quarter}.docx")
    src = os.path.join(templates_dir, "meeting-template-quarterly.docx")
    if not os.path.isfile(src):
        return None
    os.makedirs(folder, exist_ok=True)
    if any(f.lower().endswith(".docx") for f in os.listdir(folder)):
        return None                                   # a meeting Word already there
    shutil.copyfile(src, dest)
    return os.path.relpath(dest, os.path.dirname(inputs_root))


# map dashboard metric key -> (label, computed value) for the targets block
def actual_for(report):
    c, p, cash, prof, ck = (report["leads"], report["pipeline"], report["cash"],
                            report["profit"], report["cohort"])
    return {
        "leads": (sc.M_LEADS, c["total"]),
        "deals": (sc.M_DEALS, p["deals_total"]),
        "pipe_value": (sc.M_PIPEVAL, p["value_total"]),
        "net": (sc.M_NET, cash["net_total"]),
        "gross": (sc.M_GROSS, cash["gross_total"]),
        "profit": (sc.M_PROFIT, prof["total"]),
        "conversion": (sc.M_CONV, ck["headline_rate"]),
        "tasks": (sc.M_TASKS, report["tasks"]["total"]),
        "exp_ceiling": (sc.M_EXP, report["expenses"]["total"]),
    }
