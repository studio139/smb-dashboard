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

from scripts import metrics
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


# the targets Word's filename tag — mirrors TYPE_QUARTERLY ("רבעוני") on that quarter's reports
TARGETS_SUFFIX = "יעדים"


def prev_quarter_close(period):
    """(year, quarter) of the quarter-closing whose targets doc a report READS — the quarter
    BEFORE the report's own. Cross-year aware: Q1 of year Y reads (Y-1, Q4). Annual → None.
    A quarter's doc is filled at its close and carries the NEXT quarter's targets, so a report
    in quarter Q reads quarter Q-1's doc (e.g. April→Q2 reads Q1; January→Q1 reads prior-year Q4)."""
    zoom, months, year, quarter = metrics.parse_period(period)
    if zoom == "annual":
        return None
    if quarter is None:                                   # monthly: derive the month's quarter
        quarter = (int(months[0].split("-")[1]) - 1) // 3 + 1
    if quarter == 1:
        return (year - 1, 4)
    return (year, quarter - 1)


def targets_doc_name(year, quarter):
    """The targets Word's filename — mirrors the report naming (e.g. 2026-Q1_יעדים.docx)."""
    return f"{year}-Q{quarter}_{TARGETS_SUFFIX}.docx"


def targets_doc_path(outputs_root, year, quarter):
    """Full path of a quarter's targets Word in the outputs tree:
    outputs/<year>/<רבעוני>/<year>-Q<quarter>/<year>-Q<quarter>_יעדים.docx."""
    leaf = f"{year}-Q{quarter}"
    return os.path.join(outputs_root, str(year), sc.TYPE_QUARTERLY, leaf, targets_doc_name(year, quarter))


def find_targets(outputs_root, period):
    """Read the targets the report should apply: resolve the PREVIOUS quarter-close doc in the
    outputs tree (cross-year aware) and parse its table. Returns {key: value} from a FILLED doc,
    or None when that doc is missing or still the blank template (no crash either way). One doc
    serves all three months of its target quarter — that single mapping IS the carry-forward."""
    src = prev_quarter_close(period)
    if src is None:
        return None
    return read_targets(targets_doc_path(outputs_root, src[0], src[1]))


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


def ensure_targets_doc(templates_dir, period_dir, year, quarter):
    """On a quarter close, emit the blank targets Word into that quarter's OUTPUT folder
    (period_dir, e.g. outputs/2026/רבעוני/2026-Q1/). The studio fills it after the quarter
    meeting; those become the NEXT quarter's targets, read forward by find_targets. Created
    ONLY if absent — the output folder is rebuilt every run, but a filled doc is preserved.
    Returns the path when freshly emitted, else None (already present, or no template)."""
    if quarter is None:
        return None
    src = os.path.join(templates_dir, "meeting-template-quarterly.docx")
    if not os.path.isfile(src):
        return None
    dest = os.path.join(period_dir, targets_doc_name(year, quarter))
    if os.path.isfile(dest):
        return None                                   # a (maybe filled) doc is already here
    os.makedirs(period_dir, exist_ok=True)
    shutil.copyfile(src, dest)
    return dest


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
