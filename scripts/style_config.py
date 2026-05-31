# -*- coding: utf-8 -*-
"""Frozen visual configuration — the single machine-readable source of style.

Both the Excel generator and the HTML preview import from here so every period
looks identical. Mirrors references/style.md. Change style HERE (and there), never
ad-hoc inside the generator.

Design language (modern-professional, restrained):
  * ONE primary accent + ONE secondary accent + grays. At most 2 accents per sheet.
  * No gradients, no 3D, no per-category color. Status colors (green/amber/red) are
    used for TARGETS only (and the profit sign).
  * Arial throughout; ₪ before the amount; zero renders as an em dash (—).
"""
import base64
import functools
import os
import struct

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---------------------------------------------------------------- colors (hex)
PRIMARY      = "1F3A4D"   # accent 1 — deep petrol-blue (section bands, KPI top-border, Picture A)
PRIMARY_DK   = "162C3A"   # darker shade of PRIMARY for the main title
SECONDARY    = "3E7C8C"   # accent 2 — muted turquoise (2nd chart series, highlights, Picture B)
TEXT         = "1A1A1A"   # ink / body text
MUTED        = "5C6670"   # secondary text, labels, Δ chips, notes
WHITE        = "FFFFFF"   # base background
CARD         = "F4F6F8"   # KPI cards, zebra band, light strips
DIVIDER      = "E2E6EA"   # separator lines, cell borders, column-header strip
GOOD         = "2E7D5B"   # target met / positive (TARGETS + profit sign only)
WARN         = "C28A2B"   # behind target
BAD          = "B23A3A"   # over / loss

# back-compat aliases (kept so existing imports resolve to the new palette)
HEADER_FILL  = DIVIDER    # inner column-header strip (light gray)
SUBHEAD_FILL = CARD       # secondary header
BAND         = CARD       # zebra band for dynamic (open-schema) rows
GRID         = DIVIDER    # cell borders
PIC_A        = PRIMARY    # Picture A — pipeline ("money closed")
PIC_B        = SECONDARY  # Picture B — cash ("money in") — kept visually distinct

FONT_NAME = "Arial"
STUDIO_NAME = "סטודיו מרים בסקון"

# ---------------------------------------------------------------- type sizes
SZ_TITLE    = 20   # main dashboard title (bold)
SZ_SUBTITLE = 12   # studio subtitle / period (muted)
SZ_SECTION  = 13   # section header (bold, white on accent)
SZ_KPI      = 26   # KPI card value (bold)
SZ_KPILABEL = 10   # KPI card label (muted)
SZ_DELTA    = 10   # KPI Δ chip
SZ_HEAD     = 11   # column header (bold)
SZ_BODY     = 11   # body
SZ_NOTE      = 9   # notes / missing-items / disclaimers (muted)

# ---------------------------------------------------------------- number formats
# 3-section: positive ; negative ; zero(→ em dash). ₪ precedes the amount.
FMT_MONEY  = '"₪ "#,##0;"-₪ "#,##0;"—"'
FMT_MONEY0 = FMT_MONEY
FMT_PCT    = '0.0%;-0.0%;"—"'
FMT_INT    = '#,##0;-#,##0;"—"'
FMT_DAYS   = '#,##0" ימים";-#,##0" ימים";"—"'  # N ימים (days)
FMT_MONTH  = 'mm/yyyy'
FMT_DATE   = 'dd/mm/yyyy'                          # detail-sheet date columns

# ---------------------------------------------------------------- canonical labels
T_OVERVIEW = "סקירה"                             # overview sheet name
# Block titles
T_LEADS   = "פניות (לידים)"
T_PIPE    = "עסקאות ושווי — תמונה א׳ (פייפליין)"
T_CASH    = "הכנסה בפועל — תמונה ב׳ (תזרים)"
T_EXP     = "הוצאות ורווח"
T_TASKS   = "משימות / תפוקה"
T_COHORT  = "המרת קוהורט"
T_ATTR    = "שיוך עסקאות ללידים (אינפורמטיבי)"
T_MISSING = "דוח חוסרים"
T_TARGETS = "עמידה ביעדים"
# Side-by-side income block headers (the two pictures, never merged)
T_PIC_A   = "פייפליין (נסגר) — תמונה א׳"
T_PIC_B   = "תזרים (נכנס) — תמונה ב׳"
T_TREND   = "מגמות"
T_KPI     = "מדדי מפתח"

# Metric labels (also used to match the quarterly Word targets table — DO NOT rename)
M_LEADS   = "מספר פניות"
M_DEALS   = "מספר עסקאות שנסגרו"
M_PIPEVAL = "שווי עסקאות (₪)"
M_NET     = "הכנסה נטו (₪)"
M_GROSS   = "הכנסה ברוטו (₪)"
M_PAYMENTS= "מספר תשלומים"
M_EXP     = "סך הוצאות (₪)"
M_PROFIT  = "רווח (₪)"
M_CONV    = "אחוז המרה"
M_CONV_SETTLED = "שיעור המרה מיוצב"
M_TASKS   = "משימות שהושלמו"
M_TTC     = "זמן המרה (חציון ימים)"

# Short KPI-card labels
K_NET    = "הכנסה נטו"
K_PROFIT = "רווח"
K_DEALS  = "עסקאות שנסגרו"
K_LEADS  = "פניות"
K_CONV   = "המרה מיוצבת"
K_TASKS  = "משימות שהושלמו"

# Detail sheet names
DT_PROJECTS = "פרויקטים"
DT_LEADS    = "לידים"
DT_INCOME   = "הכנסות"
DT_PART     = "שתפים"
DT_TASKS    = "משימות"
DT_EXPENSES = "הוצאות"   # row-level expenses detail sheet (≠ T_EXP "הוצאות ורווח")

# Misc strings
S_TOTAL   = "סה״כ"
S_SUMMARY = "סיכום"
S_AVG     = "ממוצע חודשי"
S_NA      = "לא צוין"
S_NEW     = "(חדש)"
S_MATURING= "עוד מבשיל"
S_MONTH   = "חודש"
S_VALUE   = "ערך"
S_TARGET  = "יעד"
S_ACTUAL  = "בפועל"
S_PROGRESS= "התקדמות"
S_GENERATED = "תאריך הפקה"
S_NODATA  = "אין נתונים לתקופה"
S_METRIC  = "מדד"
# Time-to-close caveat (the ~3-day figure is biased low until lead history grows).
S_TTC_CAVEAT = "מבוסס על עסקאות שחוברו לליד; מוטה כלפי מטה עד שהיסטוריית הלידים תתרחב"

HEB_MONTHS = ["ינואר", "פברואר", "מרץ",
              "אפריל", "מאי", "יוני",
              "יולי", "אוגוסט", "ספטמבר",
              "אוקטובר", "נובמבר", "דצמבר"]


def heb_month(ym):
    """'2026-03' -> 'מרץ 2026'."""
    y, m = ym.split("-")
    return f"{HEB_MONTHS[int(m) - 1]} {y}"


# ---------------------------------------------------------------- openpyxl helpers
def font(size=SZ_BODY, bold=False, color=TEXT, italic=False):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color, italic=italic)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def center(wrap=True):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap, readingOrder=2)


def right(wrap=True):
    return Alignment(horizontal="right", vertical="center", wrap_text=wrap, readingOrder=2)


def left(wrap=False):
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap, readingOrder=2)


def side_thin(color=DIVIDER):
    return Side(style="thin", color=color)


def side_thick(color=PRIMARY):
    return Side(style="thick", color=color)   # ~3px — KPI card accent top border


def thin_border(color=GRID):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def no_border():
    return Border()


LRI = "⁦"  # left-to-right isolate
PDI = "⁩"  # pop directional isolate


def lrm(s):
    """Wrap a Latin/number token (e.g. 'Q1 2026') in an LTR isolate so it keeps its
    internal order AND cannot affect surrounding RTL text (balanced, self-closing)."""
    return LRI + str(s) + PDI


# ---------------------------------------------------------------- studio logo
# A brand asset (e.g. "לוגו ללא רקע.png") placed in the project root. Discovered by
# extension so the Hebrew filename never has to be typed on a shell line; embedded as
# a base64 data-URI in the HTML (keeps it self-contained) and as an openpyxl image in
# the Excel header. Deterministic: same bytes → identical output every run.
_IMG_EXT = (".png", ".svg", ".jpg", ".jpeg")
_MIME = {".png": "image/png", ".svg": "image/svg+xml", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}


@functools.lru_cache(maxsize=1)
def logo_path():
    """Absolute path to the studio logo in the project root, or None if absent."""
    cands = [f for f in os.listdir(ROOT) if os.path.splitext(f)[1].lower() in _IMG_EXT]
    if not cands:
        return None
    cands.sort(key=lambda f: (_IMG_EXT.index(os.path.splitext(f)[1].lower()), f))  # prefer .png, then name
    return os.path.join(ROOT, cands[0])


@functools.lru_cache(maxsize=1)
def logo_bytes():
    p = logo_path()
    if not p:
        return None
    with open(p, "rb") as f:
        return f.read()


@functools.lru_cache(maxsize=1)
def logo_data_uri():
    """'data:image/png;base64,...' for inline HTML embedding, or None."""
    b = logo_bytes()
    if b is None:
        return None
    ext = os.path.splitext(logo_path())[1].lower()
    return "data:%s;base64,%s" % (_MIME.get(ext, "image/png"), base64.b64encode(b).decode("ascii"))


@functools.lru_cache(maxsize=1)
def logo_dims():
    """(width, height) in px for a PNG logo (reads IHDR; no extra deps), else None."""
    b = logo_bytes()
    if not b or b[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    return struct.unpack(">II", b[16:24])


# CSS palette for the HTML preview (mirrors the above)
def css_vars():
    return {
        "primary": "#" + PRIMARY, "primary_dk": "#" + PRIMARY_DK,
        "secondary": "#" + SECONDARY,
        "card": "#" + CARD, "divider": "#" + DIVIDER,
        "text": "#" + TEXT, "muted": "#" + MUTED, "white": "#" + WHITE,
        "good": "#" + GOOD, "warn": "#" + WARN, "bad": "#" + BAD,
        "pic_a": "#" + PIC_A, "pic_b": "#" + PIC_B, "font": FONT_NAME,
        # legacy keys still referenced by older templates
        "header": "#" + HEADER_FILL, "subhead": "#" + SUBHEAD_FILL,
        "band": "#" + BAND, "grid": "#" + GRID,
    }
