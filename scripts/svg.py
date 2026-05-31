# -*- coding: utf-8 -*-
"""Dependency-free inline SVG charts for the HTML dashboard.

Pure functions: same inputs -> byte-identical SVG (every coordinate passes through the
single `_f` rounding gate; no randomness, stable order). Charts carry visible data
labels. RTL-aware: horizontal bars grow leftward with the category label on the right;
line categories run right-to-left (first period on the right) to mirror the RTL Excel
sheet. No external/CDN/network references — the SVG is embedded directly in the page.

This module imports nothing project-specific (colors and formatted labels are passed
in) so it stays a small, testable rendering primitive.
"""
import math

# palette defaults (mirror style_config; callers may override)
_INK = "#162C3A"
_MUTED = "#5C6670"
_TITLE = "#162C3A"
_DIVIDER = "#E2E6EA"
_GRID = "#EAEDEF"
_DEFAULT_BAR = "#1F3A4D"
_LINE_COLORS = ["#1F3A4D", "#3E7C8C", "#5C6670"]


def _f(x):
    """The single coordinate-rounding gate — keeps output deterministic and compact."""
    return "{:.2f}".format(float(x))


def _esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _clip(label, n):
    s = str(label)
    return s if len(s) <= n else s[: n - 1] + "…"


def _money(v):
    return "—" if v is None or round(v) == 0 else "₪ {:,.0f}".format(v)


def _int(v):
    return "—" if v is None or round(v) == 0 else "{:,.0f}".format(v)


def _text(x, y, s, anchor="start", size=11, color=_INK, bold=False, rtl=False):
    w = ' font-weight="bold"' if bold else ""
    d = ' direction="rtl"' if rtl else ""
    return ('<text x="{x}" y="{y}" text-anchor="{a}" font-family="Arial,sans-serif" '
            'font-size="{s}" fill="{c}"{w}{d}>{t}</text>').format(
        x=_f(x), y=_f(y), a=anchor, s=size, c=color, w=w, d=d, t=_esc(s))


def _open(width, height):
    # width:100% + height:auto -> fills the container width, height scales with the
    # viewBox aspect (no letterboxing, no hard px cap).
    return ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" '
            'width="100%" height="auto" role="img" preserveAspectRatio="xMidYMid meet" '
            'style="display:block;font-family:inherit">').format(w=_f(width), h=_f(height))


def bar_chart(series, width=580, color=None, money=False, title="", max_label=22):
    """Horizontal bars (RTL) — used for the single-month income composition. Roomy
    label/value gutters and large type so it never crowds."""
    color = color or _DEFAULT_BAR
    fmt = _money if money else _int
    n = len(series)
    title_h = 26 if title else 8
    rh = 48
    height = title_h + n * rh + 8
    lab_w, val_w = 156, 96
    x_base = width - lab_w
    x_min = val_w
    maxv = max((abs(v or 0) for _, v in series), default=0) or 1

    p = [_open(width, height)]
    if title:
        p.append(_text(width - 8, 18, title, anchor="end", size=13, color=_TITLE, bold=True, rtl=True))
    for i, (label, value) in enumerate(series):
        cy = title_h + i * rh + rh / 2.0
        barlen = (abs(value or 0) / maxv) * (x_base - x_min)
        x1 = x_base - barlen
        p.append('<rect x="{x}" y="{y}" width="{w}" height="22" rx="5" fill="{c}"/>'.format(
            x=_f(x1), y=_f(cy - 11), w=_f(max(barlen, 2)), c=color))
        p.append(_text(width - 8, cy + 5, _clip(label, max_label), anchor="end", size=13, color=_INK, rtl=True))
        p.append(_text(x1 - 9, cy + 5, fmt(value), anchor="end", size=12, color=_MUTED))
    p.append("</svg>")
    return "".join(p)


def line_chart(series_set, categories, width=540, height=250, colors=None,
               money=False, labels=True, title=""):
    """Multi-series line chart. `series_set` = [(name, [values...]), ...]; `categories`
    are pre-formatted axis labels. Categories run right-to-left. Optional data labels
    above each vertex (caller gates by month count to avoid clutter)."""
    colors = colors or _LINE_COLORS
    fmt = _money if money else _int
    n = len(categories)
    pad_l, pad_r, pad_t, pad_b = 18, 18, 24, 30
    plot_top, plot_bottom = pad_t, height - pad_b
    plot_left, plot_right = pad_l, width - pad_r

    allvals = [val for _, vals in series_set for val in vals if val is not None]
    vmax = max(allvals, default=0)
    vmin = min(0, min(allvals, default=0))
    span = (vmax - vmin) or 1

    def yof(val):
        return plot_bottom - (val - vmin) / span * (plot_bottom - plot_top)

    def xof(i):
        return plot_right if n == 1 else plot_right - i / (n - 1.0) * (plot_right - plot_left)

    p = [_open(width, height)]
    for g in range(5):                                  # soft horizontal gridlines
        gy = plot_top + g / 4.0 * (plot_bottom - plot_top)
        p.append('<line x1="{xl}" y1="{y}" x2="{xr}" y2="{y}" stroke="{c}" stroke-width="1"/>'.format(
            xl=_f(plot_left), y=_f(gy), xr=_f(plot_right), c=_GRID))
    vals0 = series_set[0][1]                             # area wash under the first series
    pts0 = [(xof(i), yof(val)) for i, val in enumerate(vals0) if val is not None]
    if len(pts0) >= 2:
        d = "M " + " L ".join("{} {}".format(_f(x), _f(y)) for x, y in pts0)
        d += " L {} {} L {} {} Z".format(_f(pts0[-1][0]), _f(plot_bottom), _f(pts0[0][0]), _f(plot_bottom))
        p.append('<path d="{d}" fill="{c}" opacity=".08"/>'.format(d=d, c=colors[0]))
    for i, cat in enumerate(categories):
        p.append(_text(xof(i), plot_bottom + 18, cat, anchor="middle", size=11, color=_MUTED, rtl=True))
    for si, (name, vals) in enumerate(series_set):
        col = colors[si % len(colors)]
        pts = [(xof(i), yof(val)) for i, val in enumerate(vals) if val is not None]
        if len(pts) >= 2:
            p.append('<polyline points="{pts}" fill="none" stroke="{c}" stroke-width="2.5" '
                     'stroke-linejoin="round" stroke-linecap="round"/>'.format(
                         pts=" ".join("{},{}".format(_f(x), _f(y)) for x, y in pts), c=col))
        for i, val in enumerate(vals):
            if val is None:
                continue
            x, y = xof(i), yof(val)
            p.append('<circle cx="{x}" cy="{y}" r="3.2" fill="#fff" stroke="{c}" stroke-width="2"/>'.format(
                x=_f(x), y=_f(y), c=col))
            if labels:
                p.append('<text x="{x}" y="{y}" text-anchor="middle" font-size="10" font-weight="600" '
                         'fill="{c}" paint-order="stroke" stroke="#fff" stroke-width="3" '
                         'stroke-linejoin="round">{t}</text>'.format(x=_f(x), y=_f(y - 9), c=col, t=_esc(fmt(val))))
    p.append("</svg>")
    return "".join(p)


def donut(pct, size=150, color=None, track=None, sub=""):
    """A single-value gauge — the headline % in the center (text-anchor=middle, so no
    RTL edge issues). Renders for every period, keeping an SVG with a numeric label
    present in the page."""
    color = color or "#3E7C8C"
    track = track or _GRID
    pct = max(0.0, min(1.0, float(pct or 0)))
    c = size / 2.0
    r = size / 2.0 - 15
    circ = 2 * math.pi * r
    dash = circ * pct
    p = [_open(size, size)]
    p.append('<circle cx="{c}" cy="{c}" r="{r}" fill="none" stroke="{t}" stroke-width="13"/>'.format(
        c=_f(c), r=_f(r), t=track))
    p.append('<circle cx="{c}" cy="{c}" r="{r}" fill="none" stroke="{col}" stroke-width="13" '
             'stroke-linecap="round" stroke-dasharray="{d} {g}" transform="rotate(-90 {c} {c})"/>'.format(
                 c=_f(c), r=_f(r), col=color, d=_f(dash), g=_f(circ - dash)))
    p.append('<text x="{c}" y="{y}" text-anchor="middle" font-size="27" font-weight="700" '
             'fill="{ink}">{pct}</text>'.format(c=_f(c), y=_f(c + 3), ink=_INK,
                                                pct="{:.1f}%".format(pct * 100)))
    if sub:
        p.append('<text x="{c}" y="{y}" text-anchor="middle" font-size="11" fill="{m}">{s}</text>'.format(
            c=_f(c), y=_f(c + 22), m=_MUTED, s=_esc(sub)))
    p.append("</svg>")
    return "".join(p)
