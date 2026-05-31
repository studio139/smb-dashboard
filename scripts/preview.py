# -*- coding: utf-8 -*-
"""Self-contained, interactive RTL BI dashboard — mirrors the Excel exactly.

Design split (the contract that keeps everything deterministic & test-safe):
  * PYTHON renders ALL content — every number, table and chart (inline <svg>) and the
    KPI `data-fig` anchors — into one deterministic HTML string, sourced only from the
    shared `viewmodel` (so Excel == HTML == metrics by construction).
  * Inline JS drives ONLY behavior — tab switching, KPI→section jumps, accordion
    expand/collapse, tooltips, the on-load reveal. It never produces a number.

One self-contained file: inline CSS + inline JS, no external/CDN/network, no web fonts.
Deterministic: identical bytes each run except the generation-date stamp.
"""
import datetime
import html

import pandas as pd

from scripts import style_config as sc
from scripts import targets as tg
from scripts import viewmodel as vm
from scripts import svg

esc = html.escape

TABS = [("overview", "סקירה"), ("finance", "כספים"), ("leads", "פניות והמרה"),
        ("tasks", "משימות"), ("detail", "פירוט")]

# terse HTML-only labels (Excel keeps the full sc.M_*/K_* labels; numbers unaffected)
_SHORT = {
    sc.M_NET: "נטו", sc.M_GROSS: "ברוטו", sc.M_PIPEVAL: "שווי", sc.M_DEALS: "עסקאות",
    sc.M_PAYMENTS: "תשלומים", sc.M_EXP: "סך הוצאות", sc.M_PROFIT: "רווח",
    sc.M_CONV: "המרה", sc.M_TASKS: "משימות", sc.S_TOTAL: "סה״כ",
}


def _short(label):
    raw = str(label)
    body = raw.strip()
    short = _SHORT.get(body, body.replace(" (₪)", "").replace("(₪)", "").strip())
    indent = len(raw) - len(raw.lstrip(" "))
    return short, indent


def _delta(key, cur, prev):
    """(text, direction) vs the previous reporting period. dir ∈ up|down|flat|none."""
    if prev is None or cur is None:
        return ("", "none")
    if key == "conversion":
        d = (cur - prev) * 100.0
        if abs(d) < 0.05:
            return ("ללא שינוי", "flat")
        return (("+" if d > 0 else "−") + "{:.1f} נק׳".format(abs(d)), "up" if d > 0 else "down")
    if not prev:
        return ("", "none")
    p = (cur - prev) / prev * 100.0
    if abs(p) < 0.5:
        return ("ללא שינוי", "flat")
    return (("+" if p > 0 else "−") + "{:.0f}%".format(abs(p)), "up" if p > 0 else "down")


def _root_vars(v):
    return (
        ":root{"
        "--ink:%s;--petrol:%s;--accent:%s;--text:%s;--muted:%s;--line:%s;--surface:%s;"
        "--bg:#F4F5F2;--good:%s;--bad:%s;--warn:%s;--pic-a:%s;--pic-b:%s;"
        "--s1:4px;--s2:8px;--s3:12px;--s4:16px;--s5:24px;--s6:40px;--s7:60px;--r:14px;"
        "--shadow:0 1px 2px rgba(22,44,58,.05),0 8px 24px rgba(22,44,58,.07);"
        "--shadow-h:0 2px 8px rgba(22,44,58,.10),0 18px 44px rgba(22,44,58,.14);"
        "--good-t:rgba(46,125,91,.12);--bad-t:rgba(178,58,58,.12);--accent-t:rgba(62,124,140,.12)"
        "}"
    ) % (v["primary_dk"], v["primary"], v["secondary"], v["text"], v["muted"], v["divider"],
         v["white"], v["good"], v["bad"], v["warn"], v["pic_a"], v["pic_b"])


_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{font-family:"Segoe UI","Assistant","Rubik",Arial,sans-serif;background:var(--bg);
  color:var(--text);direction:rtl;font-size:14px;line-height:1.5;
  -webkit-print-color-adjust:exact;print-color-adjust:exact;
  background-image:radial-gradient(rgba(31,58,77,.035) 1px,transparent 1px);background-size:22px 22px}
.tnum{font-variant-numeric:tabular-nums;font-feature-settings:"tnum"}
.wrap{max-width:1320px;margin:0 auto;padding:0 var(--s5)}

/* ---------- sticky header + tabs ---------- */
.appbar{position:sticky;top:0;z-index:30;color:#fff;
  background:linear-gradient(135deg,var(--petrol),var(--ink));
  box-shadow:0 10px 30px rgba(22,44,58,.18)}
.appbar::before{content:"";position:absolute;inset:0;opacity:.5;pointer-events:none;
  background-image:linear-gradient(rgba(255,255,255,.05) 1px,transparent 1px),
    linear-gradient(90deg,rgba(255,255,255,.05) 1px,transparent 1px);background-size:30px 30px}
.appbar .wrap{position:relative}
.bar-top{display:flex;align-items:center;justify-content:space-between;gap:var(--s4);
  padding:var(--s4) 0 var(--s3)}
.brand{display:flex;align-items:center;gap:var(--s3);min-width:0}
.brand .mark{flex:none;display:grid;place-items:center}
.brand b{font-weight:700;font-size:15px;letter-spacing:.2px;white-space:nowrap}
.brand .logochip{flex:none;display:inline-grid;place-items:center;background:#fff;
  border-radius:11px;padding:6px 13px;box-shadow:0 2px 10px rgba(0,0,0,.22)}
.brand .logochip img{display:block;height:43px;width:auto}
.ctx{display:flex;align-items:baseline;gap:var(--s3);min-width:0}
.ctx h1{font-size:19px;font-weight:600;letter-spacing:.2px;white-space:nowrap}
.chip{font-size:12px;font-weight:600;padding:3px 10px;border-radius:999px;
  background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.2)}
.gen{font-size:12px;color:rgba(255,255,255,.72);white-space:nowrap;font-variant-numeric:tabular-nums}
.tabs{display:flex;gap:2px;position:relative;overflow-x:auto;scrollbar-width:none}
.tabs::-webkit-scrollbar{display:none}
.tab{appearance:none;background:none;border:0;color:rgba(255,255,255,.66);cursor:pointer;
  font:inherit;font-weight:600;font-size:14px;padding:11px 18px 13px;white-space:nowrap;
  border-bottom:3px solid transparent;transition:color .2s,background .2s,border-color .2s;
  border-radius:8px 8px 0 0}
.tab:hover{color:#fff;background:rgba(255,255,255,.08)}
.tab.active{color:#fff;border-bottom-color:var(--accent)}
.tab:focus-visible{outline:2px solid var(--accent);outline-offset:-2px}

/* ---------- panels ---------- */
main{padding:var(--s7) 0 64px}
.panel{display:none}
[data-js] .panel{display:none}
.panel.active{display:block}
@media (prefers-reduced-motion:no-preference){
  .panel.active{animation:fade .4s ease both}
  .reveal{opacity:0;animation:up .55s cubic-bezier(.2,.7,.2,1) forwards}
  .kpis .reveal:nth-child(1){animation-delay:.02s}.kpis .reveal:nth-child(2){animation-delay:.07s}
  .kpis .reveal:nth-child(3){animation-delay:.12s}.kpis .reveal:nth-child(4){animation-delay:.17s}
  .kpis .reveal:nth-child(5){animation-delay:.22s}.kpis .reveal:nth-child(6){animation-delay:.27s}
  #panel-overview>.reveal:nth-child(2){animation-delay:.30s}
  #panel-overview>.reveal:nth-child(3){animation-delay:.38s}
}
@keyframes fade{from{opacity:0}to{opacity:1}}
@keyframes up{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
.flash{animation:flash 1s ease}
@keyframes flash{0%,100%{box-shadow:var(--shadow)}30%{box-shadow:0 0 0 3px var(--accent),var(--shadow-h)}}

/* ---------- KPI cards ---------- */
.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:var(--s4);margin-bottom:var(--s6)}
@media(max-width:1080px){.kpis{grid-template-columns:repeat(3,1fr)}}
@media(max-width:620px){.kpis{grid-template-columns:repeat(2,1fr)}}
.kpi{display:flex;flex-direction:column;gap:6px;text-align:right;cursor:pointer;
  background:var(--surface);border:1px solid var(--line);border-radius:var(--r);
  padding:var(--s5) var(--s4);box-shadow:var(--shadow);position:relative;overflow:hidden;
  font:inherit;color:inherit;transition:transform .18s,box-shadow .18s,border-color .18s}
.kpi::before{content:"";position:absolute;inset-block:0;inset-inline-start:0;width:4px;
  background:linear-gradient(var(--petrol),var(--accent));opacity:.85}
.kpi:hover{transform:translateY(-3px);box-shadow:var(--shadow-h);border-color:#cfd6db}
.kpi:active{transform:translateY(-1px)}
.kpi:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.kpi-val{font-size:30px;font-weight:700;line-height:1.05;letter-spacing:-.3px;
  font-variant-numeric:tabular-nums}
.kpi-lab{font-size:13px;color:var(--muted);font-weight:600}
.kpi-foot{display:flex;align-items:center;gap:var(--s2);margin-top:2px;flex-wrap:wrap}
.kpi-sub{font-size:12px;color:var(--muted);font-variant-numeric:tabular-nums}
.kpi .go{position:absolute;inset-inline-end:12px;top:12px;color:var(--muted);opacity:0;
  transition:opacity .18s,transform .18s;font-size:13px}
.kpi:hover .go{opacity:.8;transform:translateX(-2px)}

/* ---------- pills ---------- */
.pill{display:inline-flex;align-items:center;gap:4px;font-size:11.5px;font-weight:700;
  padding:2px 9px;border-radius:999px;line-height:1.6;font-variant-numeric:tabular-nums}
.pill.up{color:var(--good);background:var(--good-t)}
.pill.up::before{content:"▲";font-size:8px}
.pill.down{color:var(--bad);background:var(--bad-t)}
.pill.down::before{content:"▼";font-size:8px}
.pill.flat{color:var(--muted);background:rgba(92,102,112,.12)}
.pill.flat::before{content:"•"}
.pill.count{color:var(--accent);background:var(--accent-t)}

/* ---------- section + cards ---------- */
.sec{margin-bottom:var(--s6)}
.sec-h{display:flex;align-items:center;gap:var(--s3);margin-bottom:var(--s4)}
.sec-h h2{font-size:16px;font-weight:700;color:var(--ink);letter-spacing:.2px}
.sec-h::before{content:"";width:4px;height:18px;border-radius:3px;
  background:linear-gradient(var(--petrol),var(--accent))}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:var(--s5);align-items:start}
.grid-31{display:grid;grid-template-columns:7fr 5fr;gap:var(--s5);align-items:start}
@media(max-width:900px){.grid2,.grid-31{grid-template-columns:1fr}}
.card{background:var(--surface);border:1px solid var(--line);border-radius:var(--r);
  box-shadow:var(--shadow);padding:var(--s5);overflow:hidden}
.card>.ch{display:flex;align-items:center;justify-content:space-between;gap:var(--s2);
  margin-bottom:var(--s3)}
.card>.ch .t{font-size:14px;font-weight:700;color:var(--ink)}
.card.accent-a{border-top:3px solid var(--pic-a)}
.card.accent-b{border-top:3px solid var(--pic-b)}
.tagdot{display:inline-block;width:9px;height:9px;border-radius:3px;margin-inline-start:6px;vertical-align:middle}

/* ---------- tables ---------- */
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{padding:9px 10px;text-align:right;border-bottom:1px solid var(--line);font-variant-numeric:tabular-nums}
thead th{font-size:12px;font-weight:700;color:var(--muted);background:#fbfbfa;
  position:sticky;top:0;text-transform:none}
tbody tr:last-child td{border-bottom:0}
tbody tr:hover td{background:#f8faf9}
td.c,th.c{text-align:center}
tr.tot td{border-top:2px solid var(--line);font-weight:700;background:#fafbfb}
.metric{font-weight:600;color:var(--ink)}
.metric.sub{font-weight:500;color:var(--text)}
.good{color:var(--good)}.bad{color:var(--bad)}
.mat{display:inline-block;font-size:11px;font-weight:700;color:var(--warn);
  background:rgba(194,138,43,.12);padding:1px 8px;border-radius:999px}
.bar-cell{display:flex;align-items:center;gap:8px}
.bar-cell .track{flex:1;height:7px;border-radius:999px;background:var(--line);overflow:hidden}
.bar-cell .fill{height:100%;border-radius:999px;background:linear-gradient(90deg,var(--accent),var(--petrol))}

/* ---------- charts ---------- */
.legend{display:flex;gap:var(--s4);flex-wrap:wrap;margin-bottom:var(--s2)}
.legend span{font-size:12px;color:var(--muted);display:inline-flex;align-items:center;gap:6px}
.legend i{width:11px;height:11px;border-radius:3px}
.chart svg{display:block;width:100%;height:auto}
.grid-charts{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:var(--s5)}

/* ---------- bar-lists (breakdowns) ---------- */
.barlist{display:flex;flex-direction:column;gap:13px;padding-top:2px}
.bl-row{display:grid;grid-template-columns:minmax(58px,auto) 1fr minmax(58px,auto);align-items:center;gap:12px}
.bl-lab{font-size:13px;color:var(--ink);font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bl-track{height:10px;border-radius:999px;background:var(--line);overflow:hidden}
.bl-fill{display:block;height:100%;border-radius:999px;background:linear-gradient(90deg,var(--accent),var(--petrol))}
.bl-val{font-size:13px;color:var(--muted);text-align:left;white-space:nowrap;font-variant-numeric:tabular-nums}

/* ---------- donut + key-values ---------- */
.donut-wrap{display:flex;justify-content:center;padding:4px 0 12px}
.donut-wrap svg{max-width:168px}
.kv{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:10px 2px;
  border-top:1px solid var(--line);font-size:13px}
.kv span{color:var(--muted)}
.kv b{color:var(--ink);font-variant-numeric:tabular-nums;font-weight:700}

/* ---------- accordions ---------- */
.acc{background:var(--surface);border:1px solid var(--line);border-radius:12px;
  box-shadow:var(--shadow);margin-bottom:var(--s3);overflow:hidden;transition:box-shadow .18s}
.acc:hover{box-shadow:var(--shadow-h)}
.acc-head{display:flex;align-items:center;gap:var(--s3);width:100%;cursor:pointer;
  background:none;border:0;font:inherit;color:var(--ink);font-weight:700;
  padding:14px var(--s5);text-align:right;transition:background .15s}
.acc-head:hover{background:#f7f9f8}
.acc-head:focus-visible{outline:2px solid var(--accent);outline-offset:-2px}
.acc-head .chev{flex:none;color:var(--accent);font-size:12px;transition:transform .22s;display:inline-block}
.acc.open .acc-head .chev{transform:rotate(-90deg)}
.acc-head .t{flex:1}
.acc-body{max-height:0;overflow:hidden;transition:max-height .3s ease}
[data-js] .acc-body{max-height:0}
.acc-body .inner{padding:0 var(--s5) var(--s4);overflow:auto}
.acc-body table{font-size:12.5px}

/* ---------- hint / tooltip ---------- */
.hint{display:inline-grid;place-items:center;width:17px;height:17px;border-radius:50%;
  border:1px solid var(--line);color:var(--muted);font-size:11px;font-weight:700;
  cursor:help;margin-inline-start:6px;position:relative;vertical-align:middle;background:#fff}
.hint:hover,.hint:focus,.hint.show{color:#fff;background:var(--accent);border-color:var(--accent);outline:none}
.hint::after{content:attr(data-tip);position:absolute;z-index:40;top:130%;inset-inline-end:-4px;
  width:248px;background:var(--ink);color:#fff;font-size:12px;font-weight:500;line-height:1.5;
  padding:10px 12px;border-radius:10px;box-shadow:var(--shadow-h);opacity:0;visibility:hidden;
  transform:translateY(-4px);transition:opacity .16s,transform .16s;text-align:right;pointer-events:none}
.hint:hover::after,.hint:focus::after,.hint.show::after{opacity:1;visibility:visible;transform:none}

.note{font-size:12px;color:var(--muted);margin-top:var(--s3)}
.missing{list-style:none;display:flex;flex-direction:column;gap:6px}
.missing li{font-size:13px;color:var(--text);padding:9px 12px;background:var(--surface);
  border:1px solid var(--line);border-radius:10px;border-inline-start:3px solid var(--warn)}
.empty{color:var(--muted);font-size:13px;padding:var(--s4)}

@media print{.appbar{position:static}.panel{display:block!important}.tabs,.kpi .go{display:none}}
"""


_JS = """
(function(){
  var root=document.documentElement; root.setAttribute('data-js','1');
  function ready(fn){if(document.readyState!=='loading')fn();else document.addEventListener('DOMContentLoaded',fn);}
  ready(function(){
    var tabs=[].slice.call(document.querySelectorAll('.tab'));
    var panels=[].slice.call(document.querySelectorAll('.panel'));
    function activate(name,top){
      var ok=tabs.some(function(t){return t.dataset.tab===name;}); if(!ok)name='overview';
      tabs.forEach(function(t){var on=t.dataset.tab===name;t.classList.toggle('active',on);
        t.setAttribute('aria-selected',on?'true':'false');});
      panels.forEach(function(p){p.classList.toggle('active',p.dataset.panel===name);});
      try{history.replaceState(null,'','#'+name);}catch(e){location.hash=name;}
      if(top)window.scrollTo({top:0,behavior:'smooth'});
    }
    tabs.forEach(function(t){t.addEventListener('click',function(){activate(t.dataset.tab,false);});});
    document.querySelectorAll('[data-jump]').forEach(function(k){
      k.addEventListener('click',function(){
        var d=k.dataset.jump; activate(d,true);
        var sel='#panel-'+d+' [data-focus="'+(k.dataset.focus||'')+'"]';
        var sec=(k.dataset.focus&&document.querySelector(sel))||document.querySelector('#panel-'+d+' .sec');
        if(sec)setTimeout(function(){sec.classList.add('flash');
          setTimeout(function(){sec.classList.remove('flash');},1000);},220);
      });
    });
    document.querySelectorAll('.acc-head').forEach(function(h){
      h.addEventListener('click',function(){
        var acc=h.closest('.acc'); var body=acc.querySelector('.acc-body');
        var open=acc.classList.toggle('open'); h.setAttribute('aria-expanded',open?'true':'false');
        body.style.maxHeight=open?(body.scrollHeight+'px'):'0px';
      });
    });
    document.querySelectorAll('.hint').forEach(function(hn){
      hn.addEventListener('click',function(e){e.stopPropagation();
        var on=hn.classList.contains('show');
        document.querySelectorAll('.hint.show').forEach(function(x){x.classList.remove('show');});
        if(!on)hn.classList.add('show');});
    });
    document.addEventListener('click',function(){
      document.querySelectorAll('.hint.show').forEach(function(x){x.classList.remove('show');});});
    var init=(location.hash||'').replace('#','')||'overview';
    activate(init,false);
  });
})();
"""


# ============================================================== build
def build_html(report, data, out_path):
    v = sc.css_vars()
    months = report["months"]
    n = len(months)
    multi = n > 1
    vcols = vm.value_columns(report)
    prev = report.get("prev") or {}
    semc = {"good": v["good"], "bad": v["bad"], "warn": v["warn"], "primary": v["primary"], None: v["text"]}
    cs = vm.chart_series(report)
    short_m = [sc.HEB_MONTHS[int(m.split("-")[1]) - 1] for m in months]
    labels_on = n <= 6
    line_cols = [v["pic_a"], v["pic_b"], v["muted"]]

    def fmt(kind, val):
        return esc(vm.FMT[kind](val))

    # ---- reusable fragments
    def sec(title, body, hint=None, anchor=None):
        h = ' <span class="hint" tabindex="0" data-tip="{}">?</span>'.format(esc(hint)) if hint else ""
        aid = ' id="sec-{}"'.format(anchor) if anchor else ""
        focus = ' data-focus="{}"'.format(anchor) if anchor else ""
        return ('<section class="sec"{aid}{focus}><div class="sec-h"><h2>{t}{h}</h2></div>{b}</section>'
                ).format(aid=aid, focus=focus, t=esc(title), h=h, b=body)

    def card(body, cls="", head=None):
        ch = '<div class="ch"><div class="t">{}</div></div>'.format(esc(head)) if head else ""
        return '<div class="card {c}">{ch}{b}</div>'.format(c=cls, ch=ch, b=body)

    def block_table(block):
        head = "".join('<th class="c">{}</th>'.format(
            esc(sc.S_SUMMARY if k == "sum" else sc.heb_month(m))) for k, m in vcols)
        rows = ['<table><thead><tr><th>{}</th>{}</tr></thead><tbody>'.format(esc("מדד"), head)]
        for row in block.rows:
            col = semc.get(row.color)
            sty = ' style="color:{};font-weight:700"'.format(col) if row.color else ""
            short, indent = _short(row.label)
            cls = "metric sub" if indent else "metric"
            pad = ' style="padding-inline-start:{}px"'.format(10 + indent * 5) if indent else ""
            cells = "".join('<td class="c"{}>{}</td>'.format(
                sty, fmt(row.fmt, row.summary if k == "sum"
                         else (row.by_month.get(m) if row.by_month else None))) for k, m in vcols)
            rows.append('<tr><td class="{cl}"{pad}{sty}>{lab}</td>{c}</tr>'.format(
                cl=cls, pad=pad, sty=sty, lab=esc(short), c=cells))
        rows.append("</tbody></table>")
        return "".join(rows)

    def legend(items):
        return '<div class="legend">{}</div>'.format("".join(
            '<span><i style="background:{}"></i>{}</span>'.format(c, esc(l)) for c, l in items))

    def chartcard(title, svg_html, legend_html="", cls=""):
        ch = '<div class="ch"><div class="t">{}</div></div>'.format(esc(title)) if title else ""
        return '<div class="card chart {cl}">{ch}{lg}{s}</div>'.format(cl=cls, ch=ch, lg=legend_html, s=svg_html)

    def income_card(block, accent, dotcol, title):
        dot = '<span class="tagdot" style="background:{}"></span>'.format(dotcol)
        head = '<div class="ch"><div class="t">{}{}</div></div>'.format(dot, esc(title))
        return '<div class="card {acc}">{head}{body}</div>'.format(acc=accent, head=head, body=block_table(block))

    def barlist(pairs, money=False):
        fmtf = vm.fmt_money if money else vm.fmt_int
        mx = max((abs(val or 0) for _, val in pairs), default=0) or 1
        rows = "".join(
            '<div class="bl-row"><span class="bl-lab">{lab}</span>'
            '<span class="bl-track"><span class="bl-fill" style="width:{w}%"></span></span>'
            '<span class="bl-val">{val}</span></div>'.format(
                lab=esc(str(lab)), w=round(100 * abs(val or 0) / mx), val=esc(fmtf(val)))
            for lab, val in pairs)
        return '<div class="barlist">{}</div>'.format(rows)

    P = []
    P.append('<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8">')
    P.append('<meta name="viewport" content="width=device-width,initial-scale=1">')
    P.append('<title>{}</title>'.format(esc(report["period"])))
    P.append("<style>{}{}</style></head><body>".format(_root_vars(v), _CSS))

    # ---- sticky header + tabs
    z = report["zoom"]
    kind_lbl = "דשבורד רבעוני" if z == "quarterly" else ("דשבורד שנתי" if z == "annual" else "דשבורד חודשי")
    period_chip = ("Q{} {}".format(report["quarter"], report["year"]) if z == "quarterly"
                   else str(report["year"]) if z == "annual" else sc.heb_month(months[0]))
    today = datetime.date.today().strftime("%d/%m/%Y")
    # Brand mark: the studio logo on a white chip (logo-only — the wordmark carries the
    # name). Embedded as a constant base64 data-URI so the file stays self-contained and
    # byte-deterministic. Falls back to the line-mark + name text if the asset is absent.
    logo_uri = sc.logo_data_uri()
    if logo_uri:
        brand = '<span class="logochip"><img src="{uri}" alt="{alt}"></span>'.format(
            uri=logo_uri, alt=esc(sc.STUDIO_NAME))
    else:
        brand = ('<span class="mark"><svg viewBox="0 0 24 24" width="24" height="24" fill="none" '
                 'stroke="{a}" stroke-width="1.6"><rect x="3" y="3" width="18" height="18" rx="2"/>'
                 '<path d="M3 9h18M9 3v18" stroke-opacity=".45" stroke-width="1"/></svg></span>'
                 '<b>{studio}</b>').format(a=v["secondary"], studio=esc(sc.STUDIO_NAME))
    tabs_html = "".join(
        '<button class="tab" type="button" role="tab" data-tab="{id}">{lab}</button>'.format(id=tid, lab=esc(lab))
        for tid, lab in TABS)
    P.append('<header class="appbar"><div class="wrap">'
             '<div class="bar-top"><div class="brand">{brand}</div>'
             '<div class="ctx"><h1>{kind}</h1><span class="chip tnum">{chip}</span></div>'
             '<div class="gen">עודכן {date}</div></div>'
             '<nav class="tabs" role="tablist">{tabs}</nav></div></header>'.format(
                 brand=brand, kind=esc(kind_lbl),
                 chip=esc(period_chip), date=esc(today), tabs=tabs_html))

    P.append('<main class="wrap">')

    # ============================ OVERVIEW
    L, Pp, C = report["leads"], report["pipeline"], report["cash"]
    Pr, T, K = report["profit"], report["tasks"], report["cohort"]
    pcol = v["good"] if Pr["total"] >= 0 else v["bad"]

    def kpi(figkey, jump, label, valstr, color="", delta=None, sub=""):
        dt, dd = (delta or ("", "none"))
        pill = '<span class="pill {d}">{t}</span>'.format(d=dd, t=esc(dt)) if dt else ""
        subh = '<span class="kpi-sub">{}</span>'.format(esc(sub)) if sub else ""
        foot = '<div class="kpi-foot">{}{}</div>'.format(pill, subh) if (pill or subh) else ""
        sty = ' style="color:{}"'.format(color) if color else ""
        return ('<button class="kpi reveal" type="button" data-jump="{j}" data-focus="{f}" aria-label="{lab}">'
                '<span class="go">↗</span>'
                '<div class="kpi-val tnum" data-fig="{f}"{st}>{val}</div>'
                '<div class="kpi-lab">{lab}</div>{foot}</button>').format(
            j=jump, f=figkey, st=sty, val=esc(valstr), lab=esc(label), foot=foot)

    kpis = "".join([
        kpi("net", "finance", "הכנסה נטו", vm.fmt_money(C["net_total"]), delta=_delta("net", C["net_total"], prev.get("net"))),
        kpi("profit", "finance", "רווח", vm.fmt_money(Pr["total"]), color=pcol, delta=_delta("profit", Pr["total"], prev.get("profit"))),
        kpi("deals", "finance", "עסקאות", vm.fmt_int(Pp["deals_total"]), sub="שווי " + vm.fmt_money(Pp["value_total"]),
            delta=_delta("deals", Pp["deals_total"], prev.get("deals"))),
        kpi("leads", "leads", "פניות", vm.fmt_int(L["total"]), delta=_delta("leads", L["total"], prev.get("leads"))),
        kpi("conversion", "leads", "המרה", vm.fmt_pct(K["headline_rate"]), delta=_delta("conversion", K["headline_rate"], prev.get("conversion"))),
        kpi("tasks", "tasks", "משימות", vm.fmt_int(T["total"]), delta=_delta("tasks", T["total"], prev.get("tasks"))),
    ])

    a, b = vm.income_blocks(report)
    income = '<div class="grid2"><div class="reveal">{}</div><div class="reveal">{}</div></div>'.format(
        income_card(a, "accent-a", v["pic_a"], "פייפליין · נסגר"),
        income_card(b, "accent-b", v["pic_b"], "תזרים · נכנס"))

    if multi:
        trends = '<div class="grid-charts">{}{}{}</div>'.format(
            '<div class="reveal">{}</div>'.format(chartcard("מגמת הכנסה ושווי", svg.line_chart(
                [("net", cs["net"]), ("gross", cs["gross"]), ("value", cs["value"])],
                short_m, money=True, labels=labels_on, colors=line_cols),
                legend([(v["pic_a"], "נטו"), (v["pic_b"], "ברוטו"), (v["muted"], "שווי")]))),
            '<div class="reveal">{}</div>'.format(chartcard("מגמת פניות · עסקאות · משימות", svg.line_chart(
                [("leads", cs["leads"]), ("deals", cs["deals"]), ("tasks", cs["tasks"])],
                short_m, money=False, labels=labels_on, colors=line_cols),
                legend([(v["pic_a"], "פניות"), (v["pic_b"], "עסקאות"), (v["muted"], "משימות")]))),
            '<div class="reveal">{}</div>'.format(chartcard("מגמת רווח חודשי", svg.line_chart(
                [("profit", cs["profit"])], short_m, money=True, labels=labels_on, colors=[v["primary"]]))))
    else:
        comp = barlist([("שווי שנסגר", cs["value"][0]), ("הכנסה ברוטו", cs["gross"][0]),
                        ("הכנסה נטו", cs["net"][0])], money=True)
        trends = '<div class="reveal">{}</div>'.format(card(comp, head="הרכב ההכנסה החודש"))

    P.append('<section class="panel active" id="panel-overview" data-panel="overview">'
             '<div class="kpis">{kpis}</div>{income}<div style="height:var(--s6)"></div>{trends}</section>'.format(
                 kpis=kpis, income=income, trends=trends))

    # ============================ FINANCE
    eb = vm.expenses_block(report)
    fin = sec("הוצאות ורווח", card(block_table(eb)), anchor="net")
    if multi:
        fin += sec("פירוק וקצב", '<div class="grid2"><div>{}</div><div>{}</div></div>'.format(
            card(barlist(cs["cat"], money=True), head="הוצאות לפי קטגוריה"),
            chartcard("מגמת רווח חודשי", svg.line_chart(
                [("profit", cs["profit"])], short_m, money=True, labels=labels_on, colors=[v["primary"]]))))
    else:
        fin += sec("הוצאות לפי קטגוריה", card(barlist(cs["cat"], money=True)))
    P.append('<section class="panel" id="panel-finance" data-panel="finance">{}</section>'.format(fin))

    # ============================ LEADS & CONVERSION
    leadsec = sec("פניות לפי מקור", '<div class="grid-31"><div>{}</div><div>{}</div></div>'.format(
        card(block_table(vm.source_matrix(report))),
        card(barlist(cs["src"], money=False), head="התפלגות מקורות")))
    # cohort table
    crows = []
    vmax = max((r["n_conv"] for r in K["rows"]), default=0) or 1
    for row in K["rows"]:
        mat = '<span class="mat">{}</span>'.format(esc(sc.S_MATURING)) if row["maturing"] else ""
        w = round(100 * (row["n_conv"] or 0) / vmax)
        crows.append('<tr><td class="metric">{m}</td><td class="c">{nl}</td><td class="c">{nc}</td>'
                     '<td class="c">{r}</td><td><div class="bar-cell"><div class="track">'
                     '<div class="fill" style="width:{w}%"></div></div>{mat}</div></td></tr>'.format(
                         m=esc(sc.heb_month(row["month"])), nl=fmt("int", row["n_leads"]),
                         nc=fmt("int", row["n_conv"]), r=fmt("pct", row["rate"]), w=w, mat=mat))
    cohort_tbl = ('<table><thead><tr><th>חודש הגעה</th><th class="c">לידים</th><th class="c">נסגרו</th>'
                  '<th class="c">המרה</th><th>קצב</th></tr></thead><tbody>{rows}</tbody></table>').format(
        rows="".join(crows))
    ttc_kv = ""
    if K["ttc_median"] is not None:
        ttc_kv = ('<div class="kv"><span>זמן המרה <span class="hint" tabindex="0" data-tip="{tip}">?</span></span>'
                  '<b>{val}</b></div>').format(tip=esc(sc.S_TTC_CAVEAT), val=fmt("days", K["ttc_median"]))
    donut_card = card(
        '<div class="donut-wrap">{donut}</div>'
        '<div class="kv"><span>לידים בקוהורט</span><b>{hn}</b></div>'
        '<div class="kv"><span>נסגרו</span><b>{hc}</b></div>{ttc}'.format(
            donut=svg.donut(K["headline_rate"] or 0, color=v["secondary"], sub="מיוצב (בשלות)"),
            hc=fmt("int", K["headline_conv"]), hn=fmt("int", K["headline_n"]), ttc=ttc_kv),
        head="שיעור המרה מיוצב")
    leadsec += sec("המרת לידים", '<div class="grid-31"><div>{}</div><div>{}</div></div>'.format(
        card(cohort_tbl), donut_card),
        hint="הקוהורט מעוגן לחודש הגעת הליד; ~3 החודשים האחרונים עוד מבשילים, "
             "ולכן ההמרה המיוצבת מחושבת רק על קוהורטות בשלות.", anchor="conversion")
    leadsec += sec("שיוך לידים", card(block_table(vm.attribution_block(report, data))),
                   hint="כמה מהעסקאות שנסגרו מקורן בליד מתויג (חיבור לפי שם). היתר — הפניה / לקוח "
                        "חוזר / ישיר / טרם חלון המעקב — אינו חוסר נתונים אלא מדד הקשר.")
    P.append('<section class="panel" id="panel-leads" data-panel="leads">{}</section>'.format(leadsec))

    # ============================ TASKS
    tasksec = sec("משימות לפי מבצע", '<div class="grid-31"><div>{}</div><div>{}</div></div>'.format(
        card(block_table(vm.performer_matrix(report))),
        card(barlist(cs["perf"], money=False), head="התפלגות מבצעים")),
        anchor="tasks")
    P.append('<section class="panel" id="panel-tasks" data-panel="tasks">{}</section>'.format(tasksec))

    # ============================ DETAIL
    detail = vm.detail_tables(data, months)
    detail.append({"name": vm.EXPENSE_SPEC["name"], "cols": vm.EXPENSE_SPEC["cols"],
                   "rows": vm.expense_detail_rows(data, months)})
    accs = "".join(_accordion(spec) for spec in detail)
    miss = "".join("<li>{}</li>".format(esc(str(it))) for it in (report.get("missing") or ["אין חוסרים מהותיים"]))
    detail_html = accs + sec("חוסרים", '<ul class="missing">{}</ul>'.format(miss))
    P.append('<section class="panel" id="panel-detail" data-panel="detail">{}</section>'.format(detail_html))

    P.append('</main><script>{}</script></body></html>'.format(_JS))

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("".join(P))
    return out_path


# ============================================================== detail helpers
def _detail_cell(kind, vv):
    if kind == "date":
        if vv is None or (not hasattr(vv, "strftime") and pd.isna(vv)):
            return "—"
        ts = vv.to_pydatetime() if hasattr(vv, "to_pydatetime") else vv
        return ts.strftime("%d/%m/%Y")
    if kind == "money":
        return vm.fmt_money(None if (vv is None or pd.isna(vv)) else float(vv))
    if kind == "int":
        return vm.fmt_int(None if (vv is None or pd.isna(vv)) else float(vv))
    if vv is None or (isinstance(vv, float) and pd.isna(vv)):
        return "—"
    return esc(str(vv))


def _accordion(spec):
    cols, rows = spec["cols"], spec["rows"]
    head = "".join('<th class="{}">{}</th>'.format(
        "c" if kind != "text" else "", esc(h)) for _, h, kind, _ in cols)
    body = []
    for row in rows:
        cells = "".join('<td{}>{}</td>'.format(
            "" if kind == "text" else ' class="c"', _detail_cell(kind, row.get(key)))
            for key, _, kind, _ in cols)
        body.append("<tr>{}</tr>".format(cells))
    inner = ('<table><thead><tr>{}</tr></thead><tbody>{}</tbody></table>'.format(head, "".join(body))
             if rows else '<div class="empty">{}</div>'.format(esc(sc.S_NODATA)))
    return ('<div class="acc"><button class="acc-head" type="button" aria-expanded="false">'
            '<span class="chev">▾</span><span class="t">{name}</span>'
            '<span class="pill count">{ct}</span></button>'
            '<div class="acc-body"><div class="inner">{inner}</div></div></div>').format(
        name=esc(spec["name"]), ct=len(rows), inner=inner)
