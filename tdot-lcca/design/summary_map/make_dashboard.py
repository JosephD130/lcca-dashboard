"""Generates SummaryDashboard.dc.html: the whole Summary sheet drawn as a dashboard.

Every number and every chart comes from the populated MKL example run, so the artboard shows what
the sheet actually produces rather than invented values. The locator map is lifted verbatim from
the Dashboard artboard so the two stay identical.
"""
import re, warnings, os
from openpyxl import load_workbook
warnings.filterwarnings('ignore')

HERE = os.path.dirname(os.path.abspath(__file__))
RUN = '/home/user/lcca-dashboard/tdot-lcca/verification/examples/MKL_Runway2-20/MKL_LCCA_run.xlsx'
ALT = ['#2A78D6', '#EB6834', '#6DA7EC', '#F39C7A']      # the series colors the workbook charts use
CAT = ['#3E4C59', '#6B7A88', '#98A4AF', '#C4CDD5', '#E4E9ED']   # proposed: one ramp, dark to light
LAYER = ['#DCCBA0', '#C8A96E', '#3A3A3A', '#C6CBD0']
INK, MUTE, RULE = '#1D2733', '#595959', '#D0D4D9'
FF = 'Arial, Helvetica, sans-serif'

wb = load_workbook(RUN, data_only=True)
sm = wb['Summary']
cell = lambda r, c: sm.cell(r, c).value
# the Summary row map, the same constants build_summary.py lays the sheet out with
T0, T1, VER, CMP0, SEC_T = 12, 15, 17, 22, 90


def money(v, dp=0):
    if v is None: return ''
    return ('-$' if v < 0 else '$') + format(round(abs(v), dp), ',.%df' % dp)


def m1(v):
    return ('-$' if v < 0 else '$') + '%.1fM' % (abs(v) / 1e6)


# ---------------------------------------------------------------- data off the sheet
names = [cell(r, 8) for r in range(T0, T1 + 1)]
sheets = [cell(r, 7) for r in range(T0, T1 + 1)]
init = [cell(r, 10) for r in range(T0, T1 + 1)]
npw = [cell(r, 15) for r in range(T0, T1 + 1)]
vs = [cell(r, 16) for r in range(T0, T1 + 1)]
days = [cell(r, 17) for r in range(T0, T1 + 1)]
avail = [cell(r, 18) for r in range(T0, T1 + 1)]
euac = [cell(r, 13) for r in range(CMP0, CMP0 + 4)]
agency = [cell(r, 8) for r in range(CMP0, CMP0 + 4)]
agencyE = [cell(r, 9) for r in range(CMP0, CMP0 + 4)]
user = [cell(r, 10) for r in range(CMP0, CMP0 + 4)]
userE = [cell(r, 11) for r in range(CMP0, CMP0 + 4)]
pct = [cell(r, 15) for r in range(CMP0, CMP0 + 4)]
cats = [(cell(r, 23), [cell(r, 24 + i) for i in range(4)]) for r in range(5, 10)]
sens = [(cell(r, 23), [cell(r, 24 + i) for i in range(4)]) for r in range(12, 37)]
years = [(cell(r, 24), [cell(r, 25 + i) for i in range(4)], [cell(r, 29 + i) for i in range(4)],
          [cell(r, 33 + i) for i in range(4)]) for r in range(41, 72)]
sec_m = [(cell(r, 23), [cell(r, 24 + i) for i in range(4)]) for r in range(76, 80)]
sec_s = [(cell(r, 23), [cell(r, 24 + i) for i in range(4)]) for r in range(83, 87)]
secrow = [[cell(r, c) for c in range(7, 18)] for r in range(SEC_T + 3, SEC_T + 7)]
desc = [cell(r, 16) for r in range(SEC_T + 3, SEC_T + 7)]
desc_s = [cell(r, 17) for r in range(SEC_T + 3, SEC_T + 7)]
low = npw.index(min(npw))
gi = wb['General Information']
PERIOD, RATE = int(gi['D33'].value), gi['D34'].value
AREA, SHOULDER = gi['D26'].value, gi['D27'].value
YEAR0 = int(gi['D25'].value)
short = ['Alt 1', 'Alt 2', 'Alt 3', 'Alt 4']


# ---------------------------------------------------------------- svg helpers
def t(x, y, s, size=9, fill=MUTE, anchor='start', weight='normal'):
    return ('<text x="%.1f" y="%.1f" font-family="%s" font-size="%s" fill="%s" text-anchor="%s" '
            'font-weight="%s">%s</text>' % (x, y, FF, size, fill, anchor, weight, s))


def frame(W, H, L, R, T, B, ymin, ymax, ylab, xlab, ticks, fmt):
    """Axes, y ticks and both axis titles. Returns (svg parts, y-scale function)."""
    py = lambda v: (H - B) - (v - ymin) / (ymax - ymin) * (H - B - T)
    o = []
    for v in ticks:
        y = py(v)
        o.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#ECEFF2" stroke-width="1"/>'
                 % (L, y, W - R, y))
        o.append(t(L - 6, y + 3, fmt(v), 9, MUTE, 'end'))
    o.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#9AA0A6" stroke-width="1"/>' % (L, T, L, H - B))
    zero = py(0) if ymin < 0 < ymax else H - B
    o.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#9AA0A6" stroke-width="1"/>' % (L, zero, W - R, zero))
    o.append('<text x="%.1f" y="%.1f" font-family="%s" font-size="9" fill="%s" text-anchor="middle" '
             'transform="rotate(-90 %.1f %.1f)">%s</text>' % (11, (T + H - B) / 2, FF, MUTE, 11, (T + H - B) / 2, ylab))
    o.append(t((L + W - R) / 2, H - 3, xlab, 9, MUTE, 'middle'))
    return o, py, zero


def legend(x, y, labels, colors, size=9, gap=13):
    o = []
    for i, (lb, c) in enumerate(zip(labels, colors)):
        o.append('<rect x="%.1f" y="%.1f" width="8" height="8" fill="%s"/>' % (x, y + i * gap - 7, c))
        o.append(t(x + 12, y + i * gap, lb, size, MUTE))
    return o


def svg(W, H, body, label=''):
    return ('<svg viewBox="0 0 %d %d" width="100%%" height="%d" role="img" aria-label="%s" '
            'style="display: block;">%s</svg>' % (W, H, H, label, ''.join(body)))


# ---------------------------------------------------------------- the seven charts
W, H = 660, 208
L, R, T, B = 62, 128, 12, 30


def chart1():
    o, py, zero = frame(W, H, L, R, T, B, -2e6, 20e6, 'Present worth ($)', 'Alternative',
                        [-2e6, 0, 5e6, 10e6, 15e6, 20e6], lambda v: m1(v))
    span = (W - R - L) / 4.0
    for i in range(4):
        x = L + i * span + span * 0.22
        w = span * 0.56
        up = dn = 0.0
        for (lab, vals), c in zip(cats, CAT):
            v = vals[i]
            if v >= 0:
                y0, y1 = py(up + v), py(up); up += v
            else:
                y0, y1 = py(dn), py(dn + v); dn += v
            o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s"/>' % (x, y0, w, max(y1 - y0, 0.6), c))
        o.append(t(x + w / 2, H - 16, short[i], 9, INK if i == low else MUTE, 'middle',
                   'bold' if i == low else 'normal'))
    o += legend(W - R + 8, 22, [c[0] for c in cats], CAT)
    return svg(W, H, o, 'Present worth of each alternative broken into cost categories')


def chart2():
    lo = min(min(v[1]) for v in sens); hi = max(max(v[1]) for v in sens)
    o, py, _ = frame(W, H, L, R, T, B, 13.5e6, 20e6, 'Net present worth ($)', 'Discount rate (%)',
                     [14e6, 16e6, 18e6, 20e6], lambda v: m1(v))
    px = lambda r: L + (r - 2.0) / 6.0 * (W - R - L)
    for mark, lab in [(2.0, 'FAA 2%'), (3.0, 'TDOT 3%'), (7.0, 'pre-2022 7%')]:
        o.append('<line x1="%.1f" y1="%d" x2="%.1f" y2="%.1f" stroke="#C6CBD0" stroke-width="1" '
                 'stroke-dasharray="3 3"/>' % (px(mark), T, px(mark), H - B))
        o.append(t(px(mark) + 3, T + 9, lab, 9, '#8A8F98'))
    for i, c in enumerate(ALT):
        pts = ' '.join('%.1f,%.1f' % (px(r), py(v[i])) for r, v in sens)
        o.append('<polyline points="%s" fill="none" stroke="%s" stroke-width="2"/>' % (pts, c))
    for r in [2, 3, 4, 5, 6, 7, 8]:
        o.append(t(px(r), H - 16, '%d%%' % r, 9, MUTE, 'middle'))
    o += legend(W - R + 8, 22, names, ALT)
    return svg(W, H, o, 'Net present worth of each alternative against the discount rate')


def chart3():
    o, py, zero = frame(W, H, L, R, T, B, -4e6, 16e6, 'Spend, undiscounted ($)', 'Calendar year',
                        [-4e6, 0, 5e6, 10e6, 15e6], lambda v: m1(v))
    span = (W - R - L) / len(years)
    for k, (yr, spend, _, _) in enumerate(years):
        for i, c in enumerate(ALT):
            v = spend[i] or 0
            if v == 0: continue
            x = L + k * span + span * 0.12 + i * span * 0.19
            y0, y1 = (py(v), zero) if v > 0 else (zero, py(v))
            o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s"/>'
                     % (x, y0, span * 0.19, max(y1 - y0, 0.8), c))
        if k % 5 == 0:
            o.append(t(L + k * span + span / 2, H - 16, str(yr), 9, MUTE, 'middle'))
    o += legend(W - R + 8, 22, names, ALT)
    return svg(W, H, o, 'Undiscounted spend by calendar year for each alternative')


def chart4():
    o, py, _ = frame(W, H, L, R, T, B, 12e6, 20e6, 'Cumulative discounted cost ($)', 'Calendar year',
                     [12e6, 14e6, 16e6, 18e6, 20e6], lambda v: m1(v))
    span = (W - R - L) / (len(years) - 1)
    for i, c in enumerate(ALT):
        pts = ' '.join('%.1f,%.1f' % (L + k * span, py(row[2][i])) for k, row in enumerate(years))
        o.append('<polyline points="%s" fill="none" stroke="%s" stroke-width="2"/>' % (pts, c))
    for k, row in enumerate(years):
        if k % 5 == 0: o.append(t(L + k * span, H - 16, str(row[0]), 9, MUTE, 'middle'))
    o += legend(W - R + 8, 22, names, ALT)
    return svg(W, H, o, 'Cumulative discounted cost by calendar year for each alternative')


def chart5():
    o, py, zero = frame(W, H, L, R, T, B, 0, 30, 'Closure days', 'Calendar year',
                        [0, 10, 20, 30], lambda v: '%d' % v)
    span = (W - R - L) / len(years)
    for k, (yr, _, _, d) in enumerate(years):
        for i, c in enumerate(ALT):
            v = d[i] or 0
            if v == 0: continue
            x = L + k * span + span * 0.12 + i * span * 0.19
            o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s"/>'
                     % (x, py(v), span * 0.19, zero - py(v), c))
        if k % 5 == 0: o.append(t(L + k * span + span / 2, H - 16, str(yr), 9, MUTE, 'middle'))
    o += legend(W - R + 8, 22, names, ALT)
    return svg(W, H, o, 'Runway closure days by calendar year for each alternative')


def section_chart(rows):
    o, py, zero = frame(W, H, L, R, T, B, 0, 30, 'Thickness (inches)', 'Alternative',
                        [0, 10, 20, 30], lambda v: '%d' % v)
    span = (W - R - L) / 4.0
    for i in range(4):
        x = L + i * span + span * 0.25
        w = span * 0.5
        up = 0.0
        for (lab, vals), c in zip(rows, LAYER):
            v = vals[i] or 0
            if v <= 0: continue
            o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s"/>'
                     % (x, py(up + v), w, py(up) - py(up + v), c))
            up += v
        o.append(t(x + w / 2, py(up) - 5, '%.0f"' % up, 9, INK, 'middle', 'bold'))
        o.append(t(x + w / 2, H - 16, short[i], 9, MUTE, 'middle'))
    o += legend(W - R + 8, 22, [r[0] for r in rows], LAYER)
    return svg(W, H, o, 'Pavement section of each alternative, stacked by layer')


def benchmark():
    """The four alternatives' pavement unit cost against published Tennessee runway work."""
    X0, X1, TOP = 96, 566, 350.0
    px = lambda v: X0 + v / TOP * (X1 - X0)
    o = []
    o.append('<rect x="%.1f" y="26" width="%.1f" height="140" fill="#EDF1F5"/>' % (px(210), px(280) - px(210)))
    o.append(t((px(210) + px(280)) / 2, 20, '$210&#8211;280 all-in, TN 2024&#8211;25', 9, MUTE, 'middle'))
    order = sorted(range(4), key=lambda i: init[i])
    for k, i in enumerate(order):
        u = init[i] / 100083.0
        isl = i == low
        y = 40 + k * 30
        o.append('<rect x="%d" y="%d" width="%.1f" height="17" fill="%s"/>'
                 % (X0, y, px(u) - X0, '#C1440E' if isl else '#BBD4EE'))
        o.append(t(X0 - 8, y + 13, names[i], 9.5, INK if isl else MUTE, 'end', 'bold' if isl else 'normal'))
        o.append(t(px(u) + 6, y + 13, '$%d' % round(u), 9.5, INK, 'start', 'bold' if isl else 'normal'))
    for v in [0, 100, 200, 300]:
        o.append('<line x1="%.1f" y1="166" x2="%.1f" y2="171" stroke="#9AA0A6" stroke-width="1"/>' % (px(v), px(v)))
        o.append(t(px(v), 182, '$%d' % v, 9, MUTE, 'middle'))
    o.append('<line x1="%d" y1="166" x2="%d" y2="166" stroke="#9AA0A6" stroke-width="1"/>' % (X0, X1))
    o.append(t((X0 + X1) / 2, 199, 'Initial construction (with mobilization and engineering) / mainline S.Y.',
               9, MUTE, 'middle'))
    return svg(W, H, o, 'Unit cost of each alternative against published Tennessee runway costs')


# ---------------------------------------------------------------- the map, lifted from the Dashboard artboard
src = open(os.path.join(HERE, 'Dashboard.dc.html'), encoding='utf-8').read()
MAP = re.search(r'<svg viewBox="0 0 940 205".*?</svg>', src, re.S).group(0)
MAP = MAP.replace('width="583" height="127"', 'width="100%" height="auto" style="display: block;"', 1)


# ---------------------------------------------------------------- markup
def card(title, body, note=''):
    return ('<div style="border: 1px solid #D4D8DD; background: #FFFFFF;">'
            '<div style="font-size: 11px; font-weight: bold; color: #1D2733; padding: 7px 10px 5px;">%s</div>'
            '<div style="padding: 0 8px 8px;">%s</div>%s</div>'
            % (title, body, '<div style="font-size: 9.5px; color: #595959; padding: 0 10px 8px; line-height: 1.4;">%s</div>' % note if note else ''))


def kpi(label, value, sub, tone=''):
    box = ('border: 1px solid #E0C97A; background: #FFF3CD;' if tone == 'warn'
           else 'border: 1px solid #D4D8DD; background: #FFFFFF;')
    vc = '#7F6000' if tone == 'warn' else ('#2E8B1F' if tone == 'good' else INK)
    return ('<div style="%s padding: 8px 10px 9px;">'
            '<div style="font-size: 9.5px; color: #595959; letter-spacing: .04em; text-transform: uppercase;">%s</div>'
            '<div style="font-size: 18px; font-weight: bold; color: %s; line-height: 1.15; margin-top: 2px;">%s</div>'
            '<div style="font-size: 10px; color: #595959; margin-top: 1px;">%s</div></div>'
            % (box, label, vc, value, sub))


SENS_LOW = {r: min(range(4), key=lambda i: v[i]) for r, v in sens}
SENS_NOTE = ('%s stays lowest from %g to %g percent, so the recommendation does not turn on the rate.'
             % (names[low], sens[0][0], sens[-1][0])) if set(SENS_LOW.values()) == {low} else \
            ('The lowest-cost alternative changes within the range: %s.'
             % ', '.join('%g%% %s' % (r, names[i]) for r, i in sorted(SENS_LOW.items())))
SHLD_NOTE = ('The same quantities spread over mainline plus shoulder, which is the lower bound on each thickness.'
             if SHOULDER else 'Empty until a shoulder area is entered on General Information.')

maxn = max(npw)
rows = []
for i in range(4):
    isl = i == low
    bar = ('<div style="display: flex; align-items: center; gap: 8px;">'
           '<div style="background: %s; height: 12px; width: %.0f%%;"></div><span>%s</span></div>'
           % ('#2A78D6' if isl else '#BBD4EE', npw[i] / maxn * 68, money(npw[i])))
    rows.append('<tr%s><td>%s</td><td>%s</td><td>%s</td><td class="n">%s</td><td>%s</td>'
                '<td class="n">%s</td><td class="n">%s</td><td class="n">%.1f%%</td></tr>'
                % (' class="low"' if isl else '', sheets[i], names[i], cell(4 + i, 9), money(init[i]),
                   bar, '&ndash;' if isl else money(vs[i]), days[i], avail[i] * 100))

comp = []
for i in range(4):
    comp.append('<tr%s><td>%s</td><td class="n">%s</td><td class="n">%s</td><td class="n">%s</td>'
                '<td class="n">%s</td><td class="n">%s</td><td class="n">%s</td><td class="n">%s</td>'
                '<td class="n">%s</td><td>%s</td></tr>'
                % (' class="low"' if i == low else '', names[i], money(agency[i]), money(agencyE[i]),
                   money(user[i]), money(userE[i]), money(npw[i]), money(euac[i]),
                   '&ndash;' if i == low else money(vs[i]), '&ndash;' if i == low else '%.1f%%' % (pct[i] * 100),
                   'lowest' if i == low else ''))

secr = []
for i, row in enumerate(secrow):
    secr.append('<tr%s><td>%s</td><td class="n">%.1f</td><td class="n">%.1f</td><td class="n">%.1f</td>'
                '<td>%s</td><td class="n">%.1f</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                % (' class="low"' if i == low else '', row[0], row[1], row[2], row[3],
                   row[4], row[5], row[7], desc[i], desc_s[i]))

CHARTS = [
    ('1. Present worth by category (salvage below zero)', chart1(), ''),
    ('2. Net present worth vs. discount rate', chart2(), SENS_NOTE),
    ('3. Expenditure stream by calendar year (undiscounted)', chart3(), ''),
    ('4. Cumulative discounted cost', chart4(),
     'Where the lines cross is the year the higher first cost is paid back.'),
    ('5. Runway closure days by calendar year', chart5(), ''),
    ('6. Pavement section, mainline (inches, surface on top)', section_chart(sec_m), ''),
    ('7. Pavement section over mainline plus shoulder', section_chart(sec_s), SHLD_NOTE),
    ('Unit cost against recent Tennessee work', benchmark(),
     'The published range is all-in: pavement, lighting and grading. This workbook covers the pavement contract, '
     'so every marker should sit below the band. Far below it, or above it, is worth a second look at the quantities.'),
]

chart_cells = ''.join(card(ti, s, n) for ti, s, n in CHARTS)

HTML = '''<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <script src="./support.js"></script>
</head>
<body>
<x-dc>
<helmet>
  <style>
    body {{ margin: 0; background: #8A8F98; font-family: Arial, Helvetica, sans-serif; color: #1D2733; }}
    a {{ color: #2A78D6; }} a:hover {{ color: #1D5BA6; }}
    .sheet {{ background: #FFFFFF; box-sizing: border-box; padding: 18px 20px 22px; }}
    .band {{ font-size: 11px; font-weight: bold; letter-spacing: .05em; color: #FFFFFF; background: #1D2733;
             padding: 4px 10px; }}
    .sect {{ font-size: 11.5px; font-weight: bold; letter-spacing: .05em; color: #1D2733;
             border-bottom: 2px solid #1D2733; padding-bottom: 4px; }}
    .note {{ font-size: 10px; color: #595959; line-height: 1.45; }}
    table.grid {{ border-collapse: collapse; width: 100%; }}
    table.grid th {{ background: #D9D9D9; border: 1px solid #BFBFBF; font-size: 10px; padding: 4px 6px;
                     text-align: left; font-weight: bold; }}
    table.grid td {{ border: 1px solid #D0D4D9; font-size: 10.5px; padding: 3px 6px; }}
    td.n {{ text-align: right; }}
    tr.low td {{ background: #DDEBF7; color: #1F3864; font-weight: bold; }}
    .card {{ border: 1px solid #D4D8DD; background: #FFFFFF; }}
    .cardhd {{ background: #1D2733; color: #FFFFFF; font-size: 10.5px; font-weight: bold; letter-spacing: .05em;
               padding: 5px 9px; }}
    .kv {{ display: grid; grid-template-columns: 104px minmax(0, 1fr); gap: 3px 10px; font-size: 10.5px;
           padding: 8px 9px; }}
    .kv div.k {{ color: #595959; }}
    .kv div.v {{ font-weight: bold; }}
    .legend {{ display: flex; gap: 13px; align-items: center; font-size: 10px; padding: 6px 9px 8px;
               flex-wrap: wrap; }}
    .legend span.s {{ display: inline-block; width: 10px; height: 10px; margin-right: 5px; vertical-align: -1px; }}
    .flag {{ background: #FFF3CD; color: #7F6000; font-size: 10px; line-height: 1.45; padding: 7px 9px; }}
  </style>
</helmet>
<div class="sheet">

  <div style="display: flex; gap: 8px; align-items: center;">
    <span style="background: #1D2733; color: #FFFFFF; font-size: 11px; font-weight: bold; padding: 3px 9px;">&#9668; General Information</span>
    <span style="background: #2A78D6; color: #FFFFFF; font-size: 11px; font-weight: bold; padding: 3px 9px;">Instructions</span>
    <span style="font-size: 14px; font-weight: bold; margin-left: 6px; letter-spacing: .04em;">LCCA SUMMARY</span>
    <span style="flex: 1 1 auto;"></span>
    <span class="note">Everything on this sheet calculates from the alternative worksheets and updates by itself.</span>
  </div>
  <div style="font-size: 12.5px; font-weight: bold; margin: 8px 0 12px;">{ident}</div>

  <div style="display: flex; gap: 14px; align-items: flex-start;">
    <div style="flex: 1 1 0; min-width: 0;">

      <div style="display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px;">{kpis}</div>

      <div style="margin-top: 12px;">
        <div class="sect">RESULTS</div>
        <table class="grid" style="margin-top: 6px;">
          <tr><th style="width: 96px;">Worksheet</th><th style="width: 86px;">Alternative</th><th style="width: 66px;">Type</th>
              <th style="width: 96px;">Initial construction</th><th style="width: 220px;">Net present worth</th>
              <th style="width: 84px;">vs. lowest</th><th style="width: 56px;">Closure days</th>
              <th style="width: 62px;">Runway availability</th></tr>
          {rows}
        </table>
        <div style="font-size: 11.5px; font-weight: bold; background: #EAF2FB; padding: 6px 8px; margin-top: 7px;">{verdict}</div>
        <div class="note" style="margin-top: 4px;">{verdict2}</div>
      </div>
    </div>

    <div class="card" style="width: 372px; flex: none;">
      <div class="cardhd">PROJECT LOCATION</div>
      <div style="padding: 8px 9px 0;">{map}</div>
      <div class="legend">
        <span><span class="s" style="background: #7FA8CF;"></span>West</span>
        <span><span class="s" style="background: #2A78D6;"></span>Middle</span>
        <span><span class="s" style="background: #1D3F6E;"></span>East</span>
        <span><span class="s" style="background: #C1440E; border-radius: 50%;"></span>This project</span>
      </div>
      <div class="kv">
        <div class="k">Airport</div><div class="v">McKellar-Sipes Regional (MKL)</div>
        <div class="k">City / county</div><div class="v">Jackson / Madison County</div>
        <div class="k">Grand Division</div><div class="v">West</div>
        <div class="k">Coordinates</div><div class="v">35.5999&deg; N, 88.9156&deg; W</div>
        <div class="k">Elevation</div><div class="v">435 ft</div>
        <div class="k">Branch / project</div><div class="v">Runway 2-20, Reconstruction</div>
        <div class="k">Mainline area</div><div class="v">{area} S.Y. &nbsp;+&nbsp; {shoulder} S.Y. shoulder</div>
        <div class="k">Runway width</div><div class="v">150 ft <span style="font-weight: normal; color: #595959;">(input, used by the Google Earth export)</span></div>
      </div>
      <div class="flag">
        <b>Pricing basis:</b> every unit cost comes from the single Pay_Items <i>Unit Cost</i> column. The Middle,
        West and East average-cost columns are empty, so this West division project is priced on the same statewide
        numbers as any other.
      </div>
      <div style="padding: 8px 9px 10px;">
        <span style="background: #1D2733; color: #FFFFFF; font-size: 10.5px; font-weight: bold; padding: 4px 9px;">Export to Google Earth &#9658;</span>
        <div class="note" style="margin-top: 5px;">Writes a KML beside the workbook: the runway footprint of each alternative, a bar for its present worth, and every maintenance and rehabilitation event on the time slider.</div>
      </div>
    </div>
  </div>

  <div style="margin-top: 16px;">
    <div class="sect">COMPARISON &nbsp;<span style="font-weight: normal; color: #595959; letter-spacing: 0;">RealCost layout: agency cost, user cost, total; present worth and equivalent uniform annual cost</span></div>
    <table class="grid" style="margin-top: 6px;">
      <tr><th>Alternative</th><th>Agency cost PW</th><th>Agency EUAC</th><th>User cost PW (lost revenue)</th>
          <th>User EUAC</th><th>Total PW</th><th>Total EUAC</th><th>vs. lowest ($)</th><th>vs. lowest (%)</th><th>Lowest?</th></tr>
      {comp}
    </table>
    <div class="note" style="margin-top: 5px;">{compnote}</div>
  </div>

  <div style="margin-top: 16px;">
    <div class="sect">CHARTS</div>
    <div class="note" style="margin: 5px 0 9px;">{charthow}</div>
    <div style="display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px;">{charts}</div>
  </div>

  <div style="margin-top: 16px;">
    <div class="sect">PAVEMENT SECTION &nbsp;<span style="font-weight: normal; color: #595959; letter-spacing: 0;">read back from the quantities already entered; nothing here feeds a cost</span></div>
    <div class="note" style="margin: 5px 0 6px;">Asphalt unit weight, pcf: <span style="background: #D9D9D9; border: 1px solid #BFBFBF; padding: 1px 7px; font-weight: bold; color: #1D2733;">145</span> &nbsp; Only asphalt needs an assumption; 145 pcf reproduces the Murfreesboro section.</div>
    <table class="grid">
      <tr><th>Alternative</th><th>Surface course (in)</th><th>Aggregate base (in)</th><th>Subbase (in)</th>
          <th>Treated subgrade</th><th>Total section (in)</th><th>Excavation check</th>
          <th>Section from the quantities</th><th>Same quantities over mainline + shoulder</th></tr>
      {secr}
    </table>
    <div class="note" style="margin-top: 5px;">{secnote}</div>
  </div>

  <div class="note" style="margin-top: 14px; border-top: 1px solid #D4D8DD; padding-top: 8px;">
    Columns A to F are hidden: the Alternative Setup form still writes its own small table and &ldquo;Chart 1&rdquo;
    there, and this sheet repeats all of it with more detail. Unhide A:F to type an alternative description.
    The Method sheet carries every formula behind these numbers, and the three assumptions that are a decision
    rather than a calculation: the salvage basis, the lost-revenue treatment, and the statewide unit costs.
  </div>
</div>
</x-dc>
</body>
</html>
'''

unit = init[low] / AREA
margin = sorted(npw)[1] - npw[low]
close = margin / npw[low] * 100
mob = gi['D36'].value or 0
eng = gi['D37'].value or 0
kpis = ''.join([
    kpi('Lowest present worth', money(npw[low]), '%s, %s' % (names[low], desc[low])),
    kpi('Margin to next', money(margin),
        '%.1f%% %s' % (close, '&mdash; under 5%, treat the two as tied' if close < 5 else 'of the lowest present worth'),
        'warn' if close < 5 else ''),
    kpi('Equivalent annual cost', money(euac[low]), 'per year, %d years at %g%%' % (PERIOD, RATE)),
    kpi('Initial construction', money(init[low]), 'including %g%% mobilization and %g%% engineering' % (mob, eng)),
    kpi('Unit cost', '$%d / S.Y.' % round(unit), 'initial construction over the mainline area'),
    kpi('Rate sensitivity', 'holds %g&ndash;%g%%' % (sens[0][0], sens[-1][0]),
        'the lowest-cost alternative does not change', 'good') if set(SENS_LOW.values()) == {low} else
    kpi('Rate sensitivity', 'the winner changes', 'the lowest-cost alternative is not the same at every rate', 'warn'),
])

out = HTML.format(
    ident='McKellar-Sipes Regional Airport (MKL) &nbsp;|&nbsp; Jackson, West region &nbsp;|&nbsp; Runway 2-20 '
          '(Reconstruction) &nbsp;|&nbsp; construction %d &nbsp;|&nbsp; %d years at %g%%' % (YEAR0, PERIOD, RATE),
    area=format(round(AREA), ','), shoulder=format(round(SHOULDER or 0), ','),
    kpis=kpis, rows=''.join(rows), map=MAP, comp=''.join(comp), secr=''.join(secr), charts=chart_cells,
    verdict='Lowest present worth: %s &nbsp;|&nbsp; margin to next: %s &nbsp;|&nbsp; %g%% over %d years '
            '&nbsp;|&nbsp; lost revenue: %s' % (names[low], money(margin), RATE, PERIOD, gi['D38'].value or 'No'),
    verdict2='Same lowest-cost alternative at 2% (OMB A-94 real rate, FAA PGL 22-01) &nbsp;|&nbsp; same at 7% '
             '(the pre-2022 AIP rule).',
    compnote='Agency cost = initial construction + maintenance + rehabilitation + salvage. User cost = lost airport '
             'revenue during runway closures, the airport-side counterpart of the user delay cost in the '
             'Caltrans/FHWA RealCost layout. EUAC = PW &times; r(1+r)<sup>P</sup> / ((1+r)<sup>P</sup> &minus; 1) at '
             'the discount rate and analysis period on General Information.',
    charthow='How to read: 1 shows where each alternative&rsquo;s cost sits; 2 whether the lowest-cost alternative '
             'holds at other discount rates; 3 and 4 when the money is spent and when the higher first cost is paid '
             'back; 5 how often and how long the runway closes; 6 and 7 the section the quantities describe. The last '
             'panel is not on the sheet today: it puts all four unit costs on one scale against published Tennessee '
             'runway work, so a quantity that is out by an order of magnitude shows up here.',
    secnote='Thickness is not an input. Volume items give inches = 36 &times; C.Y. / mainline S.Y.; asphalt gives '
            'inches = 2,666.67 &times; tons / (pcf &times; mainline S.Y.). The last column is the section as the '
            'quantities describe it: paste it into the alternative description so the two can never disagree.',
)
open(os.path.join(HERE, 'SummaryDashboard.dc.html'), 'w', encoding='utf-8').write(out)
print('wrote SummaryDashboard.dc.html (%d bytes)' % len(out))
