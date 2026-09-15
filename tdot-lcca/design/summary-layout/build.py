"""Emit the four .dc.html artboards for the Summary layout review.

Everything here is measured, not guessed: BAND_WIDE is what the sheet can span once the empty
S:V columns join the band, BAND_NARROW is what it spans today, and every panel height is a whole
number of 20 px Excel rows. Panels sit in a 1 px hairline grid so the plots touch with no gutter.
"""
HEAD = open('_head.txt').read()
TAIL = '</x-dc>\n</body>\n</html>\n'

BAND_WIDE, BAND_NARROW = 1482, 1239      # the band as built, and as it was
RAIL_W = 364                             # P:R, the project-location rail
KEY_W = (BAND_WIDE - RAIL_W - 2) // 2    # 558: each key plot, less the two hairlines
THIRD = (BAND_WIDE - 2) // 3             # 493: a panel on either of the lower rows
ROW = 20                                 # one Excel row at 15 pt


def band(w):
    return ('<div style="position:relative;width:%dpx;background:var(--paper);'
            'font-family:\'Segoe UI\',Calibri,system-ui,sans-serif;">' % w)


HEADER = '''
  <div style="display:flex;align-items:center;gap:12px;background:var(--navy);height:53px;padding:0 16px;">
    <div style="display:flex;gap:6px;">
      <div style="color:#cdd6e2;font-size:11px;border:1px solid #3a4757;border-radius:4px;padding:4px 8px;">&#9664; General Information</div>
      <div style="color:#cdd6e2;font-size:11px;border:1px solid #3a4757;border-radius:4px;padding:4px 8px;">Instructions</div>
    </div>
    <div style="flex-grow:1;display:flex;align-items:baseline;gap:10px;">
      <div style="color:#fff;font-size:16px;font-weight:800;">LCCA Summary</div>
      <div style="color:#9fb0c4;font-size:11px;">Step 5 of 5</div>
    </div>
    <div style="display:flex;align-items:center;gap:7px;">
      <div style="background:var(--red);color:#fff;font-weight:800;font-size:12px;padding:3px 5px;border-radius:3px;">TN</div>
      <div style="line-height:1.05;"><div style="color:#fff;font-weight:700;font-size:11px;">TDOT</div>
        <div style="color:#9fb0c4;font-size:8px;">Department of Transportation</div></div>
    </div>
  </div>'''

CONTEXT = '''
  <div style="height:20px;display:flex;align-items:center;gap:9px;background:#fff;border-bottom:1px solid var(--line);padding:0 16px;font-size:10.5px;">
    <span style="font-weight:700;">Gatlinburg&#8211;Pigeon Forge (GKT)</span><span style="color:var(--grey);">&#183;</span>
    <span style="color:var(--muted);">Runway 10&#8211;28, Reconstruction</span><span style="color:var(--grey);">&#183;</span>
    <span style="color:var(--muted);">45,883 S.Y.</span><span style="color:var(--grey);">&#183;</span>
    <span style="color:var(--muted);">30 years at 3%</span><span style="color:var(--grey);">&#183;</span>
    <span style="color:var(--muted);">construction 2028</span>
  </div>'''

TILE = ('<div class="p"%s><div class="kicker">%s</div>'
        '<div style="%s">%s</div>'
        '<div class="sub">%s</div></div>')
TILES = [('Lowest present worth', '$5,852,247', 'Alternative 3', '', ''),
         ('Margin to next', '$183,970', '3.1% &#183; under 5%, treat as tied', '', ''),
         ('Equivalent annual', '$298,577', 'per year, 30 yr at 3%', '', ''),
         ('Initial construction', '$5,681,474', 'winner, mobilisation and engineering in', '', ''),
         ('Unit cost', '$124<span style="font-size:9.5px;font-weight:600;color:var(--muted);"> /S.Y.</span>',
          'over mainline area', '', ''),
         ('Rate sensitivity', 'Winner changes', 'not lowest at every rate',
          ' style="border-left:3px solid var(--amber);"',
          'font-size:13px;font-weight:800;margin-top:4px;color:var(--amber);')]


def tiles(n=6):
    cells = ''.join(TILE % (t[3], t[0], t[4] or 'font-size:16px;font-weight:800;margin-top:2px;',
                        t[1], t[2]) for t in TILES[:n])
    return ('\n  <div class="flush" style="grid-template-columns:repeat(%d,minmax(0,1fr));'
            'height:60px;border-top:none;">%s</div>' % (n, cells))


VERDICT = '''
  <div style="height:40px;display:flex;align-items:center;gap:10px;background:var(--green-soft);border-bottom:1px solid #bfe0cf;padding:0 16px;">
    <div style="display:flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;background:var(--green);flex-shrink:0;">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3"><path d="M20 6 9 17l-5-5"/></svg></div>
    <div style="flex-grow:1;font-size:12.5px;"><span style="font-weight:800;">Alternative 3 &#8212; lowest present worth</span>
      <span style="color:var(--muted);">&#160;&#183;&#160; 11&#8243; P-501 on 6&#8243; P-209 &#160;&#183;&#160; $183,970 ahead of the next option</span></div>
    <div style="font-size:16px;font-weight:800;color:var(--green);">$5,852,247</div>
  </div>'''


def panel(title, note, svg, key=False, style=''):
    tag = '<span class="key">Key</span>' if key else ''
    return ('<div class="p" style="display:flex;flex-direction:column;%s">'
            '<div style="display:flex;align-items:baseline;justify-content:space-between;gap:6px;">'
            '<div class="ctitle">%s</div>%s</div>'
            '<div class="sub" style="margin:1px 0 3px;">%s</div>%s</div>' % (style, title, tag, note, svg))


def svg(vb_w, vb_h, body, grow=True):
    return ('<svg width="100%%" viewBox="0 0 %d %d" preserveAspectRatio="none" '
            'style="%sdisplay:block;">%s</svg>' % (vb_w, vb_h, 'flex-grow:1;' if grow else '', body))


PER_M = 18.33      # px per $1M, from the axis: 0 at y=150, $6M at y=40


def chart_category(h=176):
    """1. Present worth by category. Drawn to the same numbers the table carries: gross stack up,
    salvage down, so the net is what the callout says."""
    # initial, maintenance, rehabilitation, lost revenue, salvage - $M, from the worked example
    alts = [(4.286, 1.095, 0.647, 0.190, 0.181, 'Alternative 1'),
            (5.919, 0.360, 0.332, 0.064, 0.610, 'Alternative 2'),
            (5.681, 0.360, 0.332, 0.064, 0.585, 'Alternative 3')]
    cols = ('#1f4e79', '#3e7cc0', '#6aa9df', '#d98a2b')
    out, x, base = [], 86, 150
    for i, (a, b, c, d, sal, label) in enumerate(alts):
        y = base
        for v, col in zip((a, b, c, d), cols):
            hh = max(round(v * PER_M), 2)
            y -= hh
            out.append('<rect x="%d" y="%d" width="86" height="%d" fill="%s"/>' % (x, y, hh, col))
        sh = max(round(sal * PER_M), 2)
        out.append('<rect x="%d" y="%d" width="86" height="%d" fill="#9aa7b4"/>' % (x, base, sh))
        win = i == 2
        out.append('<text x="%d" y="182" font-size="11" font-weight="%d" fill="%s" text-anchor="middle">%s</text>'
                   % (x + 43, 700 if win else 400, '#0f7b4f' if win else '#5b6675',
                      label + (' &#10003;' if win else '')))
        if win:
            out.append('<rect x="%d" y="%d" width="92" height="%d" fill="none" stroke="#0f7b4f" '
                       'stroke-width="2" rx="3"/>' % (x - 3, y - 4, base + sh - y + 8))
            out.append('<text x="%d" y="%d" font-size="10.5" font-weight="700" fill="#0f7b4f" '
                       'text-anchor="middle">$5.85M net</text>' % (x + 43, y - 9))
        x += 152
    grid = ''.join('<line x1="52" y1="%d" x2="530" y2="%d" stroke="#eef2f6"/>' % (y, y) for y in (40, 77, 114))
    axis = ('<line x1="52" y1="150" x2="530" y2="150" stroke="#cdd6e2"/>'
            '<text x="46" y="154" font-size="9" fill="#9aa7b4" text-anchor="end">0</text>'
            '<text x="46" y="80" font-size="9" fill="#9aa7b4" text-anchor="end">$4M</text>'
            '<text x="46" y="44" font-size="9" fill="#9aa7b4" text-anchor="end">$6M</text>')
    return svg(540, 190, grid + axis + ''.join(out))


def legend(items):
    sw = ''.join('<span style="display:inline-flex;align-items:center;gap:3px;">'
                 '<span style="width:8px;height:8px;background:%s;display:inline-block;"></span>%s</span>'
                 % (c, t) for c, t in items)
    return ('<div style="display:flex;gap:9px;flex-wrap:wrap;font-size:9px;color:var(--muted);'
            'margin-top:2px;">%s</div>' % sw)


CAT_LEGEND = legend([('#1f4e79', 'Initial'), ('#3e7cc0', 'Maintenance'), ('#6aa9df', 'Rehab'),
                     ('#d98a2b', 'Lost revenue'), ('#9aa7b4', 'Salvage, below the line')])
ALT_LEGEND = legend([('#3e7cc0', 'Alternative 1'), ('#1c2634', 'Alternative 2'), ('#0f7b4f', 'Alternative 3')])


def chart_rate(h=176):
    """2. Net present worth vs. discount rate. Axis runs $5.4M to $6.4M, 100 px per $M, so the
    curves start where the sensitivity block actually puts them at 2%."""
    top, per = 40, 100.0                      # y=40 is $6.4M
    y = lambda v: round(top + (6.4 - v) * per)
    ser = [('#3e7cc0', 6.311, 5.45), ('#1c2634', 6.048, 5.70), ('#0f7b4f', 5.844, 5.60)]
    lines, xs = '', [68, 143, 218, 293, 368, 443, 512]
    for col, v0, v1 in ser:
        pts = ' '.join('%d,%d' % (x, y(v0 + (v1 - v0) * (i / 6.0) ** 1.08))
                       for i, x in enumerate(xs))
        lines += ('<polyline points="%s" fill="none" stroke="%s" stroke-width="%s"/>'
                  % (pts, col, '2.6' if col == '#0f7b4f' else '2'))
    cross = ('<line x1="404" y1="28" x2="404" y2="150" stroke="#d98a2b" stroke-width="1.4" '
             'stroke-dasharray="4 3"/>'
             '<text x="398" y="38" font-size="9.5" font-weight="700" fill="#d98a2b" '
             'text-anchor="end">Alternative 1 takes over</text>')
    grid = ''.join('<line x1="68" y1="%d" x2="520" y2="%d" stroke="#eef2f6"/>' % (v, v)
                   for v in (y(6.2), y(6.0), y(5.8), y(5.6)))
    # one label a percentage point, not one every quarter point: twenty-five rotated labels are
    # unreadable and cost the plot its bottom fifth
    ticks = ''.join('<text x="%d" y="166" font-size="9" fill="#9aa7b4" text-anchor="middle">%d%%</text>'
                    % (68 + i * 75.3, 2 + i) for i in range(7))
    ax = ('<line x1="68" y1="150" x2="520" y2="150" stroke="#cdd6e2"/>'
          '<text x="62" y="%d" font-size="9" fill="#9aa7b4" text-anchor="end">$6.4M</text>'
          '<text x="62" y="%d" font-size="9" fill="#9aa7b4" text-anchor="end">$5.6M</text>'
          % (y(6.4) + 4, y(5.6) + 4)) + ticks
    return svg(540, 190, grid + cross + lines + ax)


YEARS = 31


def year_bars(spikes, color='#3e7cc0', y0=None):
    """A bar per year of the analysis period. `spikes` maps year -> height in px."""
    out = ''.join('<rect x="%d" y="%d" width="7" height="%d" fill="%s"/>'
                  % (26 + yr * 9.4, 100 - h, h, color)
                  for yr, h in sorted(spikes.items()))
    ticks = ''.join('<text x="%d" y="112" font-size="8" fill="#9aa7b4" text-anchor="middle">%d</text>'
                    % (26 + yr * 9.4 + 3, 2028 + yr) for yr in (0, 5, 10, 15, 20, 25, 30))
    return svg(320, 118, '<line x1="20" y1="100" x2="316" y2="100" stroke="#cdd6e2"/>' + out + ticks)


SPEND = {0: 86, 8: 9, 12: 14, 16: 9, 20: 26, 24: 9, 28: 14}     # construction, then maintenance and rehab
CLOSURE = {0: 62, 12: 14, 20: 22, 28: 14}                        # the years the runway is actually shut


def chart_cumulative():
    """4. Cumulative discounted cost. Each line has to END in the order the table gives:
    Alternative 3 lowest, then Alternative 1, then Alternative 2."""
    ser = [('#3e7cc0', 6.036), ('#1c2634', 6.065), ('#0f7b4f', 5.852)]
    init = {'#3e7cc0': 4.29, '#1c2634': 5.92, '#0f7b4f': 5.68}
    y = lambda v: round(100 - v * 13.0)
    lines = ''
    for col, end in ser:
        a = init[col]
        pts = ' '.join('%d,%d' % (26 + i * 9.1, y(a + (end - a) * (i / 30.0) ** 0.55))
                       for i in range(0, 31, 3))
        lines += ('<polyline points="26,100 %s" fill="none" stroke="%s" stroke-width="%s"/>'
                  % (pts, col, '2.4' if col == '#0f7b4f' else '1.8'))
    return svg(320, 118, '<line x1="20" y1="100" x2="316" y2="100" stroke="#cdd6e2"/>' + lines)


def chart_benchmark():
    rows = [('TN 2024&#8211;25 band', 210, 280, '#9aa7b4'), ('Alternative 1', 93, 93, '#3e7cc0'),
            ('Alternative 2', 129, 129, '#3e7cc0'), ('Alternative 3', 124, 124, '#0f7b4f')]
    out, y = [], 16
    for name, a, b, c in rows:
        x0, x1 = 96 + a * 0.62, 96 + b * 0.62
        out.append('<text x="90" y="%d" font-size="9" fill="#5b6675" text-anchor="end">%s</text>' % (y + 8, name))
        out.append('<rect x="%.0f" y="%d" width="%.0f" height="11" fill="%s" rx="1"/>'
                   % (96, y, max(x1 - 96, 4) if a == b else x1 - x0, c) if a == b else
                   '<rect x="%.0f" y="%d" width="%.0f" height="11" fill="%s" rx="1"/>' % (x0, y, x1 - x0, c))
        lab = '$%d&#8211;$%d' % (a, b) if a != b else '$%d' % b
        out.append('<text x="%.0f" y="%d" font-size="9" font-weight="700" fill="#3a4757">%s</text>'
                   % (x1 + 5, y + 9, lab))
        y += 24
    return svg(320, 118, ''.join(out))


def chart_section(tag):
    cols = [(0, 34, '#1c2634'), (1, 26, '#cdb894'), (2, 22, '#cdb894')]
    out, x = [], 40
    for i in range(3):
        y = 100
        for hh, c in ((26 if i == 0 else 30, '#1c2634' if i == 0 else '#cdd6e2'),
                      (22, '#cdb894'), (20, '#e3d7bd')):
            y -= hh
            out.append('<rect x="%d" y="%d" width="52" height="%d" fill="%s"/>' % (x, y, hh, c))
        out.append('<text x="%d" y="112" font-size="9" fill="#5b6675" text-anchor="middle">Alt %d</text>'
                   % (x + 26, i + 1))
        x += 82
    return svg(320, 118, '<line x1="24" y1="100" x2="300" y2="100" stroke="#cdd6e2"/>' + ''.join(out))


TN = ('<path d="M24 26 L20 40 L26 52 L22 66 L28 76 L120 76 L200 74 L250 72 L300 66 L330 56 '
      'L348 44 L336 34 L310 24 L250 22 L160 23 L80 24 Z" fill="#eaf0f6" stroke="#9fb6d0" stroke-width="1.2"/>')
TN_DOTS = ('<g fill="#9fb6d0"><circle cx="56" cy="58" r="2"/><circle cx="92" cy="40" r="2"/>'
           '<circle cx="140" cy="62" r="2"/><circle cx="178" cy="38" r="2"/><circle cx="214" cy="60" r="2"/>'
           '<circle cx="244" cy="36" r="2"/><circle cx="308" cy="40" r="2"/></g>'
           '<circle cx="276" cy="49" r="5.5" fill="#c8102e"/>'
           '<circle cx="276" cy="49" r="10" fill="none" stroke="#c8102e" stroke-width="1.2" opacity=".45"/>'
           '<text x="276" y="38" font-size="10" font-weight="800" fill="#c8102e" text-anchor="middle">GKT</text>')
DIVISIONS = ('<line x1="118" y1="24" x2="122" y2="76" stroke="#c6d2e0" stroke-width="1"/>'
             '<line x1="232" y1="22" x2="236" y2="73" stroke="#c6d2e0" stroke-width="1"/>'
             '<text x="70" y="90" font-size="8.5" fill="#9aa7b4" text-anchor="middle">West</text>'
             '<text x="176" y="90" font-size="8.5" fill="#9aa7b4" text-anchor="middle">Middle</text>'
             '<text x="292" y="90" font-size="8.5" font-weight="700" fill="#5b6675" text-anchor="middle">East</text>')

FACTS = [('City / county', 'Sevierville / Sevier'), ('Grand Division', 'East'),
         ('Coordinates', '35.8578&#176; N, 83.5287&#176; W'), ('Elevation', '1,014 ft'),
         ('Mainline area', '45,883 S.Y.')]


def fact_rows(extra_input=True):
    out = ''.join('<div style="display:flex;justify-content:space-between;gap:8px;font-size:10.5px;'
                  'padding:2px 0;border-bottom:1px solid #f0f3f7;">'
                  '<span style="color:var(--muted);">%s</span><span style="font-weight:600;">%s</span></div>'
                  % kv for kv in FACTS)
    if extra_input:
        out += ('<div style="display:flex;justify-content:space-between;align-items:center;gap:8px;'
                'font-size:10.5px;padding:3px 0;">'
                '<span style="color:var(--muted);">Runway width</span>'
                '<span style="background:#f2f4f7;border:1px solid #bfbfbf;padding:1px 10px;font-weight:700;">75 ft</span></div>')
    return out


def location_rail():
    """The right-hand panel of the chart row. It has to fit 240 px: 16 for the title, 84 for the
    map, 107 for six facts, against a 228 px content box. The runway width is last and must stay
    visible &#8212; it is the one editable cell in the block and the export macro reads it."""
    facts = ''.join('<div style="display:flex;justify-content:space-between;gap:8px;font-size:10.5px;'
                    'line-height:1.35;padding:1px 0;border-bottom:1px solid #f0f3f7;">'
                    '<span style="color:var(--muted);">%s</span><span style="font-weight:600;">%s</span></div>'
                    % kv for kv in FACTS)
    facts += ('<div style="display:flex;justify-content:space-between;align-items:center;gap:8px;'
              'font-size:10.5px;padding:2px 0;">'
              '<span style="color:var(--muted);">Runway width</span>'
              '<span style="background:#f2f4f7;border:1px solid #bfbfbf;padding:0 9px;font-weight:700;">75 ft</span></div>')
    return ('<div class="p" style="display:flex;flex-direction:column;background:var(--panel);">'
            '<div style="display:flex;align-items:baseline;justify-content:space-between;">'
            '<div class="ctitle">Project location</div>'
            '<span style="font-size:8.5px;font-weight:800;text-transform:uppercase;letter-spacing:.06em;'
            'color:var(--blue-dk);background:var(--blue-soft);padding:1px 6px;border-radius:20px;">Context</span></div>'
            '<svg width="100%" height="84" viewBox="0 0 360 96" preserveAspectRatio="xMidYMid meet" '
            'style="display:block;margin-top:2px;">' + TN + DIVISIONS + TN_DOTS + '</svg>'
            '<div style="margin-top:2px;">' + facts + '</div></div>')


def location_strip():
    """The compact form for the frozen header band: map inline, facts in two columns."""
    half = FACTS[:3], FACTS[3:]
    cols = ''
    for group in half:
        cols += '<div style="display:flex;flex-direction:column;gap:1px;min-width:0;">'
        for k, v in group:
            cols += ('<div style="display:flex;justify-content:space-between;gap:6px;font-size:10px;">'
                     '<span style="color:var(--muted);">%s</span><span style="font-weight:600;">%s</span></div>' % (k, v))
        cols += '</div>'
    return ('<div class="p" style="display:flex;align-items:center;gap:10px;background:#fbfcfd;">'
            '<div style="width:150px;flex-shrink:0;">' + svg(360, 96, TN + TN_DOTS, grow=False) + '</div>'
            '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;flex-grow:1;min-width:0;">'
            + cols + '</div>'
            '<div style="text-align:right;flex-shrink:0;"><div class="kicker">Runway width</div>'
            '<div style="background:#f2f4f7;border:1px solid #bfbfbf;padding:1px 10px;font-weight:700;'
            'font-size:11px;margin-top:2px;">75 ft</div></div></div>')


def sectitle(t, note=''):
    return ('<div style="display:flex;align-items:baseline;gap:10px;background:var(--paper);padding:7px 16px 5px;">'
            '<div style="font-size:10.5px;font-weight:800;color:var(--muted);text-transform:uppercase;'
            'letter-spacing:.1em;">%s</div><div class="sub">%s</div></div>' % (t, note))


RESULTS = [('Alt 1 &#183; New HMA', 'HMA-New', '$4,285,512', '$1,094,529', '$647,019', '$189,698',
            '&#8722;$180,541', '$6,036,217', '+$183,970', '56 d'),
           ('Alt 2 &#183; New PCC', 'PCC-New', '$5,918,906', '$359,652', '$332,277', '$64,017',
            '&#8722;$609,628', '$6,065,225', '+$212,978', '24 d'),
           ('Alt 3 &#183; New PCC &#10003;', 'PCC-New', '$5,681,474', '$359,652', '$332,277', '$64,017',
            '&#8722;$585,173', '$5,852,247', '&#8212;', '24 d')]


def tables():
    head = ('<thead><tr style="background:var(--navy);"><th>Alternative</th><th>Type</th><th>Initial</th>'
            '<th>Maint PW</th><th>Rehab PW</th><th>Lost rev</th><th>Salvage</th><th>Net PW</th>'
            '<th>vs. lowest</th><th>Closure</th></tr></thead>')
    body = ''
    for i, r in enumerate(RESULTS):
        win = i == 2
        tr = '<tr style="%s">' % ('background:var(--green-soft);' if win else 'border-bottom:1px solid var(--line);')
        cells = ''
        for j, v in enumerate(r):
            st = ''
            if j == 1: st = 'color:var(--muted);'
            if j == 7: st = 'font-weight:800;color:var(--green);' if win else 'font-weight:700;'
            if j == 8 and not win: st = 'color:var(--amber);'
            if j == 0 and win: st = 'font-weight:800;'
            cells += '<td style="%s">%s</td>' % (st, v)
        body += tr + cells + '</tr>'
    return (sectitle('The numbers &#8212; every alternative, in full',
                     'Everything above is a picture of this table.')
            + '<div style="background:#fff;border-top:1px solid var(--line);border-bottom:1px solid var(--line);">'
            '<table style="width:100%;border-collapse:collapse;">' + head + '<tbody>' + body + '</tbody></table></div>'
            + sectitle('Comparison', 'RealCost layout: agency cost, user cost, present worth and equivalent uniform annual cost.')
            + '<div style="background:#fff;border-top:1px solid var(--line);border-bottom:1px solid var(--line);height:74px;"></div>'
            + sectitle('Pavement section', 'Read back from the quantities already entered; nothing here changes a cost.')
            + '<div style="background:#fff;border-top:1px solid var(--line);border-bottom:1px solid var(--line);height:74px;"></div>')


def freeze_line(label='Frozen here &#8212; everything above stays put while the rest scrolls'):
    return ('<div style="display:flex;align-items:center;gap:8px;background:var(--paper);padding:3px 16px;">'
            '<div style="flex-grow:1;border-top:2px solid var(--navy);"></div>'
            '<div style="font-size:8.5px;font-weight:700;color:var(--navy);text-transform:uppercase;'
            'letter-spacing:.07em;">%s</div>'
            '<div style="flex-grow:1;border-top:2px solid var(--navy);"></div></div>' % label)


def fold(y):
    return ('<div class="fold" style="top:{{foldY}}px;"></div>'
            '<div class="foldtag" style="top:{{foldY}}px;">{{foldLabel}}</div>')


LOGIC = '''<script data-dc-script data-props='{"screen":{"editor":"enum",
  "options":["15 in FHD at 125% (624 px)","15 in FHD at 150% (480 px)","15 in FHD at 100% (840 px)"],
  "default":"15 in FHD at 125% (624 px)","section":"Screen"}}'>
class Component extends DCLogic {
  renderVals() {
    const s = this.props.screen ?? '15 in FHD at 125% (624 px)';
    const px = Number((s.match(/\\((\\d+) px\\)/) || [0, 624])[1]);
    return { foldY: px, foldLabel: 'bottom of a ' + s.replace(/ \\(.*\\)/, '') + ' screen' };
  }
}
</script>'''


# Titles say what the chart is; the axes should not say it again. Chart 1 carried an
# "Alternative" category-axis title over labels that already read "Alternative 1", and a
# "Present worth ($)" value-axis title under a chart called "Present worth by category". Both
# come off, the second line of chart 2's title moves into its sub line, and the plot takes the room.
C1 = lambda: panel('1. Present worth by category',
                   'Where the money goes. Salvage sits below the line.',
                   chart_category() + CAT_LEGEND, key=True)
C2 = lambda: panel('2. Net present worth vs. discount rate',
                   'Does the winner hold? TDOT 3%, FAA 2%, pre-2022 rule 7%.',
                   chart_rate() + ALT_LEGEND, key=True)
SUP = [lambda: panel('3. Expenditure by year', 'Undiscounted, as it falls.', year_bars(SPEND)),
       lambda: panel('4. Cumulative discounted cost', 'Each line ends at its net present worth.',
                     chart_cumulative() + ALT_LEGEND),
       lambda: panel('5. Runway closure days', 'The years the runway is shut, and for how long.',
                     year_bars(CLOSURE, '#d98a2b'))]
SEC = [lambda: panel('6. Section, mainline', 'Inches, read back from the quantities.', chart_section('m')),
       lambda: panel('7. Section with shoulder', 'Blank when the project has no paved shoulder.',
                     chart_section('s')),
       lambda: panel('8. Unit cost vs. published TN work', 'Against the 2024&#8211;25 all-in band.',
                     chart_benchmark())]


def row(cols, height, cells):
    return ('<div class="flush" style="grid-template-columns:%s;height:%dpx;">%s</div>'
            % (cols, height, ''.join(c() for c in cells)))


def write(name, w, body, logic=True):
    open(name, 'w').write(HEAD + band(w) + body + ('\n' + LOGIC if logic else '') + '\n</div>\n' + TAIL)
    print('%-22s %d px wide' % (name, w))


# ---------------------------------------------------------------- Main: the refinement
MAIN_BODY = (fold(624) + HEADER + CONTEXT + tiles() + VERDICT + freeze_line()
             + row('%dpx %dpx %dpx' % (KEY_W, KEY_W, RAIL_W), 13 * ROW, [C1, C2, location_rail])
             + row('repeat(3,minmax(0,1fr))', 9 * ROW, SUP)
             + row('repeat(3,minmax(0,1fr))', 9 * ROW, SEC)
             + tables())
write('Main.dc.html', BAND_WIDE, MAIN_BODY)

# ---------------------------------------------------------------- Option B: inside today's band
write('OptionB.dc.html', BAND_NARROW,
      fold(624) + HEADER + CONTEXT + tiles() + VERDICT
      + freeze_line()
      + row('433px 434px 371px', 12 * ROW, [C1, C2, location_rail])
      + row('repeat(4,minmax(0,1fr))', 9 * ROW, SUP)
      + sectitle('The numbers &#8212; every alternative, in full', 'Tables continue below.'))

# ---------------------------------------------------------------- Option C: location in the frozen band
write('OptionC.dc.html', BAND_NARROW,
      fold(624) + HEADER + CONTEXT
      + '<div class="flush" style="grid-template-columns:1fr;height:60px;">' + location_strip() + '</div>'
      + tiles() + VERDICT + freeze_line()
      + row('repeat(2,minmax(0,1fr))', 12 * ROW, [C1, C2])
      + row('repeat(4,minmax(0,1fr))', 9 * ROW, SUP)
      + sectitle('The numbers &#8212; every alternative, in full', 'Tables continue below.'))


# ---------------------------------------------------------------- Grid: the sizing and spacing review
def tbl(headers, rows, widths=None, right_from=1):
    cols = ''.join('<col style="width:%s"/>' % w for w in widths) if widths else ''
    th = ''.join('<th style="text-align:%s">%s</th>' % ('left' if i < right_from else 'right', h)
                 for i, h in enumerate(headers))
    tr = ''
    for r in rows:
        tds = ''.join('<td style="text-align:%s;%s">%s</td>'
                      % ('left' if i < right_from else 'right',
                         'font-weight:700;' if i == len(r) - 1 and isinstance(r[-1], str)
                         and r[-1].startswith('&#9733;') else '', c)
                      for i, c in enumerate(r))
        tr += '<tr style="border-bottom:1px solid var(--line);">%s</tr>' % tds
    return ('<table style="width:100%%;border-collapse:collapse;background:#fff;">%s'
            '<thead><tr style="background:var(--navy);">%s</tr></thead><tbody>%s</tbody></table>'
            % (cols, th, tr))


def stack(items, scale=0.42, total_label=''):
    """A measured vertical stack: each band drawn to scale with its pixel cost beside it."""
    out = ''
    for label, px, color, note in items:
        out += ('<div style="display:flex;align-items:center;gap:9px;">'
                '<div style="width:230px;height:%.0fpx;background:%s;border-radius:2px;flex-shrink:0;"></div>'
                '<div style="font-size:10.5px;"><span style="font-weight:700;">%d px</span>'
                '<span style="color:var(--muted);">&#160;&#160;%s</span>'
                '<span style="color:var(--grey);">&#160;&#160;%s</span></div></div>'
                % (max(px * scale, 3), color, px, label, note))
    tot = sum(i[1] for i in items)
    out += ('<div style="margin-top:6px;padding-top:6px;border-top:1px solid var(--line);font-size:11px;">'
            '<span style="font-weight:800;">%d px</span><span style="color:var(--muted);">&#160;&#160;%s</span></div>'
            % (tot, total_label))
    return out


NAVY, BLUE, LT, GREY, AMB, GRN = '#1c2634', '#3e7cc0', '#6aa9df', '#c9d2dc', '#d98a2b', '#0f7b4f'
TODAY = [('header band, rows 1&#8211;2', 40, NAVY, ''),
         ('six tiles in two rows of three, rows 3&#8211;9', 169, BLUE, 'six numbers, three rows deep'),
         ('verdict banner, row 10', 35, GRN, ''),
         ('KEY RESULTS title, row 11', 29, GREY, 'costs first-screen height'),
         ('how-to-read line, row 12', 35, GREY, 'costs first-screen height'),
         ('two key charts, rows 13&#8211;28', 320, LT, '10 px gutter between them'),
         ('spacer, row 29', 13, GREY, ''),
         ('SUPPORTING DETAIL title, row 30', 30, GREY, ''),
         ('four supporting charts, rows 31&#8211;40', 200, LT, '3 gutters, 30 px of grey'),
         ('spacer, row 41', 13, GREY, ''),
         ('two section charts, rows 42&#8211;51', 200, LT, '')]
PROPOSED = [('header band, 2 rows', 40, NAVY, ''),
            ('project context strip, 1 row', 20, NAVY, 'airport, runway, area, rate, year'),
            ('six tiles in one row, 3 rows', 60, BLUE, 'same six numbers, 109 px cheaper'),
            ('verdict banner, 2 rows', 40, GRN, '&#9660; freeze here, 160 px'),
            ('key charts + project location, 12 rows', 240, LT, 'three flush panels, no gutter'),
            ('four supporting charts, 9 rows', 180, LT, 'flush, no gutter')]

SCREENS = [('1920&#215;1080 at 125%', '1,536&#215;864', '1,488 &#215; 624', 'Windows default for 15.6 in &#8212; the target'),
           ('1920&#215;1080 at 150%', '1,280&#215;720', '1,232 &#215; 480', 'no band fits this without side-scroll'),
           ('1920&#215;1080 at 100%', '1,920&#215;1080', '1,872 &#215; 840', 'room to spare'),
           ('1600&#215;900 at 100%', '1,600&#215;900', '1,552 &#215; 660', 'comfortable')]

PLACE = [('1. Present worth by category', 'column G', '555 &#215; 240', '12'),
         ('2. Net present worth vs. rate', 'at 556 px', '555 &#215; 240', '12'),
         ('Project location', 'at 1,112 px', '371 &#215; 240', '12'),
         ('3 &#183; 4 &#183; 5 &#183; 8, supporting', 'the row below', '370 &#215; 180 each', '9'),
         ('6 &#183; 7, section', 'the row below that', '741 &#215; 160 each', '8')]


def block(title, lede, body, w='1fr'):
    return ('<div style="background:#fff;border:1px solid var(--line);border-radius:8px;overflow:hidden;">'
            '<div style="padding:10px 14px 8px;">'
            '<div style="font-size:13px;font-weight:800;">%s</div>'
            '<div class="sub" style="margin-top:2px;">%s</div></div>'
            '<div style="padding:0 14px 13px;">%s</div></div>' % (title, lede, body))


FINDINGS = [('Two rows of three tiles cost 169 px for six numbers.',
             'One flush row of six does the same job in 60 px.', '&#8722;109 px'),
            ('A section title and a how-to line sit between the verdict and the first plot.',
             'Both belong in the header band or a cell note, not in the first screen.', '&#8722;64 px'),
            ('Every plot is inset 10 px from its neighbour.',
             'Three gutters across the supporting row is 30 px of grey between instruments. A 1 px '
             'hairline reads as one panel and hands 27 px of width back to the plots &#8212; width, not '
             'height, so it does not shorten the first screen.', '&#8722;27 px wide'),
            ('The band is 1,239 px on a screen that offers 1,488.',
             'The 244 px right of column R is empty and can join the band. It is most of the project '
             'location card but not all of it: the card is 371 px, so the other 127 px comes off the two '
             'key plots, which go from about 615 px each to 555.', '+244 px')]


def findings():
    out = ''
    for a, b, d in FINDINGS:
        col = GRN if d.startswith('&#8722;') else BLUE
        out += ('<div style="display:flex;gap:11px;align-items:flex-start;padding:8px 0;'
                'border-bottom:1px solid var(--line);">'
                '<div style="flex-grow:1;"><div style="font-size:11.5px;font-weight:700;">%s</div>'
                '<div class="sub" style="margin-top:1px;">%s</div></div>'
                '<div style="font-size:13px;font-weight:800;color:%s;flex-shrink:0;">%s</div></div>'
                % (a, b, col, d))
    return out


grid_body = (
    '<div style="padding:18px 20px 22px;display:flex;flex-direction:column;gap:14px;">'
    '<div><div style="font-size:19px;font-weight:800;">Summary sizing and spacing</div>'
    '<div class="sub" style="margin-top:3px;font-size:11px;">Measured from the workbook as it stands. '
    'Every number below is the sheet&#8217;s own: column widths from &lt;cols&gt;, row heights from '
    '&lt;row ht&gt;, chart sizes from the drawing anchors.</div></div>'
    + '<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">'
    + block('What a 15-inch laptop actually gives Excel',
            'Grid area left after the ribbon, formula bar, column headers, sheet tabs and status bar '
            '(about 240 px), and the row headers and scrollbar (about 48 px).',
            tbl(['Screen setting', 'Window', 'Excel grid', 'Note'], SCREENS, right_from=3))
    + block('Where the proposed panels land',
            'Panel edges are pixel anchors on the drawing, not column boundaries, so they can sit flush.',
            tbl(['Panel', 'Starts', 'Size, px', 'Rows'], PLACE, right_from=1))
    + '</div>'
    + block('Four findings', 'What the first screen is spending its height and width on.', findings())
    + '<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">'
    + block('Today &#8212; 1,084 px before the first table',
            'On a 624 px grid that is 1.7 screens of scrolling before any number appears.',
            stack(TODAY, total_label='before the results table, against 624 px of screen'))
    + block('Proposed &#8212; 580 px to the end of the supporting plots',
            'Fits the 624 px grid with about 40 px to spare once the frozen-pane rule is drawn, '
            'and the project location is on it.',
            stack(PROPOSED, total_label='with the section charts and the tables below the fold'))
    + '</div></div>')

open('Grid.dc.html', 'w').write(HEAD + band(BAND_WIDE) + grid_body + '\n</div>\n' + TAIL)
print('Grid.dc.html           %d px wide' % BAND_WIDE)


# ---------------------------------------------------------------- Direction B: one answer, then the evidence
def decision_panel():
    """The six tiles and the verdict banner say the same thing twice: the banner names the winner
    and its margin, and two of the tiles repeat both. Said once, in one panel, the answer stops
    being a wall of numbers and the top of the sheet gives 133 px back."""
    rows = [('Margin to next', '$183,970', '3.1% &#183; under 5%, treat the two as tied'),
            ('Equivalent annual', '$298,577', 'per year, 30 years at 3%'),
            ('Initial construction', '$5,681,474', 'mobilisation and engineering in'),
            ('Unit cost', '$124 / S.Y.', 'over the mainline area')]
    body = ''.join('<div style="display:flex;justify-content:space-between;align-items:baseline;'
                   'gap:8px;padding:5px 0;border-bottom:1px solid #eef1f5;">'
                   '<div><div style="font-size:11px;font-weight:700;">%s</div>'
                   '<div class="sub">%s</div></div>'
                   '<div style="font-size:14px;font-weight:800;white-space:nowrap;">%s</div></div>'
                   % (k, note, v) for k, v, note in rows)
    return ('<div class="p" style="display:flex;flex-direction:column;background:var(--panel);">'
            '<div style="display:flex;align-items:center;gap:8px;">'
            '<div style="display:flex;align-items:center;justify-content:center;width:20px;height:20px;'
            'border-radius:50%;background:var(--green);flex-shrink:0;">'
            '<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3">'
            '<path d="M20 6 9 17l-5-5"/></svg></div>'
            '<div class="kicker" style="color:var(--green);">Lowest present worth</div></div>'
            '<div style="font-size:26px;font-weight:800;line-height:1.1;margin-top:3px;">$5,852,247</div>'
            '<div style="font-size:12px;font-weight:700;">Alternative 3 &#183; new concrete</div>'
            '<div class="sub" style="margin-bottom:5px;">11&#8243; P-501 on 6&#8243; P-209</div>'
            + body +
            '<div style="display:flex;gap:7px;align-items:flex-start;margin-top:6px;padding:6px 8px;'
            'background:#fdf6ea;border-left:3px solid var(--amber);">'
            '<div style="font-size:10.5px;color:#7a5a20;">'
            '<span style="font-weight:800;">The winner changes with the rate.</span> '
            'Alternative 1 is lowest above about 6.5%. See chart 2.</div></div>'
            '<div style="flex-grow:1;"></div>'
            + svg(360, 76, TN + TN_DOTS, grow=False) +
            '<div class="sub" style="text-align:center;">Sevierville, East division &#183; 45,883 S.Y. '
            '&#183; runway 75 ft</div></div>')


write('DirectionB.dc.html', BAND_WIDE,
      fold(624) + HEADER + CONTEXT + freeze_line('Frozen here &#8212; only the band and the project line')
      + row('%dpx %dpx %dpx' % (RAIL_W, KEY_W, KEY_W), 17 * ROW, [decision_panel, C1, C2])
      + row('repeat(3,minmax(0,1fr))', 9 * ROW, SUP)
      + row('repeat(3,minmax(0,1fr))', 9 * ROW, SEC)
      + tables())


# ---------------------------------------------------------------- Direction C: numbers on the first screen
def compact_results():
    head = ('<thead><tr style="background:var(--navy);"><th>Alternative</th><th>Initial</th>'
            '<th>Maint PW</th><th>Rehab PW</th><th>Lost rev</th><th>Salvage</th><th>Net PW</th>'
            '<th>vs. lowest</th><th>Closure</th></tr></thead>')
    body = ''
    for i, r in enumerate(RESULTS):
        win = i == 2
        cells = ''
        for j, v in enumerate([r[0]] + list(r[2:])):
            st = 'font-weight:800;' if (j == 0 and win) else ''
            if j == 6: st = 'font-weight:800;color:var(--green);' if win else 'font-weight:700;'
            if j == 7 and not win: st = 'color:var(--amber);'
            cells += '<td style="%s">%s</td>' % (st, v)
        body += ('<tr style="%s">%s</tr>'
                 % ('background:var(--green-soft);' if win else 'border-bottom:1px solid var(--line);', cells))
    return ('<div style="background:#fff;border-bottom:1px solid var(--line);">'
            '<table style="width:100%;border-collapse:collapse;">' + head + '<tbody>' + body
            + '</tbody></table></div>')


write('DirectionC.dc.html', BAND_WIDE,
      fold(624) + HEADER + CONTEXT + tiles() + VERDICT + freeze_line()
      + row('%dpx %dpx %dpx' % (KEY_W, KEY_W, RAIL_W), 12 * ROW, [C1, C2, location_rail])
      + sectitle('The numbers &#8212; every alternative, in full',
                 'Above the fold, so the sheet opens on figures as well as pictures.')
      + compact_results()
      + sectitle('Supporting detail', 'Charts 3 to 8 and the comparison block continue below.')
      + row('repeat(3,minmax(0,1fr))', 9 * ROW, SUP)
      + row('repeat(3,minmax(0,1fr))', 9 * ROW, SEC))


# ---------------------------------------------------------------- Review: what is still being paid for
TWICE = [('1. Present worth by category', 'A category axis titled &#8220;Alternative&#8221; over labels '
          'that already read Alternative 1, 2, 3', '20 px of height'),
         ('1. Present worth by category', 'A value axis titled &#8220;Present worth ($)&#8221; under a '
          'chart called Present worth by category', '18 px of width'),
         ('2. Net present worth vs. rate', 'A 71-character title that wraps to two lines at 558 px',
          '17 px of height'),
         ('2. Net present worth vs. rate', 'A value axis titled &#8220;Net present worth ($)&#8221;, '
          'again', '18 px of width'),
         ('2. Net present worth vs. rate', 'A tick label every quarter point &#8212; twenty-five of them, '
          'rotated', 'the bottom fifth'),
         ('Locator map', '&#8220;Latitude&#8221; and &#8220;Longitude&#8221; on a 365 &#215; 100 panel',
          'a quarter of the map')]

FOLD = [('Header band', 53, 53, 53), ('Project context', 20, 20, 20),
        ('Six tiles', 60, 0, 60), ('Verdict banner', 40, 0, 40),
        ('Decision panel with the key plots', 0, 340, 0),
        ('Key plots and the project location', 260, 0, 240),
        ('Results table', 0, 0, 129),
        ('Supporting plots', 180, 180, 0)]

TRADE = [('Main &#8212; quieter charts',
          'Nothing moves. Every panel gains the space its own labels were holding, and the sheet '
          'reads calmer for it.',
          'The least it could be. If the tile strip or the reading order is what bothers you, this '
          'does not touch either.'),
         ('Direction B &#8212; one answer, then the evidence',
          'The six tiles and the verdict banner say the same thing twice. Said once, in one panel '
          'beside the plots, the top of the sheet gives 133 px back and the answer stops being a wall.',
          'The answer no longer stays frozen while you scroll, and six separate numbers become one '
          'block to read rather than six to scan.'),
         ('Direction C &#8212; numbers on the first screen',
          'The results table clears the fold, so the sheet opens on figures as well as pictures. '
          'Everything else keeps its place.',
          'Charts 3 to 8 all drop below the fold. The sheet stops reading picture-first, which is '
          'the order it was rearranged into two rounds ago.')]


def fold_table():
    rows = []
    for label, a, b, c in FOLD:
        rows.append((label, a or '&#8212;', b or '&#8212;', c or '&#8212;'))
    tot = ['%d px' % sum(x[i] for x in FOLD) for i in range(1, 4)]
    rows.append(('<b>Above the fold</b>', '<b>%s</b>' % tot[0], '<b>%s</b>' % tot[1], '<b>%s</b>' % tot[2]))
    return tbl(['Band', 'Main', 'Direction B', 'Direction C'], rows, right_from=1)


def trades():
    out = ''
    for name, why, cost in TRADE:
        out += ('<div style="padding:9px 0;border-bottom:1px solid var(--line);">'
                '<div style="font-size:12px;font-weight:800;">%s</div>'
                '<div style="font-size:11px;margin-top:2px;">%s</div>'
                '<div class="sub" style="margin-top:3px;"><b>The cost:</b> %s</div></div>'
                % (name, why, cost))
    return out


review_body = (
    '<div style="padding:18px 20px 22px;display:flex;flex-direction:column;gap:14px;">'
    '<div><div style="font-size:19px;font-weight:800;">What the dashboard is still paying for</div>'
    '<div class="sub" style="margin-top:3px;font-size:11px;">Measured from the workbook as built. '
    'The layout is settled; what is left is labelling that repeats itself and one open question '
    'about what belongs on the first screen.</div></div>'
    + block('Said twice', 'Every one of these is a chart telling you something its own title '
            'already told you, in space the plot could have had.',
            tbl(['Chart', 'What it says twice', 'Costs'], TWICE, right_from=2))
    + '<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">'
    + block('What clears a 624 px screen', 'The grid a 15.6 in laptop gives Excel at 125% scaling. '
            'Each column is one of the three directions.', fold_table())
    + block('And what each one costs', 'A set where only the favourite gets a case made for it is '
            'not a choice.', trades())
    + '</div></div>')

open('Review.dc.html', 'w').write(HEAD + band(BAND_WIDE) + review_body + '\n</div>\n' + TAIL)
print('Review.dc.html         %d px wide' % BAND_WIDE)
