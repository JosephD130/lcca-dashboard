"""Generates the four artboards for the logo-placement pass.

Sizes, anchors and the aspect-ratio finding are read out of the workbook package, so the audit
artboard states what is actually in the file rather than what anyone remembers.
"""
import os, re, zipfile, struct

HERE = os.path.dirname(os.path.abspath(__file__))
WB = '/home/user/lcca-dashboard/tdot-lcca/TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm'
EMU = 914400.0
INK, MUTE, NAVY, BLUE, PALE, RULE = '#1D2733', '#595959', '#1D2733', '#2A78D6', '#EAF2FB', '#D4D8DD'
GOOD, WARN = '#2E8B1F', '#C1440E'

z = zipfile.ZipFile(WB)
png = z.read('xl/media/image4.png')
open(os.path.join(HERE, 'tdot-logo.png'), 'wb').write(png)
NW, NH = struct.unpack('>II', png[16:24])
NATIVE = NW / float(NH)

SHEETS = {'drawing2': 'Overview', 'drawing3': 'Instructions',
          'drawing4': 'General Information', 'drawing5': 'Maintenance Policies'}
placed = []
for d, sheet in SHEETS.items():
    x = z.read('xl/drawings/%s.xml' % d).decode('utf8', 'replace')
    anchor = re.search(r'<xdr:twoCellAnchor[^>]*>(?:(?!</xdr:twoCellAnchor>).)*?<xdr:pic>.*?</xdr:twoCellAnchor>',
                       x, re.S).group(0)                      # the logo's own anchor, not another shape's
    cx, cy = (int(v) for v in re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', anchor).groups())
    col, coloff, row = (int(v) for v in re.search(
        r'<xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>(\d+)</xdr:colOff><xdr:row>(\d+)</xdr:row>',
        anchor).groups())
    w, h = cx / EMU, cy / EMU
    placed.append(dict(sheet=sheet, w=w, h=h, aspect=w / h, col=col, coloff=coloff, row=row,
                       stretch=(h / (w / NATIVE) - 1) * 100))

MISSING = ['Pay_Items', 'Summary', 'Typical Values', 'Method', 'Alternative worksheets (5 templates)']

HEAD = '''<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <script src="./support.js"></script>
</head>
<body>
<x-dc>
<helmet>
  <style>
    body { margin: 0; background: #8A8F98; font-family: Arial, Helvetica, sans-serif; color: #1D2733; }
    a { color: #2A78D6; } a:hover { color: #1D5BA6; }
    .sheet { background: #FFFFFF; box-sizing: border-box; padding: 20px 22px 24px; }
    .title { font-size: 15px; font-weight: bold; letter-spacing: .03em; }
    .lede { font-size: 11px; color: #595959; line-height: 1.5; }
    .lab { font-size: 9.5px; color: #595959; letter-spacing: .05em; text-transform: uppercase; }
    .card { border: 1px solid #D4D8DD; }
    .hd { background: #1D2733; color: #FFFFFF; font-size: 10.5px; font-weight: bold;
          letter-spacing: .05em; padding: 5px 9px; }
    .rule { border-top: 2px solid #1D2733; margin-top: 4px; }
    table.kv { border-collapse: collapse; width: 100%; }
    table.kv th { background: #D9D9D9; border: 1px solid #BFBFBF; font-size: 10px; padding: 4px 7px;
                  text-align: left; }
    table.kv td { border: 1px solid #D0D4D9; font-size: 10.5px; padding: 4px 7px; }
    td.n { text-align: right; font-variant-numeric: tabular-nums; }
    .yes { color: #2E8B1F; font-weight: bold; }
    .no { color: #C1440E; font-weight: bold; }
    .note { font-size: 10px; color: #595959; line-height: 1.45; }
    .pill { display: inline-block; font-size: 9.5px; font-weight: bold; letter-spacing: .04em;
            padding: 3px 8px; text-transform: uppercase; }
    /* a schematic sheet: row bands, so a placement can be shown without redrawing the sheet */
    .mini { border: 1px solid #C6CBD0; background: #FFFFFF; }
    .r { height: 7px; background: #EDF0F3; margin: 3px 5px; }
    .r.hdr { background: #1D2733; height: 9px; margin: 5px 5px 4px; }
    .r.band { background: #EAF2FB; }
    .r.tall { height: 22px; }
    .frozen { border-bottom: 2px dashed #9AA0A6; }
    .hdrzone { background: repeating-linear-gradient(135deg, #F4F6F8, #F4F6F8 6px, #EAEDF1 6px, #EAEDF1 12px);
               border-bottom: 1px solid #C6CBD0; padding: 5px; }
  </style>
</helmet>
<div class="sheet">
'''
TAIL = '</div>\n</x-dc>\n</body>\n</html>\n'


def write(name, body):
    open(os.path.join(HERE, name), 'w', encoding='utf-8').write(HEAD + body + TAIL)
    print('wrote', name)


def logo(width_px, stretch=0.0):
    """The lockup at a given rendered width; stretch>0 shows it squashed the way the file has it."""
    h = width_px / NATIVE * (1 + stretch / 100.0)
    return ('<img src="tdot-logo.png" alt="TDOT" style="width: %.0fpx; height: %.1fpx; display: block;"/>'
            % (width_px, h))


def mini(rows, head=None, label=''):
    """rows: list of (kind, content_html). kind in {'', 'hdr', 'band', 'tall', 'frozen'}"""
    body = ''
    if head is not None:
        body += '<div class="hdrzone">%s</div>' % head
    for kind, inner in rows:
        cls = 'r' + (' ' + kind if kind else '')
        body += ('<div class="%s" style="%s">%s</div>'
                 % (cls, 'display: flex; align-items: center; padding: 0 4px;' if inner else '', inner or ''))
    return ('<div><div class="mini" style="padding-bottom: 6px;">%s</div>'
            '<div class="lab" style="margin-top: 5px;">%s</div></div>' % (body, label))


# ---------------------------------------------------------------- 1. what is there today
rows = ''.join(
    '<tr><td>%s</td><td class="n">%.2f</td><td class="n">%.2f</td><td class="n">%.3f</td>'
    '<td class="n" style="color: %s; font-weight: bold;">+%.1f%%</td><td>column %s, row %d</td></tr>'
    % (p['sheet'], p['w'], p['h'], p['aspect'], WARN, p['stretch'], chr(65 + p['col']), p['row'] + 1)
    for p in sorted(placed, key=lambda q: q['sheet']))
missing = ''.join('<tr><td>%s</td><td colspan="5" class="no">no logo</td></tr>' % m for m in MISSING)

write('LogoAudit.dc.html',
      '<div class="title">WHERE THE LOGO IS TODAY</div>'
      '<div class="lede" style="margin: 6px 0 16px; max-width: 780px;">Four of the nine sheets a user opens carry '
      'the mark, each anchored and sized a little differently, and all four are drawn taller than the artwork. '
      'The file is %d &times; %d, an aspect of %.3f; every placement is squarer than that, so the lettering is '
      'stretched vertically.</div>'
      '<div style="display: flex; gap: 26px; align-items: flex-start;">'
      '<div style="flex: none;">'
      '<div class="lab">The artwork</div>%s'
      '<div class="note" style="margin-top: 6px; width: 210px;">%d &times; %d px<br/>aspect %.3f</div></div>'
      '<div style="flex: none;">'
      '<div class="lab">As placed, +11%%</div>'
      '<div style="border: 1px dashed %s; padding: 4px; display: inline-block;">%s</div>'
      '<div class="note" style="margin-top: 6px; width: 210px; color: %s;">The tallest of the four. The '
      '&ldquo;TN&rdquo; square is no longer square.</div></div>'
      '<div style="flex: 1 1 0; min-width: 0;">'
      '<table class="kv"><tr><th>Sheet</th><th>Width in</th><th>Height in</th><th>Aspect</th>'
      '<th>Stretch</th><th>Anchored at</th></tr>%s%s</table>'
      '<div class="note" style="margin-top: 7px;">At the native aspect a %.2f in wide lockup is %.2f in tall, '
      'not %.2f in. Whatever placement rule is chosen, this is worth correcting in the same pass.</div>'
      '</div></div>'
      % (NW, NH, NATIVE, logo(210), NW, NH, NATIVE,
         WARN, logo(150, 11.1), WARN, rows, missing,
         placed[0]['w'], placed[0]['w'] / NATIVE, placed[0]['h']))


# ---------------------------------------------------------------- the three options
def option(letter, name, rule, minis, pros, cons, covers, recommended=False):
    pill = ('<span class="pill" style="background: %s; color: #FFFFFF;">Recommended</span>' % GOOD
            if recommended else '')
    return ('<div style="display: flex; gap: 10px; align-items: baseline;">'
            '<span class="pill" style="background: %s; color: #FFFFFF;">Option %s</span>'
            '<span class="title">%s</span>%s</div>'
            '<div class="lede" style="margin: 8px 0 4px; max-width: 820px;">%s</div>'
            '<div class="rule"></div>'
            '<div style="display: grid; grid-template-columns: repeat(%d, minmax(0, 1fr)); gap: 16px; '
            'margin-top: 14px;">%s</div>'
            '<div style="display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; '
            'margin-top: 18px;">'
            '<div class="card"><div class="hd">WHAT IT COVERS</div><div class="note" style="padding: 8px 10px;">%s</div></div>'
            '<div class="card"><div class="hd">WHAT IT COSTS</div><div class="note" style="padding: 8px 10px;">%s</div></div>'
            '<div class="card"><div class="hd">THE TRADEOFF</div><div class="note" style="padding: 8px 10px;">%s</div></div>'
            '</div>' % (NAVY, letter, name, pill, rule, len(minis), ''.join(minis), covers, pros, cons))


HEADER_MINI = mini([('hdr', ''), ('', ''), ('band', ''), ('', ''), ('', ''), ('', ''), ('', '')],
                   head='<div style="display: flex; justify-content: flex-end;">%s</div>' % logo(74),
                   label='Any sheet, page layout view')
SCREEN_MINI = mini([('hdr', ''), ('', ''), ('band', ''), ('', ''), ('', ''), ('', ''), ('', '')],
                   label='The same sheet on screen: unchanged')

write('LogoHeader.dc.html',
      option('A', 'The logo lives in the page header, not in a cell',
             'Every sheet gets the mark through Excel&rsquo;s own page header, the same way a letterhead works. '
             'It sits above the grid, so it costs no row, moves no cell reference, and cannot land above a frozen '
             'pane. The five alternative-worksheet templates carry it too, so every alternative the form creates '
             'inherits it without the macro changing.',
             [HEADER_MINI, SCREEN_MINI],
             covers='All sixteen sheets, including each alternative worksheet as it is created. One size, one '
                    'position, set once per sheet.',
             pros='Nothing. No row, no column, no cell reference, no space above a freeze line. The Summary keeps '
                  'its dashboard strip and Pay_Items keeps its frozen header exactly as they are.',
             cons='It is not visible in the normal editing view &mdash; only in Page Layout, print preview and on '
                  'paper. If the workbook is mostly read on screen rather than printed, the mark is mostly unseen.',
             recommended=False))

MASTHEAD_MINI = mini([('tall', logo(96)), ('', ''), ('band', ''), ('', ''), ('', ''), ('', ''), ('', '')],
                     label='Overview, Instructions, General Information, Maintenance Policies')
WORKING_MINI = mini([('hdr', ''), ('', ''), ('', ''), ('', ''), ('', ''), ('', ''), ('', '')],
                    head='<div style="display: flex; justify-content: flex-end;">%s</div>' % logo(74),
                    label='Summary, Pay_Items, Typical Values, Method, alternatives')

write('LogoMasthead.dc.html',
      option('B', 'Masthead where the sheet is read, page header everywhere else',
             'Keep the visible logo on the four sheets that open like a document and are read before anything is '
             'typed, at one anchor and one size instead of the four that exist now. The working grids &mdash; the '
             'Summary, Pay_Items, the alternative worksheets &mdash; get it from the page header as in option A. '
             'One rule, stated plainly: visible where the sheet is read, printed everywhere.',
             [MASTHEAD_MINI, WORKING_MINI],
             covers='Visible on the four document sheets; printed on all sixteen. The four existing placements are '
                    'squared up to the artwork and anchored identically at B2, 1.95 by 0.85 in.',
             pros='Nothing beyond what is already spent: those four sheets already give the logo that space. It '
                  'takes away the size and anchor drift rather than adding anything.',
             cons='On screen the workbook is not uniformly branded. Whether that matters depends on whether anyone '
                  'looks at the Summary and thinks it should carry the mark.'))

BAND = ('<div style="display: flex; align-items: center; gap: 6px; width: 100%;">'
        '<span style="background: #1D2733; height: 11px; width: 52px;"></span>'
        '<span style="background: #2A78D6; height: 11px; width: 30px;"></span>'
        '<span style="flex: 1 1 auto;"></span>' + logo(58) + '</div>')
BAND_MINI = mini([('tall', BAND), ('frozen', ''), ('', ''), ('band', ''), ('', ''), ('', ''), ('', '')],
                 label='Every sheet: nav buttons left, logo right, 40pt')
BAND_COST = mini([('tall', BAND), ('frozen', ''), ('hdr', ''), ('', ''), ('', ''), ('', ''), ('', '')],
                 label='Pay_Items: the band sits above the frozen header, for good')

write('LogoBand.dc.html',
      option('C', 'A branded band in row 1 of every sheet',
             'Force the consistency on screen. Row 1 already exists on every sheet and already carries the '
             'navigation buttons; raise it from 21 to 40 points and put the lockup at the right-hand end of it. '
             'Forty points is the floor: below it &ldquo;Department of Transportation&rdquo; stops being legible, '
             'and at 40pt the lockup is 1.27 in wide.',
             [BAND_MINI, BAND_COST],
             covers='Every sheet, visibly, in the same place, at the same size. The strongest answer to &ldquo;it '
                    'should look like one workbook&rdquo;.',
             pros='Nineteen points of height on every sheet. On Pay_Items and the alternative worksheets that space '
                  'is above the frozen header, so it is gone from every screen for good, not just the first one. '
                  'The Summary&rsquo;s row 1 already carries the buttons, the title and the project identity line, '
                  'so the logo has to share a crowded row.',
             cons='Most visible, most expensive, and the one option that changes how the working sheets feel to '
                  'use. Chosen and built in v1.2.0: row 1 is also a print title now, so the band and the mark '
                  'print at the top of every page, not only the first.',
             recommended=True))
