"""Static verification of the delivered workbook: package hygiene, sheet inventory, Summary layout,
charts (axis titles, legends, hidden-cell plotting), the Method and Typical Values sheets, and the
pavement-section block. Reads the file as a zip and with openpyxl; no Excel or LibreOffice needed.

usage: python3 verify_build.py <workbook.xlsm>
"""
import sys, re, zipfile, warnings
from openpyxl import load_workbook
warnings.filterwarnings('ignore')

WB = sys.argv[1] if len(sys.argv) > 1 else 'TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm'
z = zipfile.ZipFile(WB)
names = z.namelist()
rd = lambda n: z.read(n).decode('utf-8', 'replace')
wb = load_workbook(WB, keep_vba=True)
ws = wb['Summary']
# the Summary row map, the same constants build_summary.py lays the sheet out with
KPI0, HDR, T0, T1 = 3, 11, 12, 15
VER, VER2, CMP, CMPH, CMP0 = 17, 18, 20, 21, 22
CHT, HOW, C1, C2, C3, BMN = 28, 29, 30, 48, 66, 87
SEC_T, CH67, NOTE67, MAP0, A0 = 90, 100, 120, 130, 132
VBACH = 123       # where the chart the Alternative Setup form maintains is parked
SENS0, SENS1 = 12, 36
BM = 89           # unit-cost benchmark data block, in the chart-data columns
fails = []
checks = 0


def check(name, ok, detail=''):
    global checks
    checks += 1
    print(('PASS  ' if ok else 'FAIL  ') + name + (('   |   ' + str(detail)[:160]) if detail else ''))
    if not ok: fails.append(name)


print('=' * 78); print('PACKAGE'); print('=' * 78)
check('zip test: every part parses', all(z.testzip() is None for _ in [0]))
check('VBA project still present', 'xl/vbaProject.bin' in names)
check('all 106 ActiveX controls kept', sum(1 for n in names if re.match(r'xl/activeX/activeX\d+\.xml$', n)) == 106,
      sum(1 for n in names if re.match(r'xl/activeX/activeX\d+\.xml$', n)))
check('every ActiveX binary keeps a content type', rd('[Content_Types].xml').count('/xl/activeX/activeX') == 106 * 2 or
      len(re.findall(r'PartName="/xl/activeX/activeX\d+\.bin"', rd('[Content_Types].xml'))) == 106)
check('printer settings removed', not [n for n in names if 'printerSettings' in n])
check('no page setup still points at a printer part', not re.search(r'<pageSetup[^>]*r:id=', ''.join(rd(n) for n in names if n.startswith('xl/worksheets/sheet'))))
check('Power Query stub removed', not [n for n in names if n.startswith('customXml/')])
check('cached SharePoint path removed', 'absPath' not in rd('xl/workbook.xml'))
core = rd('docProps/core.xml')
check('document properties carry no personal names', 'Jim Bruinsma' not in core and 'Marti Denis' not in core,
      re.findall(r'<dc:creator>([^<]*)', core) + re.findall(r'<cp:lastModifiedBy>([^<]*)', core))
check('external links removed', not [n for n in names if n.startswith('xl/externalLinks')] and 'externalReferences' not in rd('xl/workbook.xml'))
check('calculation chain dropped (full recalc on open)', 'xl/calcChain.xml' not in names and 'fullCalcOnLoad="1"' in rd('xl/workbook.xml'))
check('no cell carries an invalid style id', 's="None"' not in ''.join(rd(n) for n in names if n.startswith('xl/worksheets/sheet')))

print(); print('=' * 78); print('SHEETS'); print('=' * 78)
expect = ['TMP(NewPCC)_IndirectCost', 'Overview', 'Instructions', 'General Information', 'Pay_Items', 'Database',
          'Maintenance Policies', 'TMP(NewHMA)_IndirectCost', 'RevenueData', 'TMP(NewHMA)', 'TMP(NewPCC)',
          'TMP(HMARehab)', 'TMP(PCCRehab)', 'Summary', 'Typical Values', 'Method']
check('sheet order unchanged, Typical Values and Method appended last', wb.sheetnames == expect, wb.sheetnames)
check('the four templates stay hidden', all(wb[s].sheet_state != 'visible' for s in expect if s.startswith('TMP(')))
check('Database and RevenueData stay hidden', wb['Database'].sheet_state == 'hidden' and wb['RevenueData'].sheet_state == 'hidden')

print(); print('=' * 78); print('SUMMARY LAYOUT'); print('=' * 78)
check('columns A to F hidden (the block the form writes)', all(ws.column_dimensions[c].hidden for c in 'ABCDEF'),
      {c: ws.column_dimensions[c].hidden for c in 'ABCDEFG'})
check('column G onward visible', not ws.column_dimensions['G'].hidden)
check('navigation buttons sit in the first visible column', 'HYPERLINK' in str(ws['G1'].value) and 'HYPERLINK' in str(ws['H1'].value),
      (ws['G1'].value, ws['H1'].value))
check('title moved beside the buttons', ws['J1'].value == 'LCCA SUMMARY')
check('note explains the hidden block', 'hidden' in str(ws.cell(HDR - 1, 7).value) and 'unhide' in str(ws.cell(HDR - 1, 7).value).lower())
check('results table header intact', [ws.cell(HDR, c).value for c in range(7, 19)] ==
      ['Worksheet', 'Alternative', 'Type', 'Initial construction', 'Maintenance PW', 'Rehabilitation PW', 'Lost revenue PW',
       'Salvage PW', 'Net present worth', 'vs. lowest NPW', 'Closure days in period', 'Runway availability'],
      [ws.cell(HDR, c).value for c in range(7, 19)])
check('results table reads the Database sheet by INDEX (survives a row deletion)',
      all(str(ws.cell(r, 7).value).startswith('=IF(INDEX(Database!$D:$D') for r in range(T0, T1 + 1)))
check('comparison block present (RealCost layout)', str(ws.cell(CMP, 7).value).startswith('COMPARISON') and ws.cell(CMPH, 7).value == 'Alternative')
check('section block present with both description strings', ws.cell(SEC_T + 2, 16).value == 'Section from the quantities'
      and ws.cell(SEC_T + 2, 17).value == 'Same quantities over mainline + shoulder')
check('asphalt unit weight is an input cell', ws.cell(SEC_T + 1, 10).value == 145)
check('the two verdict lines sit under the results table',
      str(ws.cell(VER, 7).value).startswith('=IF(COUNT($O$%d:$O$%d)=0' % (T0, T1))
      and 'lowest-cost alternative' in str(ws.cell(VER2, 7).value),
      (str(ws.cell(VER, 7).value)[:40], str(ws.cell(VER2, 7).value)[:40]))
check('the charts block is labelled and carries its how-to-read line',
      ws.cell(CHT, 7).value == 'CHARTS' and str(ws.cell(HOW, 7).value).startswith('How to read'),
      (ws.cell(CHT, 7).value, str(ws.cell(HOW, 7).value)[:30]))
check('the notes under the chart rows are where the row map says',
      'per S.Y.' in str(ws.cell(BMN, 7).value) and 'Charts 6 and 7' in str(ws.cell(NOTE67, 7).value)
      and 'Alternative Setup form draws' in str(ws.cell(VBACH - 1, 7).value),
      (str(ws.cell(NOTE67, 7).value)[:30], str(ws.cell(VBACH - 1, 7).value)[:30]))
draw = rd('xl/drawings/drawing10.xml')
anchors = [(int(c), int(r)) for c, r in
           re.findall(r'<xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>\d+</xdr:colOff>'
                      r'<xdr:row>(\d+)</xdr:row>', draw)]
# the laptop band: two key charts side by side, then four supporting, then the
# two section charts - all inside the 1,218 px the narrowed columns G:R now span
want = [(6, C1 - 1), (12, C1 - 1),                                  # key pair
        (6, C2 - 1), (8, C2 - 1), (12, C2 - 1), (15, C2 - 1),       # four supporting
        (6, C3 - 1), (8, C3 - 1)]                                   # the two section charts
check('the eight Summary charts are anchored where the row map says',
      all(a in anchors for a in want), sorted(set(anchors)))
check('no chart starts beyond the dashboard band (column R)',
      all(c <= 15 for c, r in anchors if r in (C1 - 1, C2 - 1, C3 - 1)),
      sorted(c for c, r in anchors if r in (C1 - 1, C2 - 1, C3 - 1)))
check('the chart the form maintains is parked below the dashboard', (6, VBACH - 1) in anchors,
      [a for a in anchors if a[0] == 6])
check('chart-data block labelled do not edit', 'do not edit' in str(ws.cell(1, 23).value))
check('the sensitivity block names the lowest alternative at each of its rates',
      ws.cell(SENS0 - 1, 23 + 1 + 4).value == 'Lowest at this rate'
      and str(ws.cell(SENS0, 28).value).startswith('=IF(COUNTIF(')
      and str(ws.cell(SENS1, 28).value).startswith('=IF(COUNTIF('),
      ws.cell(SENS0 - 1, 28).value)
TILE_C = [7, 10, 14]        # G:I, J:M, N:Q - even widths over the table's own column widths
check('the rate-sensitivity tile tests every rate in that block, not just its ends',
      'COUNTIF($AB$%d:$AB$%d' % (SENS0, SENS1) in str(ws.cell(KPI0 + 5, TILE_C[2]).value),
      str(ws.cell(KPI0 + 5, TILE_C[2]).value)[:90])
check('an alternative that does not exist gets no name', all('""' in str(ws.cell(4, c).value) for c in range(24, 28)),
      ws.cell(4, 24).value)
wbx = rd('xl/workbook.xml')
check('print areas defined for the Summary and General Information',
      'Summary!$A$1:$V$141' in wbx and "'General Information'!$A$1:$J$52" in wbx,
      re.findall(r'<definedName name="_xlnm.Print_Area"[^>]*>([^<]*)', wbx))

print(); print('=' * 78); print('CHARTS'); print('=' * 78)
charts = sorted((n for n in names if re.match(r'xl/charts/chart\d+\.xml$', n)), key=lambda n: int(re.search(r'\d+', n.split('/')[-1]).group()))
check('fifteen charts: six original, eight on the Summary, one locator map', len(charts) == 15, len(charts))
for n in charts:
    c = rd(n)
    ts = re.findall(r'<a:t>([^<]*)</a:t>', c)
    has_x = any(t in ('Calendar year', 'Alternative', 'Discount rate (%)', 'Scenario', 'Longitude') for t in ts)
    has_y = any(t in ('Cost ($)', 'Present worth ($)', 'Net present worth ($)', 'Spend, undiscounted ($)',
                      'Cumulative discounted cost ($)', 'Closure days', 'Thickness (inches)', 'Latitude',
                      'Initial construction / mainline S.Y.') for t in ts)
    check('%s names both axes' % n.split('/')[-1], has_x and has_y, ts[:4])
summary_charts = charts[6:13]      # the seven cost and section charts; the benchmark and the map carry no legend
check('the seven Summary charts put the legend beside the plot',
      all(re.search(r'<legendPos val="r"/>', rd(n)) for n in summary_charts),
      [re.findall(r'<legendPos val="(\w)"/>', rd(n)) for n in summary_charts])
check('money axes read to one decimal in millions',
      sum('$#,##0.0,,&quot;M&quot;' in rd(n) or '$#,##0.0,,"M"' in rd(n) for n in charts) >= 5)
c6 = rd([n for n in charts if n.endswith('chart6.xml')][0])
check('the original Alternatives Comparison chart plots the hidden columns', 'plotVisOnly val="0"' in c6)
check('the alternative-sheet charts keep calendar years on the category axis',
      all('Calendar year' in rd(n) for n in charts[:5]))

print(); print('=' * 78); print('DASHBOARD STRIP'); print('=' * 78)
labels = [ws.cell(KPI0 + 4 * (k // 3), TILE_C[k % 3]).value for k in range(6)]
check('six KPI tiles above the results table',
      labels == ['LOWEST PRESENT WORTH', 'MARGIN TO NEXT', 'EQUIVALENT ANNUAL COST', 'INITIAL CONSTRUCTION',
                 'UNIT COST', 'RATE SENSITIVITY'], labels)
vals = [str(ws.cell(KPI0 + 1 + 4 * (k // 3), TILE_C[k % 3]).value) for k in range(6)]
check('every tile is a formula over cells that already exist', all(v.startswith('=') for v in vals), vals[:2])
tile_rows = (KPI0, KPI0 + 1, KPI0 + 2, KPI0 + 4, KPI0 + 5, KPI0 + 6)
spans = sorted({(m.min_col, m.max_col) for m in ws.merged_cells.ranges if m.min_row in tile_rows})
check('the six tiles are even: three columns then four then four, not four each',
      spans == [(7, 9), (10, 13), (14, 17)]
      and sum(1 for m in ws.merged_cells.ranges if m.min_row in tile_rows) == 18, spans)
sx = rd('xl/worksheets/sheet14.xml')
cfs = dict(re.findall(r'<conditionalFormatting sqref="([^"]+)"><cfRule type="(\w+)"', sx))
check('bars inside the net present worth column', cfs.get('O%d:O%d' % (T0, T1)) == 'dataBar', cfs)
check('the margin tile turns amber only when the two best are within five percent',
      cfs.get('J%d:M%d' % (KPI0, KPI0 + 2)) == 'expression'
      and '<0.05' in sx.replace('&lt;', '<'), cfs)
check('lowest-cost row still highlighted in both tables',
      cfs.get('G%d:R%d' % (T0, T1)) == 'expression' and cfs.get('G%d:P%d' % (CMP0, CMP0 + 3)) == 'expression', cfs)
bm = [n for n in charts if 'Unit cost against published' in rd(n)]
bx = rd(bm[0]) if bm else ''
check('chart 8 is the unit-cost benchmark: a stacked bar whose first series is invisible',
      len(bm) == 1 and 'stacked' in bx and 'a:noFill' in bx and '<legend>' not in bx,
      (len(bm), 'stacked' in bx, 'a:noFill' in bx))
check('the published band is one grey point carrying its own value label',
      '<dPt>' in bx and '<dLbls>' in bx, (bx.count('<dPt>'), bx.count('<dLbls>')))
check('the band comes from cells, so it can be moved without touching the chart',
      ws.cell(BM + 3, 23).value == 'Unit cost / published range'
      and ws.cell(BM + 2, 28).value == 210 and ws.cell(BM + 3, 28).value == 70,
      (ws.cell(BM + 2, 28).value, ws.cell(BM + 3, 28).value))
check('chart 8 sits beside chart 5 and carries its own note',
      'per S.Y.' in str(ws.cell(BMN, 7).value) and 'Typical Values' in str(ws.cell(BMN, 7).value),
      str(ws.cell(BMN, 7).value)[:60])


print(); print('=' * 78); print('LOCATOR MAP AND PROJECT IDENTITY'); print('=' * 78)
check('project identity line on its own row under the band',
      'D$9' in str(ws['G2'].value) and 'construction' in str(ws['G2'].value), str(ws['G2'].value)[:50])
check('map data block labelled and outside the print area', 'MAP DATA' in str(ws.cell(MAP0, 23).value))
coords = [(ws.cell(r, 25).value, ws.cell(r, 26).value) for r in range(A0, A0 + 79)]
check('79 airports in the map block, 74 with published coordinates',
      len(coords) == 79 and sum(1 for a, b in coords if isinstance(a, (int, float)) and isinstance(b, (int, float))) == 74,
      (len(coords), sum(1 for a, b in coords if isinstance(a, (int, float)))))
check('coordinates fall inside Tennessee',
      all(-90.5 < a < -81.5 and 34.9 < b < 36.8 for a, b in coords if isinstance(a, (int, float))))
border = [(ws.cell(r, 34).value, ws.cell(r, 35).value) for r in range(A0, A0 + 213)]
bpts = [(a, b) for a, b in border if isinstance(a, (int, float))]
check('state outline embedded with gaps between segments', len(bpts) == 179 and len(border) > len(bpts), (len(bpts), len(border)))
check('the selected airport is looked up, not typed',
      str(ws.cell(A0, 30).value).startswith('=IFERROR(IF(INDEX(') and str(ws.cell(A0, 31).value).startswith('=IFERROR(IF(INDEX('))
mapch = [c for c in wb['Summary']._charts if c.tagname == 'scatterChart']
check('locator map has the outline, the three Grand Divisions and this project', len(mapch) == 1 and len(mapch[0].series) == 5,
      [(c.tagname, len(c.series)) for c in wb['Summary']._charts])
check('airports grouped by division so each is its own series',
      [ws.cell(r, 27).value for r in (A0, A0 + 20, A0 + 43, A0 + 73)] == ['West', 'West', 'Middle', 'East'],
      [ws.cell(r, 27).value for r in (A0, A0 + 20, A0 + 43, A0 + 73)])
check('legend in cells, one per division plus this project',
      [ws.cell(17, c).value for c in range(19, 23)] == ['\u25a0 West', '\u25a0 Middle', '\u25a0 East', '\u25a0 This project'],
      [ws.cell(17, c).value for c in range(19, 23)])
check('card header band across S2:V2', ws['S2'].value == 'PROJECT LOCATION' and ws['S2'].fill.fgColor.rgb.endswith('1D2733'))
check('pricing basis stated beside the map', 'Unit Cost' in str(ws['S28'].value) and 'empty' in str(ws['S28'].value))
check('map axes are named and their degree labels suppressed',
      mapch and mapch[0].x_axis.numFmt.formatCode == ';;;' and mapch[0].y_axis.numFmt.formatCode == ';;;')
check('project block names airport, county, region, coordinates, elevation',
      [ws.cell(20 + k, 19).value for k in range(7)] ==
      ['Airport', 'City / county', 'TDOT Grand Division', 'Coordinates', 'Elevation', 'Branch / project', 'Mainline area'],
      [ws.cell(20 + k, 19).value for k in range(7)])
check('the sheet says where the coordinates come from and how to export KML',
      'embedded' in str(ws['S29'].value) and 'ExportLCCAKML' in str(ws['S30'].value))
check('runway width for the export footprint is an input cell', ws['S27'].value == 'Runway width, ft'
      and ws['T27'].value == 100, (ws['S27'].value, ws['T27'].value))
check('the three notes beside the map are merged so they do not run into the chart data',
      all(str(rng) in [str(m) for m in ws.merged_cells.ranges] for rng in ('S28:V28', 'S29:V29', 'S30:V30')),
      [str(m) for m in ws.merged_cells.ranges][-4:])
check('print area widened to take the map', 'Summary!$A$1:$V$141' in rd('xl/workbook.xml'))

print(); print('=' * 78); print('METHOD AND TYPICAL VALUES'); print('=' * 78)
me = wb['Method']
check('Method sheet has its navigation buttons', 'HYPERLINK' in str(me['A1'].value) and 'HYPERLINK' in str(me['B1'].value))
check('Method sheet states it feeds nothing', 'nothing here feeds' in str(me['A2'].value).lower())
titles = [me.cell(r, 1).value for r in range(1, me.max_row + 1) if isinstance(me.cell(r, 1).value, str) and re.match(r'^\d+\. ', me.cell(r, 1).value)]
check('eleven sections, in order', len(titles) == 11 and titles[0].startswith('1.') and titles[-1].startswith('11.'), titles)
formula_cells = [me.cell(r, 3).value for r in range(1, me.max_row + 1) if isinstance(me.cell(r, 3).value, str) and me.cell(r, 3).value.startswith('=')]
live = [c for c in formula_cells if "'General Information'!$D$3" in c or 'COUNTA(Database' in c]
check('the live block reads the workbook (period, rate, CRF, discount factor, alternatives)', len(live) >= 5, len(live))
shown = [me.cell(r, 3) for r in range(1, me.max_row + 1) if isinstance(me.cell(r, 3).value, str)
         and me.cell(r, 3).value.startswith('=') and me.cell(r, 3).data_type == 's']
check('documented formulas are stored as text, not evaluated', len(shown) >= 20, len(shown))
check('assumptions section names the four open questions',
      all(k in ''.join(str(me.cell(r, 1).value) for r in range(1, me.max_row + 1)) for k in
          ['Engineering on initial construction', 'Salvage is not symmetric', 'Lost revenue is gross revenue', 'PCC joint length']))
tv = wb['Typical Values']
srcs = [tv.cell(r, 7).value for r in range(1, tv.max_row + 1) if isinstance(tv.cell(r, 7).value, str) and tv.cell(r, 7).value.startswith('http')]
check('Typical Values sources are plain text, not hyperlinks', len(srcs) >= 20 and not any('HYPERLINK' in str(tv.cell(r, 7).value)
      for r in range(1, tv.max_row + 1)), len(srcs))
hl = sum(len(sh._hyperlinks) for sh in wb.worksheets)
check('only the original tn.gov hyperlink object remains', hl <= 1, hl)

print(); print('=' * 78); print('GENERAL INFORMATION'); print('=' * 78)
gi = wb['General Information']
check('analysis period default 30 years', gi['D33'].value == 30)
check('discount rate default 3 percent', gi['D34'].value == 3)
gix = rd('xl/worksheets/sheet4.xml')
check('airport dropdown covers all 79 airports (D9 list = L10:L88)', '<formula1>$L$10:$L$88</formula1>' in gix,
      re.findall(r'<formula1>([^<]*)</formula1>', gix)[:4])
check('79 airports and 17 with revenue data behind it',
      sum(1 for r in range(10, 89) if gi.cell(r, 12).value) == 79 and sum(1 for r in range(2, 60) if wb['RevenueData'].cell(r, 1).value) == 17,
      (sum(1 for r in range(10, 89) if gi.cell(r, 12).value), sum(1 for r in range(2, 60) if wb['RevenueData'].cell(r, 1).value)))
check('three navigation buttons: Summary, Typical Values, Method',
      all('HYPERLINK' in str(gi[c].value) for c in ('D44', 'D46', 'D48')), [gi[c].value for c in ('D44', 'D46', 'D48')])
check('how-to card and live status line present', str(gi['F2'].value or '').startswith('HOW TO USE') and 'Still needed' in str(gi['F9'].value))
check('input hints beside the parameters', sum(1 for r in (25, 26, 27, 28, 33, 34, 36, 37, 38) if gi.cell(r, 6).value) == 9)

print(); print('=' * 78); print('SPELLING IN THE ISSUED TEXT'); print('=' * 78)
ov = rd('xl/drawings/drawing2.xml'); ss = rd('xl/sharedStrings.xml')
for bad, good, part in [('Adminimstration', 'Administration', ov), ('clossures', 'closures', ov),
                        ('associeted', 'associated', ov), ('Intial Construction', 'Initial Construction', ss)]:
    check('%s corrected to %s' % (bad, good), bad not in part and good in part)


print(); print('=' * 78); print('SETUP FLOW AND THE REFERENCE SHEETS'); print('=' * 78)
card = str(wb['General Information']['F3'].value)
check('the how-to card names all five steps in order',
      [card.find(t) for t in ('1.', '2.', '3.', '4.', '5.')] == sorted(card.find(t) for t in ('1.', '2.', '3.', '4.', '5.'))
      and all(t in card for t in ('Overview', 'General Information', 'Pay_Items', 'Alternative Setup', 'Summary')),
      card[:80])
check('five navigation buttons on General Information, one per destination',
      all('HYPERLINK' in str(gi[c].value) for c in ('D44', 'D46', 'D48', 'D50', 'D52')),
      [str(gi[c].value)[:40] for c in ('D44', 'D46', 'D48', 'D50', 'D52')])
steps = {'Overview': wb['Overview']['G1'].value, 'Instructions': wb['Instructions']['G1'].value,
         'Pay_Items': wb['Pay_Items']['C1'].value, 'TMP(NewHMA)': wb['TMP(NewHMA)']['D1'].value,
         'Summary': ws.cell(HDR - 1, 7).value}
check('every sheet in the flow says which step it is',
      [str(v).strip()[:11] for v in steps.values()] == ['STEP 1 of 5', 'STEP 1 of 5', 'STEP 3 of 5', 'STEP 4 of 5', 'STEP 5 of 5'],
      {k: str(v).strip()[:14] for k, v in steps.items()})
for name in ('Overview', 'Instructions', 'Pay_Items', 'Maintenance Policies'):
    sh = wb[name]
    b = [c for c in ('A1', 'B1') if 'HYPERLINK' in str(sh[c].value)]
    check('%s carries the navigation row' % name, len(b) >= 1, [sh['A1'].value, sh['B1'].value])
pi = wb['Pay_Items']
check('Pay_Items marks the unit-cost column as the input',
      pi['F4'].fill.fgColor.rgb.endswith('D9D9D9') and pi['F4'].number_format.startswith('"$"'),
      (pi['F4'].fill.fgColor.rgb, pi['F4'].number_format))
priced = [r for r in range(4, 60) if pi.cell(r, 6).value not in (None, '')]
check('Pay_Items keeps all 56 items, the 27 that carry a cost and the three that are formulas',
      sum(1 for r in range(4, 60) if pi.cell(r, 4).value) == 56 and len(priced) == 27
      and [pi.cell(r, 6).value for r in (10, 16, 56)] == ['=7/9', '=10/2', '=6/9'],
      (len(priced), [pi.cell(r, 6).value for r in (10, 16, 56)]))
check('its Table2 is still declared over C2:I59', 'ref="C2:I59"' in rd('xl/tables/table5.xml'))
check('Pay_Items says what the empty division columns mean and how many items carry no cost',
      'West' in str(pi['A62'].value) and 'no unit cost' in str(pi['A63'].value), str(pi['A63'].value)[:50])
check('Pay_Items part headings band the left of the table',
      all(pi.cell(r, 1).font.bold for r in (4, 26, 34, 37, 44, 48, 54)))
check('Pay_Items freezes the header and repeats the band and the header in print',
      pi.freeze_panes == 'A3' and 'Pay_Items!$1:$2' in rd('xl/workbook.xml'), pi.freeze_panes)
mp = wb['Maintenance Policies']
check('Maintenance Policies says it is live and what Rate and Year Applied mean, in one note',
      'live' in str(mp['C3'].value) and 'Rate' in str(mp['C3'].value)
      and mp['C4'].value is None and mp['C5'].value is None,
      str(mp['C3'].value)[:70])
# the sheet is read by the templates, so the note must not tell anyone otherwise
mp_refs = set()
for p_ in names:
    if not re.match(r'xl/worksheets/sheet\d+\.xml$', p_): continue
    mp_refs |= set(re.findall(r"Maintenance Policies'!\$?([DE])\$?\d+", rd(p_)))
check('and that is true: the templates read both its Rate and Year Applied columns',
      mp_refs == {'D', 'E'} and 'changes nothing' not in str(mp['C3'].value), sorted(mp_refs))
check('the four maintenance tables carry a header band',
      all(mp.cell(r, 2).font.color and mp.cell(r, 2).font.color.rgb.endswith('FFFFFF') for r in (9, 36, 50, 75)),
      [mp.cell(r, 2).value for r in (9, 36, 50, 75)])


print(); print('=' * 78); print('THE TDOT MARK'); print('=' * 78)
LOGO_SHEETS = ['xl/drawings/drawing%d.xml' % i for i in range(1, 14)]
anchors, aspects = {}, set()
for n in LOGO_SHEETS:
    if n not in names: continue
    d = rd(n)
    m = re.search(r'<xdr:oneCellAnchor><xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>(\d+)</xdr:colOff>'
                  r'<xdr:row>(\d+)</xdr:row><xdr:rowOff>(\d+)</xdr:rowOff></xdr:from>'
                  r'<xdr:ext cx="(\d+)" cy="(\d+)"/>(?:(?!</xdr:oneCellAnchor>).)*?TDOT logo', d, re.S)
    if m:
        col, off, row, roff, cx, cy = (int(v) for v in m.groups())
        anchors[n] = (col, row)
        aspects.add(round(cx / float(cy), 3))
check('the mark is on all thirteen sheets a user can reach', len(anchors) == 13, sorted(anchors))
check('every one of them sits in row 1', all(r == 0 for _, r in anchors.values()), sorted(set(r for _, r in anchors.values())))
check('and at the artwork aspect, not stretched', aspects == {2.288}, aspects)
check('no placement is left at the old B2 anchor',
      not any('<xdr:pic>' in rd(n).split('<xdr:oneCellAnchor>')[0] for n in anchors), '')
rows1 = {sh.title: wb[sh.title].row_dimensions[1].height for sh in wb.worksheets
         if sh.sheet_state == 'visible' or sh.title.startswith('TMP(')}
check('row 1 is the same height on every one of them',
      sorted(set(v for k, v in rows1.items() if k != 'TMP(PCCRehab)')) == [40.0], rows1)
titles = re.findall(r'<definedName name="_xlnm.Print_Titles" localSheetId="(\d+)">([^<]*)</definedName>', rd('xl/workbook.xml'))
check('row 1 repeats at the top of every printed page, so the mark prints throughout',
      len(titles) == 13 and all(t.endswith('$1:$1') or t.endswith('$1:$2') for _, t in titles),
      [t for _, t in titles][:3])
check('Pay_Items repeats its table header as well as the band',
      any(t.endswith("Pay_Items!$1:$2") for _, t in titles), [t for _, t in titles if 'Pay_Items' in t])


print(); print('=' * 78); print('THE PAY ITEM PICKER AND THE PRICING SOURCE'); print('=' * 78)
import struct
ALT = ['TMP(NewHMA)', 'TMP(NewPCC)', 'TMP(HMARehab)', 'TMP(NewHMA)_IndirectCost', 'TMP(NewPCC)_IndirectCost']
combos = [n for n in names if re.match(r'xl/activeX/activeX\d+\.bin$', n)
          and z.read(n)[:16] == bytes.fromhex('301dd28b42ecce119e0d00aa006002f3')]
widths, masked = set(), 0
for n in combos:
    d = z.read(n)
    if struct.unpack('<Q', d[20:28])[0] & (1 << 10):
        masked += 1
        widths.add(struct.unpack('<i', d[36:40])[0])
check('all 105 pay item pickers carry a ListWidth', len(combos) == 105 and masked == 105, (len(combos), masked))
check('and it is wide enough for the longest description at two columns',
      widths == {22860}, [round(w / 2540.0, 2) for w in widths])
check('every picker still lists the pay item number beside the description',
      all(rd(p).count('listFillRange="Pay_Items!C3:D59"') == rd(p).count('<controlPr ')
          for p in names if re.match(r'xl/worksheets/sheet\d+\.xml$', p) and 'listFillRange' in rd(p)))
widecol = {}
for t in ALT:
    sh = wb[t]
    widecol[t] = round(sh.column_dimensions['C'].width, 2)
check('the picker column is wide enough to read what was picked', set(widecol.values()) == {53.71}, widecol)
anch = set()
for p in names:
    if not re.match(r'xl/worksheets/sheet\d+\.xml$', p): continue
    for m in re.finditer(r'<from><xdr:col>2</xdr:col>.*?</from><to><xdr:col>(\d+)</xdr:col><xdr:colOff>(\d+)</xdr:colOff>', rd(p), re.S):
        anch.add(m.groups())
check('and each picker ends exactly at that column\'s edge', anch == {('3', '0')}, anch)

src = {t: (wb[t]['C11'].value, wb[t]['B11'].value, wb[t]['I11'].value) for t in ALT}
check('every alternative worksheet picks a pricing source, defaulting to Regular',
      all(v[0] == 'Regular' and v[1] == 'Price from:' for v in src.values()), src)
check('and resolves it to a column number, falling back to Regular when the cell is empty',
      all(v[2] == '=IFERROR(MATCH($C$11,PriceSources,0),1)' for v in src.values()))
dv = [rd('xl/worksheets/sheet%d.xml' % i) for i in (1, 8, 10, 11, 12)]
check('C11 is a list, not free text',
      all('sqref="C11"' in x and '&quot;Regular,Middle,West,East&quot;' in x for x in dv))
nm = dict(re.findall(r'<definedName name="(PriceSources|UnitCostGrid|PayItemKeys)">([^<]*)</definedName>', rd('xl/workbook.xml')))
check('the three names those lookups read are defined', len(nm) == 3, nm)
check('the grid is the four unit cost columns of Table2', nm.get('UnitCostGrid') == 'Table2[[Unit Cost]:[East]]')
TPL = ['xl/worksheets/sheet%d.xml' % i for i in (1, 8, 10, 11, 12)]
old = sum(rd(p).count('Table2[Pay Item Description],Table2[Unit Cost]') for p in TPL)
new = sum(rd(p).count('"N/A",IF(INDEX(UnitCostGrid,MATCH(') for p in TPL)
flags = sum(rd(p).count('="","",IF(INDEX(UnitCostGrid,MATCH(') for p in TPL)
check('all 156 unit cost lookups on the five templates read the chosen column',
      new == 156 and old == 0, (new, old))
check('and each of the fifty pay item lines flags whether that division prices it', flags == 50, flags)
check('Pay_Items marks Middle, West and East as fillable',
      all(wb['Pay_Items'].cell(r, c).fill.fgColor.rgb == 'FFEDEDED' for r in (3, 30, 59) for c in (7, 8, 9)))
check('TMP(HMARehab) hides its working columns like every other template',
      all(wb['TMP(HMARehab)'].column_dimensions[c].hidden for c in ('I', 'L', 'BL')))
check('and its chart still plots them', 'plotVisOnly val="0"' in rd('xl/charts/chart5.xml'))
gi = wb['General Information']
check('the always-blank Airport Owner row now carries the county',
      gi['C12'].value == 'County:' and gi['D12'].value.endswith('Table17[],4))'), (gi['C12'].value, gi['D12'].value))
check('and State Region below it is unmoved, so the Summary still finds it', gi['C13'].value == 'State Region:')

print(); print('=' * 78); print('SALVAGE'); print('=' * 78)
mp = wb['Maintenance Policies']
check('the HMA salvage fraction is a formula, not a constant',
      str(mp['D32'].value).startswith('=IF(E22>'), mp['D32'].value)
check('and it divides remaining life by an expected life you can see and change',
      mp['F32'].value == 16 and mp['F46'].value == 40 and mp['F9'].value == 'Asset life (yrs)',
      (mp['F32'].value, mp['F46'].value))
check('the PCC salvage fraction is a formula too',
      str(mp['D46'].value).startswith('=MAX(0,MIN(1,(0+F46'), mp['D46'].value)
check('both read the analysis period rather than assuming 30 years',
      all("'General Information'!$D$33" in str(mp[r].value) for r in ('D32', 'D46', 'E32', 'E46')))
check('the HMA fraction is guarded for a period that ends before the overlay',
      "IF(E22>'General Information'!$D$33,0," in str(mp['D32'].value))
check('the sentence on each salvage row restates itself from the numbers',
      all(str(mp[r].value).startswith('="Salvage value: "&TEXT(') for r in ('C32', 'C46')))
check('the two rehabilitation tables keep their stated zero salvage',
      mp['D71'].value == 0 and mp['D85'].value == 0 and
      'zero salvage' in str(mp['C71'].value) and 'zero salvage' in str(mp['C85'].value))
check('every salvage row shows the year the credit is actually taken',
      all(str(mp['E%d' % r].value) == "='General Information'!$D$33" for r in (32, 46, 71, 85)))
check('the alternative templates still multiply that fraction by the salvaged cost',
      all("-'Maintenance Policies'!D32*AV22" in rd(p_) for p_ in
          ('xl/worksheets/sheet8.xml', 'xl/worksheets/sheet10.xml')))

print(); print('=' * 78); print('THE FIRST SHEET, AND WHAT OPENS WHERE'); print('=' * 78)
gi = wb['General Information']
check('no sheet opens scrolled away from its own top-left',
      not [n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n)
           and 'topLeftCell' in (re.search(r'<sheetView\b[^>]*?(/>|>)', rd(n)) or type('', (), {'group': lambda *a: ''})()).group(0)],
      [n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n)
       and 'topLeftCell' in (re.search(r'<sheetView\b[^>]*?(/>|>)', rd(n)) or type('', (), {'group': lambda *a: ''})()).group(0)])
check('a frozen row split is a bottom-LEFT pane, which is what Excel honours',
      all('activePane="bottomLeft"' in m for m in re.findall(r'<pane\b[^>]*xSplit="0"[^>]*/>|<pane\b(?![^>]*xSplit)[^>]*/>',
                                                             ''.join(rd(n) for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n)))))
check('the how-to card and the status line have room to render',
      all(round(gi.column_dimensions[c].width, 1) == 13.0 for c in 'FGHIJ')
      and all(gi.row_dimensions[r].height == 22 for r in range(3, 8)),
      {c: gi.column_dimensions[c].width for c in 'FGHIJ'})
check('the step line moved off row 1 to make room for the button',
      gi['B1'].value is None and str(gi['C1'].value).startswith('STEP 2 of 5'))
check('the New Study button is on the sheet and wired to the macro',
      'macro="NewStudy"' in rd('xl/drawings/drawing4.xml') and 'btnNewStudy' in rd('xl/drawings/drawing4.xml'))
check('the New Study caption is short and says nothing about importing',
      'import' not in str(gi['B2'].value).lower() and len(str(gi['B2'].value)) < 90,
      str(gi['B2'].value))
check('the free-text inputs carry a border so they stop reading as one grey block',
      all(gi.cell(r, 4).border.left.style for r in list(range(14, 18)) + list(range(21, 25))))
check('the salvage note is out of the input column', gi['D35'].value is None and 'Remaining service life' in str(gi['F35'].value))
check('the pavement section waits for a mainline area rather than dividing by zero',
      all("N('General Information'!$D$26)<=0" in str(ws.cell(r, c).value)
          for r in range(SEC_T + 3, SEC_T + 7) for c in (8, 9, 10, 13, 16)))
# the locator map is excluded on purpose: its axes carry no tick labels (numFmt ';;;') and forcing a
# plot-area layout on a geographic scatter would stretch Tennessee out of shape
dash = [n for n in charts if 'Summary' in rd(n) and '<valAx>' in rd(n) and 'scatterChart' not in rd(n)]
check('every dashboard chart reserves its margins so axis titles clear the tick labels',
      len(dash) >= 6 and all('manualLayout' in rd(n) for n in dash),
      [n for n in dash if 'manualLayout' not in rd(n)] or len(dash))
check('and no legend is drawn over its plot',
      all('<overlay val="0"/>' in rd(n) for n in charts if '<legend>' in rd(n)),
      [n for n in charts if '<legend>' in rd(n) and '<overlay val="0"/>' not in rd(n)])

print(); print('=' * 78)
print('%d checks, %d failed' % (checks, len(fails)))
print('ALL PASS' if not fails else 'FAILED: ' + '; '.join(fails))
sys.exit(1 if fails else 0)
