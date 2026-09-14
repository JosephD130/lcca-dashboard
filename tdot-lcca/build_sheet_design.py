#!/usr/bin/env python3
"""Build the canvas design into the workbook, sheet by sheet.

The canvas at design/canvas/ draws every sheet on one 1,240 px band so a sheet fits a 15 or 17
inch laptop. This pass lays that out in Excel: a twelve-column grid, the navy header band the
sheets already carry, white cards with a thin rule, and the same colour rules throughout (grey
means you type here, blue means calculated, green means the good outcome, amber means something
needs attention).

Nothing here changes a calculation. No row or column is inserted, so every formula, defined
name and table range stays where it was; the passes write text, styles, widths and merges only.

    python3 build_sheet_design.py <workbook.xlsm>
"""
import os, re, sys, html, shutil, zipfile
import xml.dom.minidom as minidom
import sheetdesign as D
from sheetdesign import (DXF_FILL, DXF_FONT, NAVY, BLUE, PALE, WHITE, PAPER, LINE, GREY, MUTED, INPUT, GREEN,
                         GREEN_SOFT, AMBER, AMBER_SOFT, RED, RED_SOFT, FONT, FILL, BORDER, ALIGN)

# The band: twelve columns of 103 px after a narrow gutter, so the content is 1,236 px wide.
GRID = {1: 1.5}
GRID.update({i: 14.71 for i in range(2, 14)})        # B .. M
BAND_RIGHT = 13                                      # column M


# ---------------------------------------------------------------------------- Overview, Instructions
# (label, link target, how many of the twelve columns the pill spans)
FLOW = [('1  Overview & Instructions', 'Instructions!B1', 3),
        ('2  General Information', "'General Information'!D9", 2),
        ('3  Pay Items', 'Pay_Items!A1', 2),
        ('4  Alternatives', "'General Information'!D9", 2),
        ('5  Summary', 'Summary!G1', 3)]

OVERVIEW_CARDS = [
    ('What you provide', ['The airport and the runway', 'Pavement areas and markings',
                          'Two or more alternatives', 'Quantities per pay item']),
    ('What the workbook adds', ['Unit costs by division', 'Maintenance and rehabilitation timing',
                                'Salvage on remaining life', 'Lost revenue during closures']),
    ('What you get', ['Net present worth per option', 'Equivalent annual cost',
                      'Sensitivity to the discount rate', 'Eight charts and a comparison']),
]

INSTRUCTION_STEPS = [
    ('Set the study up', [
        ('1', 'Fill in General Information',
         'Pick the airport from the dropdown and the ID, city, county and region fill themselves. '
         'Then the runway areas and the LCCA parameters, D9 to D39. The amber banner lists whatever '
         'is still missing.'),
        ('2', 'Check the unit costs on Pay_Items',
         'Every alternative is priced from that sheet, and the grey Unit Cost column is the one to '
         'edit. An item that ships with no cost prices at zero.'),
        ('3', 'Review the maintenance policies',
         'Four tables set what each pavement type gets done to it and in which year. They are live: '
         'a rate or a year changed there moves every alternative of that type.')]),
    ('Run it', [
        ('4', 'Add each alternative',
         'Back on General Information, click Alternative Setup. Name it, choose the pavement type, '
         'and the form builds a worksheet for it. Enter quantities against the pay items there.'),
        ('5', 'Read the Summary',
         'It updates by itself. The banner names the lowest present worth; the two key plots show '
         'why it wins and whether it still wins at other discount rates.'),
        ('6', 'Start the next study',
         'The + New Study button saves a cleared copy wherever you choose, with every input blank '
         'and every alternative removed. The study you have open is untouched.')]),
]

INSTRUCTION_FOOTER = [
    ('Grey means yours', ['A grey cell is one you type in.', 'Everything else calculates: leave it alone.']),
    ('Where to look things up', ['Typical Values: the usual ranges, with sources.',
                                 'Method: every formula as it stands in the sheet.']),
    ('If a number looks wrong', ['Check the Unit Cost column first, then the year a policy is applied.',
                                 'Both feed every alternative of that pavement type at once.']),
]


def pill_merges():
    out, c0 = [], 2
    for _, _, span in FLOW:
        out.append('%s19:%s19' % (D.colname(c0), D.colname(c0 + span - 1)))
        c0 += span
    return out


def text_sheets(rd, wr, st):
    """Overview and Instructions: the wall of text becomes a card layout, and the original prose
    moves below it under its own heading rather than being lost."""
    s_sub = st.add(font=FONT(9, i=True, color=GREY), alignment=ALIGN())
    s_title = st.add(font=FONT(13, b=True), alignment=ALIGN(v='center'))
    s_kick = st.add(font=FONT(8, b=True, color=MUTED), alignment=ALIGN(v='bottom'))
    s_cardh = st.add(font=FONT(10, b=True, color=WHITE), fill=FILL(NAVY), alignment=ALIGN(indent=1))
    s_amberh = st.add(font=FONT(10, b=True, color=AMBER), fill=FILL(AMBER_SOFT),
                      border=BORDER(left=True, right=True, top=True, color='FFE6CF76'),
                      alignment=ALIGN(indent=1))
    s_pill = st.add(font=FONT(9, b=True, color=BLUE), fill=FILL(PALE),
                    border=BORDER(True, True, True, True, color='FFC6D9EC'), alignment=ALIGN(h='center'))
    s_pill1 = st.add(font=FONT(9, b=True, color=WHITE), fill=FILL(GREEN),
                     border=BORDER(True, True, True, True, color=GREEN), alignment=ALIGN(h='center'))
    s_num = st.add(font=FONT(14, b=True, color=BLUE), fill=FILL(PALE),
                   alignment=ALIGN(h='center', v='center'))
    s_seco = st.add(font=FONT(10, b=True), border=BORDER(bottom=True, color=BLUE), alignment=ALIGN(v='center'))

    def w_(**kw):
        """A style on the white card face, with the card's own edges where they fall."""
        return st.add(fill=FILL(WHITE), **kw)

    def card(x, r0, r1, c0, c1, title, head=None):
        """Header row r0, white face down to r1, thin rule around the whole block."""
        x = D.box(x, st, r0 + 1, r1, c0, c1, fill=WHITE)
        x = D.restyle(x, [r0], [D.colname(c) for c in range(c0, c1 + 1)], head or s_cardh)
        x = D.text(x, '%s%d' % (D.colname(c0), r0), title, head or s_cardh)
        return x

    def face(x, ref, value, size=9, color=MUTED, bold=False, wrap=True, top=False,
             left=False, bottom=False, indent=1, v='top'):
        return D.text(x, ref, value, w_(font=FONT(size, b=bold, color=color),
                                        border=BORDER(left=left, top=top, bottom=bottom),
                                        alignment=ALIGN(v=v, wrap=wrap, indent=indent)))

    # -------------------------------------------------------------- Overview
    x = rd('xl/worksheets/sheet2.xml')
    x = D.cols(x, GRID, default_style='5')
    for r in range(2, 28): x = D.ensure_row(x, r)
    # the address block that sat in rows 3-9 moves into a card at the foot of the page
    for r in range(3, 10): x = D.remove_cell(x, 'B%d' % r)
    x = x.replace('<hyperlink ref="B8"', '<hyperlink ref="B25"')

    x = D.text(x, 'G1', 'Overview', s_title)
    x = D.text(x, 'B2', 'STEP 1 of 5   \u00b7   What the framework does and the rules behind it. '
                        'Read this and Instructions, then go to General Information.', s_sub)
    x = D.row_height(x, 2, 16); x = D.row_height(x, 3, 6)

    # hero
    x = D.box(x, st, 4, 6, 2, BAND_RIGHT, fill=WHITE)
    x = face(x, 'B4', 'A life-cycle cost analysis compares alternatives over their whole life, '
                      'not their first bill.', size=14, color=NAVY, bold=True, wrap=False,
             top=True, left=True, v='center')
    x = face(x, 'B5', 'You describe two or more ways to build the same runway. The workbook prices each '
                      'one from the TDOT pay-item list, adds the maintenance and rehabilitation each will '
                      'need over the analysis period, discounts it all to today\u2019s money, and reports '
                      'which costs least in present worth.', size=10, left=True)
    for r, h in ((4, 26), (5, 17), (6, 17), (7, 8)): x = D.row_height(x, r, h)

    # three cards
    for i, (title, lines) in enumerate(OVERVIEW_CARDS):
        c0 = 2 + 4 * i
        x = card(x, 8, 12, c0, c0 + 3, title)
        for j, line in enumerate(lines):
            x = face(x, '%s%d' % (D.colname(c0), 9 + j), '\u2022  ' + line,
                     left=True, bottom=(9 + j == 12), wrap=False, v='center')
    x = D.row_height(x, 8, 20)
    for r in range(9, 13): x = D.row_height(x, r, 15)
    x = D.row_height(x, 13, 8)

    # the rules, read live out of General Information so they track the file
    x = D.text(x, 'B14', 'THE RULES BEHIND IT', s_kick); x = D.row_height(x, 14, 14)
    x = D.box(x, st, 15, 16, 2, BAND_RIGHT, fill=WHITE)
    rules = [("IF('General Information'!$D$33=\"\",\"30\",'General Information'!$D$33)&\" years\"",
              'analysis period, live from General Information D33'),
             ("IF('General Information'!$D$34=\"\",\"3\",'General Information'!$D$34)&\"%\"",
              'discount rate, D34; FAA follows OMB Circular A-94'),
             ("IF('General Information'!$D$36=\"\",\"10\",'General Information'!$D$36)&\" / \"&"
              "IF('General Information'!$D$37=\"\",\"5\",'General Information'!$D$37)&\"%\"",
              'mobilization and engineering, D36 and D37'),
             (None, 'salvage basis: remaining life over expected life (AAPTP, FAA)')]
    for i, (f, cap) in enumerate(rules):
        c0 = 2 + 3 * i
        s_v = w_(font=FONT(14, b=True), border=BORDER(left=(i == 0), top=True),
                 alignment=ALIGN(v='center', indent=1))
        if f: x = D.formula(x, '%s15' % D.colname(c0), f, s_v, t='str', v='')
        else: x = D.text(x, '%s15' % D.colname(c0), 'Remaining life', s_v)
        x = face(x, '%s16' % D.colname(c0), cap, size=8, bottom=True, left=(i == 0))
    x = D.row_height(x, 15, 22); x = D.row_height(x, 16, 20); x = D.row_height(x, 17, 8)

    # the five-step flow, each pill a link to its sheet
    x = D.text(x, 'B18', 'HOW A STUDY RUNS', s_kick); x = D.row_height(x, 18, 14)
    c0 = 2
    for i, (label, target, span) in enumerate(FLOW):
        s = s_pill1 if i == 0 else s_pill
        x = D.link(x, '%s19' % D.colname(c0), label, target, s)
        for c in range(c0 + 1, c0 + span): x = D.blank(x, '%s19' % D.colname(c), s)
        c0 += span
    x = D.row_height(x, 19, 20); x = D.row_height(x, 20, 8)

    # contact, and what to do before starting
    x = card(x, 21, 25, 2, 6, 'TDOT Aeronautics Division')
    for j, line in enumerate(['7335 Centennial Boulevard', 'Nashville, Tennessee 37209', '615-741-3208']):
        x = face(x, 'B%d' % (22 + j), line, wrap=False, v='center', left=True)
    x = D.put_cell(x, 'B25', '<c r="B25" s="%d" t="inlineStr"><is><t>tn.gov/tdot/aeronautics.html</t></is></c>'
                   % w_(font=FONT(9, color=BLUE), border=BORDER(left=True, bottom=True),
                        alignment=ALIGN(v='center', indent=1)))
    x = card(x, 21, 25, 8, BAND_RIGHT, 'Before you start', head=s_amberh)
    x = D.box(x, st, 22, 25, 8, BAND_RIGHT, fill=AMBER_SOFT, color='FFE6CF76')
    x = D.text(x, 'H22', 'Enable macros and ActiveX when Excel asks. The Alternative Setup form and the '
                         'pay-item pickers need them. If the yellow bar does not appear, check Trust Center '
                         '\u203a ActiveX Settings and restart Excel.',
               st.add(font=FONT(9, color=AMBER), fill=FILL(AMBER_SOFT),
                      border=BORDER(left=True, color='FFE6CF76'),
                      alignment=ALIGN(v='top', wrap=True, indent=1)))
    x = D.row_height(x, 21, 20)
    for r in range(22, 26): x = D.row_height(x, r, 15)
    x = D.row_height(x, 26, 10)

    x = D.text(x, 'B27', 'BACKGROUND', s_seco)
    x = D.restyle(x, [27], [D.colname(c) for c in range(3, BAND_RIGHT + 1)], s_seco)
    x = D.row_height(x, 27, 20)
    x = D.merges(x, ['B2:M2', 'B4:M4', 'B5:M6'] +
                 ['%s%d:%s%d' % (D.colname(2 + 4 * i), r, D.colname(5 + 4 * i), r)
                  for i in range(3) for r in range(8, 13)] +
                 ['B14:M14', 'B18:M18'] +
                 ['%s%d:%s%d' % (D.colname(2 + 3 * i), r, D.colname(4 + 3 * i), r)
                  for i in range(4) for r in (15, 16)] +
                 pill_merges() +
                 ['B21:F21', 'H21:M21', 'H22:M25', 'B27:M27'] +
                 ['B%d:F%d' % (r, r) for r in range(22, 26)])
    x = D.sheet_view(x, gridlines=False)
    x = D.page_setup(x)
    move_textbox('xl/worksheets/sheet2.xml', rd, wr, first_row=28)
    wr('xl/worksheets/sheet2.xml', x)

    # -------------------------------------------------------------- Instructions
    x = rd('xl/worksheets/sheet3.xml')
    x = D.cols(x, GRID, default_style='5')
    for r in range(2, 27): x = D.ensure_row(x, r)
    for r in range(3, 10): x = D.remove_cell(x, 'B%d' % r)

    x = D.text(x, 'G1', 'Instructions', s_title)
    x = D.text(x, 'B2', 'STEP 1 of 5   \u00b7   How to run an analysis, start to finish. When you are '
                        'ready, go to General Information and fill in the grey cells.', s_sub)
    x = D.row_height(x, 2, 16); x = D.row_height(x, 3, 6)

    x = D.box(x, st, 4, 5, 2, BAND_RIGHT, fill=AMBER_SOFT, color='FFE6CF76')
    x = D.text(x, 'B4', 'First, enable content.',
               st.add(font=FONT(11, b=True, color=AMBER), fill=FILL(AMBER_SOFT),
                      border=BORDER(left=True, top=True, color='FFE6CF76'),
                      alignment=ALIGN(v='center', indent=1)))
    x = D.text(x, 'B5', 'Macros and ActiveX must be on, or the Alternative Setup form and the pay-item '
                        'pickers will not run. Trust Center \u203a ActiveX Settings, then restart Excel.',
               st.add(font=FONT(9, color=AMBER), fill=FILL(AMBER_SOFT),
                      border=BORDER(left=True, bottom=True, color='FFE6CF76'),
                      alignment=ALIGN(v='center', indent=1)))
    x = D.row_height(x, 4, 18); x = D.row_height(x, 5, 16); x = D.row_height(x, 6, 8)

    # two cards of numbered steps: number cell, heading, then two rows of body
    for i, (title, steps) in enumerate(INSTRUCTION_STEPS):
        c0 = 2 + 6 * i
        x = card(x, 7, 19, c0, c0 + 5, title)
        for j, (n, head, body) in enumerate(steps):
            r = 8 + 4 * j
            x = D.text(x, '%s%d' % (D.colname(c0), r), n, s_num)
            x = D.text(x, '%s%d' % (D.colname(c0 + 1), r), head,
                       w_(font=FONT(10, b=True), alignment=ALIGN(v='center')))
            x = D.text(x, '%s%d' % (D.colname(c0 + 1), r + 1), body,
                       w_(font=FONT(9, color=MUTED), alignment=ALIGN(v='top', wrap=True)))
    for r in range(8, 20):
        x = D.row_height(x, r, 15 if (r - 8) % 4 == 0 else (6 if (r - 8) % 4 == 3 else 14))
    x = D.row_height(x, 7, 20); x = D.row_height(x, 20, 8)

    for i, (title, lines) in enumerate(INSTRUCTION_FOOTER):
        c0 = 2 + 4 * i
        x = card(x, 21, 23, c0, c0 + 3, title)
        for j, line in enumerate(lines):
            x = face(x, '%s%d' % (D.colname(c0), 22 + j), line, left=True, bottom=(22 + j == 23))
    x = D.row_height(x, 21, 20)
    for r in (22, 23): x = D.row_height(x, r, 22)
    for r in (24, 25): x = D.row_height(x, r, 5)

    x = D.text(x, 'B26', 'THE ORIGINAL STEP-BY-STEP NOTES', s_seco)
    x = D.restyle(x, [26], [D.colname(c) for c in range(3, BAND_RIGHT + 1)], s_seco)
    x = D.row_height(x, 26, 20)
    x = D.merges(x, ['B2:M2', 'B4:M4', 'B5:M5'] +
                 ['%s7:%s7' % (D.colname(2 + 6 * i), D.colname(7 + 6 * i)) for i in range(2)] +
                 ['%s%d:%s%d' % (D.colname(2 + 6 * i), 8 + 4 * j, D.colname(2 + 6 * i), 10 + 4 * j)
                  for i in range(2) for j in range(3)] +
                 ['%s%d:%s%d' % (D.colname(3 + 6 * i), 8 + 4 * j, D.colname(7 + 6 * i), 8 + 4 * j)
                  for i in range(2) for j in range(3)] +
                 ['%s%d:%s%d' % (D.colname(3 + 6 * i), 9 + 4 * j, D.colname(7 + 6 * i), 10 + 4 * j)
                  for i in range(2) for j in range(3)] +
                 ['%s%d:%s%d' % (D.colname(2 + 4 * i), r, D.colname(5 + 4 * i), r)
                  for i in range(3) for r in range(21, 24)] +
                 ['B26:M26'])
    x = D.sheet_view(x, gridlines=False)
    x = D.page_setup(x)
    move_textbox('xl/worksheets/sheet3.xml', rd, wr, first_row=27)
    wr('xl/worksheets/sheet3.xml', x)


def move_textbox(sheet_part, rd, wr, first_row):
    """Drop the original prose text box to `first_row` (1-based), keeping its height, and re-anchor
    the TDOT mark to the right edge of the band."""
    drawing = drawing_of(sheet_part, rd)
    d = rd(drawing)
    m = re.search(r'<xdr:twoCellAnchor[^>]*>.*?</xdr:twoCellAnchor>', d, re.S)
    if m:
        a = m.group(0)
        f = re.search(r'<xdr:from>.*?<xdr:row>(\d+)</xdr:row>.*?</xdr:from>', a, re.S)
        t = re.search(r'<xdr:to>.*?<xdr:row>(\d+)</xdr:row>.*?</xdr:to>', a, re.S)
        r0, r1 = int(f.group(1)), int(t.group(1))
        n0, n1 = first_row - 1, first_row - 1 + (r1 - r0)
        a2 = a[:f.start()] + re.sub(r'<xdr:row>\d+</xdr:row>', '<xdr:row>%d</xdr:row>' % n0, f.group(0)) + \
            a[f.end():t.start()] + re.sub(r'<xdr:row>\d+</xdr:row>', '<xdr:row>%d</xdr:row>' % n1, t.group(0)) + \
            a[t.end():]
        # the band is twelve columns now, so the box spans B .. M
        a2 = a2.replace('<xdr:from><xdr:col>1</xdr:col>', '<xdr:from><xdr:col>1</xdr:col>')
        a2 = re.sub(r'(<xdr:to><xdr:col>)\d+(</xdr:col>)', r'\g<1>13\g<2>', a2)
        d = d[:m.start()] + a2 + d[m.end():]
    px, dflt = D.col_px(GRID)
    col, off = D.anchor_right(px, dflt, BAND_RIGHT, 1046062)
    d = re.sub(r'(<xdr:oneCellAnchor><xdr:from><xdr:col>)\d+(</xdr:col><xdr:colOff>)\d+(</xdr:colOff>)',
               r'\g<1>%d\g<2>%d\g<3>' % (col, off), d)
    wr(drawing, d)



# ---------------------------------------------------------------------------- Pay_Items
# The table starts on row 2 and is referenced from 291 formulas, five defined names, Table2 and
# every ActiveX picker, so no row is inserted above it: the canvas's "29 without a unit cost"
# tile becomes a live status line in the band, and the amber flag moves onto the cells themselves,
# where it is what a user actually needs to see.
PAY_WIDTHS = {1: 22, 2: 16, 3: 14, 4: 64, 5: 7, 6: 11, 7: 10, 8: 10, 9: 10, 10: 18}
PAY_FIRST, PAY_LAST = 4, 59


def pay_items(rd, wr, st):
    x = rd('xl/worksheets/sheet5.xml')
    x = D.cols(x, PAY_WIDTHS, default_style='5')

    s_title = st.add(font=FONT(13, b=True), alignment=ALIGN(v='center'))
    s_step = st.add(font=FONT(9, i=True, color=GREY), alignment=ALIGN(v='center', wrap=True))
    s_stat = st.add(font=FONT(9, b=True, color=AMBER), fill=FILL(AMBER_SOFT),
                    border=BORDER(True, True, True, True, color='FFE6CF76'),
                    alignment=ALIGN(v='center', indent=1))
    s_hdr = st.add(font=FONT(10, b=True, color=WHITE), fill=FILL(NAVY), alignment=ALIGN(indent=1))
    s_hdrc = st.add(font=FONT(10, b=True, color=WHITE), fill=FILL(NAVY), alignment=ALIGN(h='center'))

    x = D.text(x, 'C1', 'Pay Items', s_title)
    x = D.formula(x, 'D1',
                  'IF(COUNTBLANK($F$%d:$F$%d)=0,"Every pay item carries a unit cost.",'
                  'COUNTBLANK($F$%d:$F$%d)&" of "&COUNTA($D$%d:$D$%d)&" pay items carry no unit cost, '
                  'flagged amber below.")'
                  % (PAY_FIRST, PAY_LAST, PAY_FIRST, PAY_LAST, PAY_FIRST, PAY_LAST),
                  s_stat, t='str', v='')
    x = D.text(x, 'E1', 'STEP 3 of 5   \u00b7   the unit costs every alternative is priced from. '
                        'The grey Unit Cost column is the one to edit.', s_step)
    for c in 'FGHI': x = D.blank(x, '%s1' % c, s_step)
    # the merged "Average Pay Item Unit Cost" label gave up row 1 to the status line; the three
    # division columns say it themselves instead
    x = re.sub(r'<mergeCell ref="G1:I1"/>', '', x)
    x = D.remove_cell(x, 'G1')
    for col, name in (('G', 'Middle avg'), ('H', 'West avg'), ('I', 'East avg')):
        x = D.text(x, '%s2' % col, name, s_hdrc)
    x = D.merges(x, ['E1:I1'])
    x = D.row_height(x, 1, 40)

    # the header reads as the band the other sheets use
    x = D.restyle(x, [2], ['A', 'B', 'C', 'D', 'J'], s_hdr)
    x = D.restyle(x, [2], ['E', 'F', 'G', 'H', 'I'], s_hdrc)
    x = D.row_height(x, 2, 22)

    # every empty Unit Cost is flagged where it is used
    dxf = st.dxf(font=DXF_FONT(b=True, color=AMBER), fill=DXF_FILL(AMBER_SOFT))
    cf = ('<conditionalFormatting sqref="F%d:F%d"><cfRule type="containsBlanks" dxfId="%d" '
          'priority="1"><formula>LEN(TRIM(F%d))=0</formula></cfRule></conditionalFormatting>'
          % (PAY_FIRST, PAY_LAST, dxf, PAY_FIRST))
    x = re.sub(r'<conditionalFormatting sqref="F\d+:F\d+">.*?</conditionalFormatting>', '', x, flags=re.S)
    x = D.insert_before(x, cf, ['dataValidations', 'hyperlinks', 'printOptions', 'pageMargins',
                                'drawing', 'legacyDrawing', 'tableParts'])

    # the note stack: the count now lives in the band, so row 63 says what to do about it instead
    x = D.text(x, 'A63', 'An item with no unit cost prices at $0. The Unit Cost column flags those in '
                         'amber: fill one in before entering a quantity against it.',
               st.add(font=FONT(9, color=AMBER), fill=FILL(AMBER_SOFT),
                      alignment=ALIGN(v='center', wrap=True, indent=1)))
    x = D.sheet_view(x, freeze='A3', gridlines=False)
    wr('xl/worksheets/sheet5.xml', x)
    relogo('xl/worksheets/sheet5.xml', rd, wr, PAY_WIDTHS, 10)


# ---------------------------------------------------------------------------- Maintenance Policies
# Where each table's rehabilitation and salvage rows sit. The years and rates are read by every
# alternative, so the rows are only restyled; not one of them moves.
POLICY_TABLES = [
    dict(title=8, header=9, first=10, last=32, rehab=[22], salvage=(32, 'green')),
    dict(title=35, header=36, first=37, last=46, rehab=[41], salvage=(46, 'green')),
    dict(title=49, header=50, first=51, last=71, rehab=[59], salvage=(71, 'amber')),
    dict(title=74, header=75, first=76, last=85, rehab=[80], salvage=(85, 'amber')),
]
POL_WIDTHS = {1: 3, 2: 29.43, 3: 56, 4: 13.71, 5: 13.71, 6: 13.71}


def formula_at(x, ref):
    m = D.CELL_RE(ref).search(x)
    return bool(m) and '<f>' in m.group(0)


def policies(rd, wr, st):
    x = rd('xl/worksheets/sheet7.xml')
    x = D.cols(x, POL_WIDTHS, default_style='5')

    s_title = st.add(font=FONT(13, b=True), alignment=ALIGN(v='center'))
    s_tbl = st.add(font=FONT(11, b=True), alignment=ALIGN(v='center'))
    s_live = st.add(font=FONT(9, color=NAVY), fill=FILL(PALE),
                    border=BORDER(left=True, color=BLUE, style='medium'),
                    alignment=ALIGN(v='center', wrap=True, indent=1))
    s_livem = st.add(font=FONT(9, color=NAVY), fill=FILL(PALE), alignment=ALIGN(v='center', wrap=True))
    s_note = st.add(font=FONT(9, i=True, color=GREY), alignment=ALIGN(v='center', wrap=True))
    s_hdr = st.add(font=FONT(10, b=True, color=WHITE), fill=FILL(NAVY), alignment=ALIGN(indent=1))
    s_hdrc = st.add(font=FONT(10, b=True, color=WHITE), fill=FILL(NAVY), alignment=ALIGN(h='center'))
    s_rate = st.add(font=FONT(10), fill=FILL(INPUT), border=BORDER(True, True, True, True, color='FFBFBFBF'),
                    alignment=ALIGN(h='right', v='center'))
    year_top, year_mid, year_mid_last = {}, {}, {}
    for key, fill, ink in (('plain', PALE, NAVY), ('rehab', RED_SOFT, RED),
                           ('green', GREEN_SOFT, GREEN), ('amber', AMBER_SOFT, AMBER)):
        year_top[key] = st.add(font=FONT(10, b=True, color=ink), fill=FILL(fill),
                               border=BORDER(left=True, right=True, top=True, color=ink),
                               alignment=ALIGN(h='center', v='center'))
        year_mid[key] = st.add(fill=FILL(fill), border=BORDER(left=True, right=True, color=ink))
        year_mid_last[key] = st.add(fill=FILL(fill),
                                    border=BORDER(left=True, right=True, bottom=True, color=ink))
    chips = {}
    for key, fill, ink in (('rehab', RED_SOFT, RED), ('green', GREEN_SOFT, GREEN), ('amber', AMBER_SOFT, AMBER)):
        chips[key] = (
            st.add(font=FONT(10, b=True, color=ink), fill=FILL(fill), alignment=ALIGN(v='center')),
            st.add(font=FONT(10, color=ink), fill=FILL(fill), alignment=ALIGN(v='center')),
            st.add(font=FONT(10, b=True, color=ink), fill=FILL(fill),
                   border=BORDER(True, True, True, True, color=ink), alignment=ALIGN(h='center', v='center')),
        )

    x = D.text(x, 'B2', 'Maintenance & Rehabilitation Policies', s_title)
    x = D.remove_cell(x, 'C2')
    x = D.text(x, 'B3', 'This sheet is live. Every alternative reads its Rate and Year Applied straight '
                        'from these four tables, so a year moved here moves that event on every '
                        'alternative of that type and every number on the Summary. The grey cells are '
                        'the ones to edit.', s_live)
    for c in 'CDEF': x = D.blank(x, '%s3' % c, s_livem)
    x = D.text(x, 'B4', 'Rate is the share of the quantity the activity covers: 1 is the whole mainline '
                        'area, the whole markings area or the whole joint length. Year Applied counts '
                        'years after construction, and an event past the analysis period drops to zero.',
               s_note)
    x = D.remove_cell(x, 'C4')
    x = D.text(x, 'B5', 'Table 1 drives a New HMA alternative, Table 2 a New PCC alternative, Table 3 an '
                        'HMA overlay and Table 4 a PCC rehabilitation.', s_note)
    x = D.remove_cell(x, 'C5')
    x = D.remove_cell(x, 'C6')
    # the legend, in the same colours the rows below carry
    x = D.text(x, 'B6', 'Rehabilitation', chips['rehab'][0])
    x = D.text(x, 'C6', 'Salvage credit', chips['green'][0])
    x = D.text(x, 'D6', 'No salvage', chips['amber'][0])
    x = D.text(x, 'E6', 'Closure days come from Typical Values, not this sheet.', s_note)
    for r, h in ((2, 22), (3, 40), (4, 28), (5, 16), (6, 18), (7, 6)): x = D.row_height(x, r, h)

    cols_ = ['B', 'C', 'D', 'E', 'F']

    def has_value(ref):
        # a formula cell may carry no cached <v> until Excel recalculates, so <f> counts too
        m = D.CELL_RE(ref).search(x)
        return bool(m) and any(t in m.group(0) for t in ('<f>', '<v>', '<is>'))

    for t in POLICY_TABLES:
        x = D.restyle(x, [t['title']], ['B'], s_tbl)
        x = D.restyle(x, [t['header']], ['B', 'C'], s_hdr)
        x = D.restyle(x, [t['header']], ['D', 'E', 'F'], s_hdrc)
        x = D.row_height(x, t['header'], 20)
        sr, kind = t['salvage']

        # an activity runs from the row carrying its Year Applied to the row before the next one
        starts = [r for r in range(t['first'], t['last'] + 1) if has_value('E%d' % r)]
        for i, r0 in enumerate(starts):
            r1 = (starts[i + 1] - 1) if i + 1 < len(starts) else t['last']
            tint = 'rehab' if r0 in t['rehab'] else (kind if r0 == sr else None)
            if tint:
                x = D.restyle(x, range(r0, r1 + 1), cols_, chips[tint][1])
                x = D.restyle(x, [r0], ['B'], chips[tint][0])
            for r in range(r0, r1 + 1):
                # the grey Rate column is the input; the salvage fraction is calculated, so it is not
                if has_value('D%d' % r) and not formula_at(x, 'D%d' % r) and r != sr:
                    x = D.restyle(x, [r], ['D'], s_rate)
            x = D.restyle(x, [r0], ['E'], year_top[tint or 'plain'])
            for r in range(r0 + 1, r1 + 1):
                x = D.restyle(x, [r], ['E'], (year_mid_last if r == r1 else year_mid)[tint or 'plain'])

    x = D.merges(x, ['B3:F3', 'B4:F4', 'B5:F5', 'E6:F6'])
    x = D.sheet_view(x, freeze='A8', gridlines=False)
    wr('xl/worksheets/sheet7.xml', x)
    relogo('xl/worksheets/sheet7.xml', rd, wr, POL_WIDTHS, 6)


def drawing_of(sheet_part, rd):
    """The drawing part a sheet points at, read out of its own rels."""
    rels = rd('xl/worksheets/_rels/%s.rels' % os.path.basename(sheet_part))
    m = re.search(r'Target="\.\./(drawings/drawing\d+\.xml)"', rels)
    return 'xl/' + m.group(1) if m else None


def relogo(sheet_part, rd, wr, widths, right_col):
    """Put the TDOT mark back at the right edge of the band after the widths change."""
    drawing = drawing_of(sheet_part, rd)
    if not drawing: return
    d = rd(drawing)
    px, dflt = D.col_px(widths)
    col, off = D.anchor_right(px, dflt, right_col, 1046062)
    d = re.sub(r'(<xdr:oneCellAnchor><xdr:from><xdr:col>)\d+(</xdr:col><xdr:colOff>)\d+(</xdr:colOff>)',
               r'\g<1>%d\g<2>%d\g<3>' % (col, off), d)
    wr(drawing, d)



# ---------------------------------------------------------------------------- alternative worksheets
# The five templates the Alternative Setup form copies. Their rows differ (the indirect-cost
# variants carry a lost-revenue row beside every event, and PCC has fewer events), so every block
# is found by what column A or B says rather than by a row number, and nothing is moved: the
# Summary reads these sheets by MATCHing "Total" in column A and "Net Present Worth" in column B.
ALT_SHEETS = ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet8.xml', 'xl/worksheets/sheet10.xml',
              'xl/worksheets/sheet11.xml', 'xl/worksheets/sheet12.xml', 'xl/worksheets/sheet13.xml']


def cell_text(x, ref, shared):
    m = D.CELL_RE(ref).search(x)
    if not m: return None
    c = m.group(0)
    head = c[:c.index('>') + 1]
    body = c[len(head):-4] if not c.endswith('/>') else ''
    if 't="s"' in head:
        v = re.search(r'<v>(\d+)</v>', body)
        return shared[int(v.group(1))] if v else None
    if 't="inlineStr"' in head or '<is>' in body:
        return html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', body, re.S))) or None
    if '<f>' in body: return None
    v = re.search(r'<v>(.*?)</v>', body, re.S)
    return html.unescape(v.group(1)) if v else None


def shared_strings(rd):
    ss = rd('xl/sharedStrings.xml')
    return [html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S)))
            for si in re.findall(r'<si>(.*?)</si>', ss, re.S)]


def alternatives(rd, wr, st):
    shared = shared_strings(rd)

    s_title = st.add(font=FONT(11, b=True), alignment=ALIGN(v='center'))
    s_step = st.add(font=FONT(9, i=True, color=GREY), alignment=ALIGN(v='center', wrap=True))
    s_ctxl = st.add(font=FONT(10, color=MUTED), fill=FILL(PALE),
                    border=BORDER(left=True, color=BLUE, style='medium'),
                    alignment=ALIGN(v='center', indent=1))
    s_ctxb = st.add(fill=FILL(PALE))
    s_ctxv = st.add(font=FONT(11, b=True), fill=FILL(PALE), alignment=ALIGN(v='center'))
    s_hdr = st.add(font=FONT(10, b=True, color=WHITE), fill=FILL(NAVY), alignment=ALIGN(indent=1))
    s_hdrc = st.add(font=FONT(10, b=True, color=WHITE), fill=FILL(NAVY), alignment=ALIGN(h='center'))
    s_roll = st.add(font=FONT(10, color=MUTED), alignment=ALIGN(v='center', indent=1))
    s_tot = st.add(font=FONT(11, b=True), border=BORDER(top=True, color=NAVY, style='medium'),
                   alignment=ALIGN(v='center', indent=1))
    s_totv = st.add(font=FONT(13, b=True, color=BLUE), border=BORDER(top=True, color=NAVY, style='medium'),
                    alignment=ALIGN(h='right', v='center'), numfmt='&quot;$&quot;#,##0')
    s_npwt = st.add(font=FONT(13, b=True), alignment=ALIGN(v='center'))
    s_rehab = st.add(font=FONT(10, b=True, color=RED), fill=FILL(RED_SOFT), alignment=ALIGN(v='center', indent=1))
    s_rehaby = st.add(font=FONT(10, b=True, color=RED), fill=FILL(RED_SOFT), alignment=ALIGN(h='center', v='center'))
    s_rehabc = st.add(font=FONT(10, color=RED), fill=FILL(RED_SOFT), alignment=ALIGN(h='right', v='center'),
                      numfmt='&quot;$&quot;#,##0')
    s_salv = st.add(font=FONT(10, b=True, color=GREEN), fill=FILL(GREEN_SOFT), alignment=ALIGN(v='center', indent=1))
    s_salvy = st.add(font=FONT(10, b=True, color=GREEN), fill=FILL(GREEN_SOFT), alignment=ALIGN(h='center', v='center'))
    s_salvc = st.add(font=FONT(10, color=GREEN), fill=FILL(GREEN_SOFT), alignment=ALIGN(h='right', v='center'),
                     numfmt='&quot;$&quot;#,##0')
    s_ind = st.add(font=FONT(9, i=True, color=MUTED), alignment=ALIGN(v='center', indent=2))
    s_end = st.add(font=FONT(12, b=True, color=WHITE), fill=FILL(NAVY), alignment=ALIGN(v='center', indent=1))
    s_endb = st.add(fill=FILL(NAVY))
    s_endv = st.add(font=FONT(14, b=True, color=WHITE), fill=FILL(NAVY),
                    alignment=ALIGN(h='right', v='center'), numfmt='&quot;$&quot;#,##0')

    for part in ALT_SHEETS:
        try: x = rd(part)
        except FileNotFoundError: continue
        A = {r: cell_text(x, 'A%d' % r, shared) for r in range(1, 60)}
        B = {r: cell_text(x, 'B%d' % r, shared) for r in range(1, 60)}
        find_a = lambda t: next((r for r, v in A.items() if v == t), None)
        find_b = lambda t: next((r for r, v in B.items() if v == t), None)

        name_row = next((r for r in sorted(A) if A.get(r - 1) == 'Net Present Worth' and A.get(r)), None)
        item_hdr = find_a('Item No.')
        subtotal, total = find_a('Subtotal'), find_a('Total')
        sched_hdr = find_b('Item')
        npw_row = find_b('Net Present Worth')
        salvage = find_b('Salvage')

        # the band names the alternative, and the guide line carries the step marker
        if name_row:
            x = D.formula(x, 'D1', '$A$%d' % name_row, s_title, t='str', v='')
        guide = cell_text(x, 'A3', shared)
        if guide and not guide.startswith('STEP'):
            x = D.text(x, 'A3', 'STEP 4 of 5   \u00b7   ' + guide, s_step)

        # the context strip: what this alternative is being priced against
        ctx = [r for r in range(4, 11) if A.get(r) and str(A[r]).endswith(':')]
        if ctx:
            r0, r1 = min(ctx), max(ctx)
            for r in range(r0, r1 + 1):
                x = D.restyle(x, [r], ['A'], s_ctxl)
                x = D.restyle(x, [r], ['B'], s_ctxb)
                m = D.CELL_RE('C%d' % r).search(x)
                x = D.restyle(x, [r], ['C'], s_ctxv) if m else x
                x = D.restyle(x, [r], ['D'], s_ctxb)

        if item_hdr:
            x = D.restyle(x, [item_hdr], ['A', 'B', 'C'], s_hdr)
            x = D.restyle(x, [item_hdr], ['D', 'E', 'F', 'G'], s_hdrc)
            x = D.row_height(x, item_hdr, 20)

        # the cost rollup ends in one figure
        if subtotal and total:
            for r in range(subtotal, total):
                x = D.restyle(x, [r], ['A', 'B'], s_roll)
            x = D.restyle(x, [total], ['A', 'B'], s_tot)
            for c in 'CDEF': x = D.restyle(x, [total], [c], s_tot)
            x = D.restyle(x, [total], ['G'], s_totv)
            x = D.row_height(x, total, 22)

        if name_row:
            x = D.restyle(x, [name_row], ['A'], s_npwt)
            x = D.row_height(x, name_row, 20)
        if sched_hdr:
            x = D.restyle(x, [sched_hdr], ['B'], s_hdr)
            x = D.restyle(x, [sched_hdr], ['C', 'D', 'E'], s_hdrc)
            x = D.row_height(x, sched_hdr, 20)

        # the schedule: the rehabilitation stands out, the salvage credit reads as a credit, and the
        # bottom line is the number the Summary quotes
        if sched_hdr and npw_row:
            for r in range(sched_hdr + 1, npw_row):
                label = B.get(r) or ''
                if label.endswith('Indirect Cost'):
                    x = D.restyle(x, [r], ['B'], s_ind)
                elif label.startswith('Rehabilitation'):
                    x = D.restyle(x, [r], ['B'], s_rehab)
                    x = D.restyle(x, [r], ['C'], s_rehaby)
                    for c in 'DE': x = D.restyle(x, [r], [c], s_rehabc)
                elif label == 'Salvage':
                    x = D.restyle(x, [r], ['B'], s_salv)
                    x = D.restyle(x, [r], ['C'], s_salvy)
                    for c in 'DE': x = D.restyle(x, [r], [c], s_salvc)
            x = D.restyle(x, [npw_row], ['B'], s_end)
            for c in 'CD': x = D.restyle(x, [npw_row], [c], s_endb)
            # the Total row's own currency cells keep their format too
            x = D.restyle(x, [npw_row], ['E'], s_endv)
            x = D.row_height(x, npw_row, 26)

        x = D.sheet_view(x, gridlines=False)
        wr(part, x)


# ---------------------------------------------------------------------------- the two reference sheets
# Typical Values and Method. Both are long sourced tables, so the work is making their structure
# readable at a glance: a reference-only band at the top, numbered section headers, the same navy
# column headers the rest of the workbook uses, and widths that fit the band.
REF_SHEETS = [
    dict(part='xl/worksheets/sheet15.xml', title='Typical Values', last_col='G',
         widths={1: 38, 2: 20, 3: 16, 4: 26, 5: 10, 6: 32, 7: 34},
         step='reference   \u00b7   what a reasonable number looks like, with its source',
         banner='Reference only: nothing here feeds the calculation. Six sourced sections, '
                'covering the system, runway geometry, unit costs, closure days, the economic '
                'parameters and the timing of each treatment.'),
    dict(part='xl/worksheets/sheet16.xml', title='Method', last_col='E',
         widths={1: 30, 2: 22, 3: 46, 4: 52, 5: 26},
         step='reference   \u00b7   every formula as it stands in the sheet, and why',
         banner='Reference only: nothing here feeds the calculation. Eleven sections, each '
                'row giving the cell a formula lives in, the formula itself, what it means and the '
                'rule it comes from.',
         chain='Quantity \u00d7 unit cost   \u203a   + mobilization & engineering   \u203a   policy year '
               'for each event   \u203a   + closure days \u00d7 daily revenue   \u203a   \u2212 salvage   '
               '\u203a   discount to today = net present worth'),
]


def reference_sheets(rd, wr, st):
    shared = shared_strings(rd)
    s_title = st.add(font=FONT(13, b=True), alignment=ALIGN(v='center'))
    s_step = st.add(font=FONT(9, i=True, color=GREY), alignment=ALIGN(v='center'))
    s_band = st.add(font=FONT(9, color=NAVY), fill=FILL(PALE),
                    border=BORDER(left=True, color=BLUE, style='medium'),
                    alignment=ALIGN(v='center', wrap=True, indent=1))
    s_bandm = st.add(font=FONT(9, color=NAVY), fill=FILL(PALE), alignment=ALIGN(v='center', wrap=True))
    s_chain = st.add(font=FONT(9, b=True, color=WHITE), fill=FILL(NAVY),
                     alignment=ALIGN(h='center', v='center'))
    s_sec = st.add(font=FONT(11, b=True), border=BORDER(bottom=True, color=BLUE),
                   alignment=ALIGN(v='center'))
    s_secb = st.add(border=BORDER(bottom=True, color=BLUE))
    s_hdr = st.add(font=FONT(10, b=True, color=WHITE), fill=FILL(NAVY),
                   alignment=ALIGN(v='center', wrap=True, indent=1))
    s_foot = st.add(font=FONT(9, i=True, color=MUTED), alignment=ALIGN(v='center', wrap=True, indent=1))

    for cfg in REF_SHEETS:
        x = rd(cfg['part'])
        x = D.cols(x, cfg['widths'])
        last = cfg['last_col']
        cols_ = [D.colname(c) for c in range(1, D.colnum(last) + 1)]

        # A1 and B1 carry the navigation buttons, so the title starts at C1
        x = D.remove_cell(x, 'A3')
        x = D.text(x, 'C1', cfg['title'], s_title)
        x = D.text(x, 'D1', cfg['step'], s_step)
        x = D.text(x, 'A2', cfg['banner'], s_band)
        for c in cols_[1:]: x = D.blank(x, '%s2' % c, s_bandm)
        x = D.row_height(x, 1, 40); x = D.row_height(x, 2, 32); x = D.row_height(x, 3, 6)
        if cfg.get('chain'):
            x = D.text(x, 'A5', cfg['chain'], s_chain)
            for c in cols_[1:]: x = D.blank(x, '%s5' % c, s_chain)
            x = D.row_height(x, 5, 22)

        # a numbered section heading, then the row under it is that section's column header
        secs = [r for r in range(4, 130)
                if re.match(r'^\d+\. ', str(cell_text(x, 'A%d' % r, shared) or ''))]
        merged = set(int(m) for m in re.findall(r'<mergeCell ref="A(\d+):%s\d+"/>' % last, x))
        for r in secs:
            x = D.restyle(x, [r], ['A'], s_sec)
            for c in cols_[1:]: x = D.restyle(x, [r], [c], s_secb)
            x = D.row_height(x, r, 22)
            x = D.restyle(x, [r + 1], cols_, s_hdr)
            x = D.row_height(x, r + 1, 24)
        for r in sorted(merged):
            if r < 4 or r in secs: continue
            x = D.restyle(x, [r], cols_, s_foot)
            x = D.row_height(x, r, 26)

        x = D.merges(x, ['A2:%s2' % last] + (['A5:%s5' % last] if cfg.get('chain') else []))
        x = D.sheet_view(x, freeze='A4', gridlines=False)
        wr(cfg['part'], x)
        relogo(cfg['part'], rd, wr, cfg['widths'], D.colnum(last))

# ---------------------------------------------------------------------------- driver
PASSES = [('text_sheets', text_sheets), ('pay_items', pay_items), ('policies', policies),
          ('alternatives', alternatives),
          ('reference_sheets', reference_sheets)]


def main(path):
    work = path + '.design'
    if os.path.isdir(work): shutil.rmtree(work)
    with zipfile.ZipFile(path) as z: z.extractall(work)
    P = lambda p: os.path.join(work, p)
    read = lambda p: open(P(p), encoding='utf-8').read()

    def write(p, s):
        if p.endswith('.xml'): minidom.parseString(s)
        open(P(p), 'w', encoding='utf-8').write(s)

    st = D.Styles(work)
    for name, fn in PASSES:
        fn(read, write, st)
        print('  %s' % name)
    st.save()

    tmp = path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(P('[Content_Types].xml'), '[Content_Types].xml')
        for root, _, files in os.walk(work):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, work).replace(os.sep, '/')
                if rel == '[Content_Types].xml': continue
                z.write(full, rel)
    os.replace(tmp, path)
    shutil.rmtree(work)
    print('wrote', path)


if __name__ == '__main__':
    main(sys.argv[1])
