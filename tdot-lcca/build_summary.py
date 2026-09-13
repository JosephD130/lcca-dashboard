#!/usr/bin/env python3
"""Formula-driven Summary sheet for the TDOA LCCA framework, transplanted into the .xlsm package.

The sheet is generated with openpyxl in a scratch workbook (cells + charts), then its sheet XML,
styles, chart parts and drawing anchors are merged into the framework package so the existing VBA
(which manages columns A:E and "Chart 1") keeps working and nothing has to be imported by hand.
All blocks are driven by the hidden Database sheet (Alt worksheet names written by the Alternative
Setup form) through INDIRECT, so they follow whichever alternatives exist.

Layout (left to right, top to bottom, as a user reads it):
  A:E   the original table written by the Alternative Setup form (unchanged, VBA-managed)
  G:R   results table, verdict line, FAA 7% note
  G12   RealCost-style comparison block (agency / user / total PW and EUAC)
  G20+  five charts, two per row
  W+    chart data (calculated; labelled "do not edit")
"""
import os, re, zipfile, shutil, tempfile, html
import xml.dom.minidom as minidom
from openpyxl import Workbook
from openpyxl.formatting.rule import Rule, DataBarRule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, ScatterChart, Reference, Series
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
import geo_data
from openpyxl.utils import get_column_letter as L

GI = "'General Information'"
HMA_RGB, PCC_RGB, HMA2, PCC2 = '2A78D6', 'EB6834', '6DA7EC', 'F39C7A'
ALT_COLORS = [HMA_RGB, PCC_RGB, HMA2, PCC2]
NALT = 4          # four alternatives, as the existing Summary/VBA assume
# Row map of the visible sheet. The dashboard strip sits above the results, so every block below it moved
# down; anything that reads the Summary from outside (verify_build, kml_twin, LCCA_KML_Export.bas,
# examples/run_example) uses the same numbers.
KPI0 = 3          # KPI strip: tiles on rows 3-5 and 7-9, with row 6 and row 10 left as gaps
HDR = 11          # results table header
T0 = 12           # first alternative row (T0 .. T0+NALT-1)
VER, VER2 = 17, 18
CMP, CMPH, CMP0, CMPN = 20, 21, 22, 26
CHT, HOW, C1, C2, C3 = 28, 29, 30, 48, 66    # charts title, how-to-read, and the three chart rows
BMN = 87          # note under chart 8
SEC_T = 90        # pavement-section block: title, pcf row, header, four rows, note
CH67, NOTE67 = 100, 120
VBACH = 123       # where the chart the Alternative Setup form maintains is parked
T1 = T0 + NALT - 1
CMP1 = CMP0 + NALT - 1
R0, R1 = 36, 52   # NPW-table window on the alternative sheets (covers HMA and PCC layouts)
DC = 23           # first chart-data column (W)

def rng(alt_cell, col, a=R0 + 1, b=R1):
    return f'INDIRECT("\'"&{alt_cell}&"\'!${col}${a}:${col}${b}")'

def build_scratch(path):
    wb = Workbook(); ws = wb.active; ws.title = 'Summary'
    F_H = Font(name='Arial', size=10, bold=True); F_B = Font(name='Arial', size=10); F_T = Font(name='Arial', size=12, bold=True)
    F_N = Font(name='Arial', size=9, italic=True, color='595959'); F_BTN = Font(name='Arial', size=10, bold=True, color='FFFFFF'); F_DATA = Font(name='Arial', size=9, color='7F7F7F')
    FILL = PatternFill('solid', fgColor='D9D9D9'); FILL_BTN = PatternFill('solid', fgColor='1D2733'); FILL_BTN2 = PatternFill('solid', fgColor='2A78D6'); FILL_VERDICT = PatternFill('solid', fgColor='EAF2FB')
    thin = Side(style='thin', color='BFBFBF'); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
    CUR = '$#,##0;($#,##0);-'
    def hdr(row, col, labels, fill=FILL, font=F_H):
        for i, l in enumerate(labels):
            c = ws.cell(row, col + i, l); c.font = font; c.fill = fill; c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical='center')
    def button(ref, text, target, fill=FILL_BTN):
        c = ws[ref]; c.value = f'=HYPERLINK("#{target}","{text}")'; c.font = F_BTN; c.fill = fill; c.alignment = Alignment(horizontal='center', vertical='center'); c.border = BOX

    # ---- navigation buttons (macro-free hyperlinks) and the VBA-managed header, kept identical
    button('G1', '◄ General Information', "'General Information'!D9")
    button('H1', 'Instructions', 'Instructions!B9', FILL_BTN2)
    ws.row_dimensions[1].height = 22
    hdr(3, 1, ['Alternative', 'Name', 'Initial Construction', 'Present Worth', 'Alternative Description'])
    ws.column_dimensions['A'].width = 22; ws.column_dimensions['B'].width = 32; ws.column_dimensions['C'].width = 17; ws.column_dimensions['D'].width = 17; ws.column_dimensions['E'].width = 34
    ws.column_dimensions['F'].width = 3
    # A:E is the block the Alternative Setup form writes; everything in it is repeated with more detail from
    # column G onward, so it is hidden (not removed: the VBA still writes to it) and the charts line up with the table
    for c in 'ABCDEF': ws.column_dimensions[c].hidden = True
    for c in 'GHIJKLMNOPQRSTUV': ws.column_dimensions[c].width = 14
    ws.column_dimensions['G'].width = 20; ws.column_dimensions['H'].width = 24; ws.column_dimensions['P'].width = 30; ws.column_dimensions['Q'].width = 30

    # ---- 1. results table
    ws['J1'] = 'LCCA SUMMARY'; ws['J1'].font = F_T
    ws['G2'] = ('STEP 5 of 5: the comparison. Everything here calculates from the alternative worksheets and updates by itself. Columns A to F are hidden: the Alternative Setup form still writes '
                'its own small table and "Chart 1" there, and this table repeats all of it with more detail. Unhide A:F if you want to type an alternative description.'); ws['G2'].font = F_N
    hdr(HDR, 7, ['Worksheet', 'Alternative', 'Type', 'Initial construction', 'Maintenance PW', 'Rehabilitation PW', 'Lost revenue PW', 'Salvage PW', 'Net present worth', 'vs. lowest NPW', 'Closure days in period', 'Runway availability'])
    for i in range(NALT):
        r = T0 + i; dbr = 4 + i; g = f'$G{r}'
        # INDEX on whole columns: the VBA deletes Database rows when an alternative is removed, and a fixed
        # cell reference (Database!$D$4) would turn into #REF! in Excel; INDEX(...,row) survives the deletion.
        ws.cell(r, 7, f'=IF(INDEX(Database!$D:$D,{dbr})="","",INDEX(Database!$D:$D,{dbr}))')
        ws.cell(r, 8, f'=IF({g}="","",INDEX(Database!$A:$A,{dbr}))')
        ws.cell(r, 9, f'=IF({g}="","",INDEX(Database!$B:$B,{dbr}))')
        B, C, D, E = rng(g, 'B'), rng(g, 'C'), rng(g, 'D'), rng(g, 'E')
        ws.cell(r, 10, f'=IF({g}="","",IFERROR(INDEX(INDIRECT("\'"&{g}&"\'!$G$1:$G$70"),MATCH("Total",INDIRECT("\'"&{g}&"\'!$A$1:$A$70"),0)),0))')
        ws.cell(r, 11, f'=IF({g}="","",IFERROR(SUMIFS({E},{B},"Maintenance*",{B},"<>*Indirect*"),0))')
        ws.cell(r, 12, f'=IF({g}="","",IFERROR(SUMIFS({E},{B},"Rehabilitation*",{B},"<>*Indirect*"),0))')
        ws.cell(r, 13, f'=IF({g}="","",IFERROR(SUMIFS({E},{B},"*Indirect*"),0))')
        ws.cell(r, 14, f'=IF({g}="","",IFERROR(SUMIFS({E},{B},"Salvage*"),0))')
        ws.cell(r, 15, f'=IF({g}="","",IFERROR(INDEX(INDIRECT("\'"&{g}&"\'!$E$1:$E$70"),MATCH("Net Present Worth",INDIRECT("\'"&{g}&"\'!$B$1:$B$70"),0)),0))')
        ws.cell(r, 16, f'=IF({g}="","",O{r}-MIN($O${T0}:$O${T1}))')
        dcol = L(DC + 2 + 2 * NALT + i)  # this alternative's closure-days column in the by-year block (period-aware)
        ws.cell(r, 17, f'=IF({g}="","",IF(IFERROR(INDIRECT("\'"&{g}&"\'!$F$2"),0)>0,SUM(${dcol}$41:${dcol}$71),IFERROR(SUM(INDIRECT("\'"&{g}&"\'!$F$4:$F$10")),0)))')
        ws.cell(r, 18, f'=IF({g}="","",1-Q{r}/({GI}!$D$33*365))')
        for c in range(10, 17): ws.cell(r, c).number_format = CUR
        ws.cell(r, 18).number_format = '0.00%'
        for c in range(7, 19): ws.cell(r, c).font = F_B; ws.cell(r, c).border = BOX

    # ---- dashboard strip: six tiles, each a label row, a value row and a one-line explanation. Every tile
    # reads cells that already exist on this sheet, so nothing new is calculated here.
    F_KLAB = Font(name='Arial', size=8, bold=True, color='595959')
    F_KVAL = Font(name='Arial', size=16, bold=True, color='1D2733')
    F_KVAL_W = Font(name='Arial', size=16, bold=True, color='7F6000')
    F_KVAL_G = Font(name='Arial', size=16, bold=True, color='2E8B1F')
    F_KSUB = Font(name='Arial', size=9, color='595959')
    TILE_COLS = [(7, 10), (11, 14), (15, 18)]          # G:J, K:N, O:R
    top = Side(style='thin', color='BFC7D1')
    NONE_ANY = f'COUNT($O${T0}:$O${T1})=0'
    LOWNAME = f'INDEX($H${T0}:$H${T1},MATCH(MIN($O${T0}:$O${T1}),$O${T0}:$O${T1},0))'
    LOWROW = f'MATCH(MIN($O${T0}:$O${T1}),$O${T0}:$O${T1},0)'
    MARGIN = f'IF(COUNT($O${T0}:$O${T1})<2,0,SMALL($O${T0}:$O${T1},2)-MIN($O${T0}:$O${T1}))'
    CRF_ = f'(({GI}!$D$34/100)*(1+{GI}!$D$34/100)^{GI}!$D$33/((1+{GI}!$D$34/100)^{GI}!$D$33-1))'
    S2R = f'${L(DC+1)}$12:${L(DC+NALT)}$12'
    S7R = f'${L(DC+1)}${12+20}:${L(DC+NALT)}${12+20}'
    holds = (f'AND({LOWNAME}=INDEX($H${T0}:$H${T1},MATCH(SMALL({S2R},COUNTIF({S2R},0)+1),{S2R},0)),'
             f'{LOWNAME}=INDEX($H${T0}:$H${T1},MATCH(SMALL({S7R},COUNTIF({S7R},0)+1),{S7R},0)))')
    TILES = [
        ('LOWEST PRESENT WORTH',
         f'=IF({NONE_ANY},"-",TEXT(MIN($O${T0}:$O${T1}),"$#,##0"))',
         f'=IF({NONE_ANY},"No alternatives yet",{LOWNAME}&IF(INDEX($P${SEC_T+3}:$P${SEC_T+6},{LOWROW})="",""," - "&INDEX($P${SEC_T+3}:$P${SEC_T+6},{LOWROW})))',
         'plain'),
        ('MARGIN TO NEXT',
         f'=IF(COUNT($O${T0}:$O${T1})<2,"-",TEXT({MARGIN},"$#,##0"))',
         f'=IF(COUNT($O${T0}:$O${T1})<2,"only one alternative so far",TEXT({MARGIN}/MIN($O${T0}:$O${T1}),"0.0%")&IF({MARGIN}/MIN($O${T0}:$O${T1})<0.05," - under 5%, treat the two as tied"," of the lowest present worth"))',
         'flag'),
        ('EQUIVALENT ANNUAL COST',
         f'=IF({NONE_ANY},"-",TEXT(MIN($O${T0}:$O${T1})*{CRF_},"$#,##0"))',
         f'="per year, "&{GI}!$D$33&" years at "&{GI}!$D$34&"%"',
         'plain'),
        ('INITIAL CONSTRUCTION',
         f'=IF({NONE_ANY},"-",TEXT(INDEX($J${T0}:$J${T1},{LOWROW}),"$#,##0"))',
         f'="of the lowest-cost alternative, including "&{GI}!$D$36&"% mobilization and "&{GI}!$D$37&"% engineering"',
         'plain'),
        ('UNIT COST',
         f'=IF(OR({NONE_ANY},{GI}!$D$26=0,{GI}!$D$26=""),"-",TEXT(INDEX($J${T0}:$J${T1},{LOWROW})/{GI}!$D$26,"$#,##0")&" / S.Y.")',
         '="initial construction over the mainline area; see the benchmark under chart 5"',
         'plain'),
        ('RATE SENSITIVITY',
         f'=IF(COUNT($O${T0}:$O${T1})<2,"-",IF({holds},"holds 2-8%","the winner changes"))',
         f'=IF(COUNT($O${T0}:$O${T1})<2,"needs two alternatives",IF({holds},"the lowest-cost alternative is the same at every rate tested","the lowest-cost alternative is not the same at every rate: see chart 2"))',
         'good'),
    ]
    for k, (label, value, sub, tone) in enumerate(TILES):
        r0 = KPI0 + 4 * (k // 3)
        c0, c1 = TILE_COLS[k % 3]
        fill = PatternFill('solid', fgColor='F5F7FA')   # the margin tile is re-coloured by the rule below
        for rr, val, fnt in ((r0, label, F_KLAB), (r0 + 1, value, {'good': F_KVAL_G}.get(tone, F_KVAL)),
                             (r0 + 2, sub, F_KSUB)):
            ws.merge_cells(start_row=rr, start_column=c0, end_row=rr, end_column=c1)
            c = ws.cell(rr, c0, val); c.font = fnt
            c.alignment = Alignment(vertical='center', wrap_text=(rr == r0 + 2))
            for cc in range(c0, c1 + 1):
                cell = ws.cell(rr, cc)
                if fill: cell.fill = fill
                cell.border = Border(top=top if rr == r0 else None, bottom=top if rr == r0 + 2 else None,
                                     left=top if cc == c0 else None, right=top if cc == c1 else None)
        ws.row_dimensions[r0].height = 14
        ws.row_dimensions[r0 + 1].height = 24
        ws.row_dimensions[r0 + 2].height = 15
    ws.row_dimensions[KPI0 + 3].height = 6         # the gap between the two bands of tiles
    ws.row_dimensions[HDR - 1].height = 6          # and between the strip and the results table
    # the margin tile turns amber when the two best alternatives are within five percent of each other
    KPI_FLAG = (f'AND(COUNT($O${T0}:$O${T1})>1,'
                f'(SMALL($O${T0}:$O${T1},2)-MIN($O${T0}:$O${T1}))/MIN($O${T0}:$O${T1})<0.05)')


    S7 = f'${L(DC+1)}${12+20}:${L(DC+NALT)}${12+20}'  # 7.00% row of the sensitivity data block
    ws[f'G{VER}'] = (f'=IF(COUNT($O${T0}:$O${T1})=0,"No alternatives yet. Go to General Information and click Alternative Setup.",'
                f'"Lowest present worth: "&INDEX($H${T0}:$H${T1},MATCH(MIN($O${T0}:$O${T1}),$O${T0}:$O${T1},0))&IF(COUNT($O${T0}:$O${T1})<2,"   |   only one alternative so far",IF(SMALL($O${T0}:$O${T1},2)=MIN($O${T0}:$O${T1}),"   |   tied with the next alternative","   |   margin to next: "&TEXT(SMALL($O${T0}:$O${T1},2)-MIN($O${T0}:$O${T1}),"$#,##0")))'
                f'&"   |   "&{GI}!$D$34&"% over "&{GI}!$D$33&" years   |   lost revenue: "&{GI}!$D$38)')
    ws[f'G{VER}'].font = F_H; ws[f'G{VER}'].fill = FILL_VERDICT
    for c in range(8, 19): ws.cell(VER, c).fill = FILL_VERDICT
    S2 = f'${L(DC+1)}$12:${L(DC+NALT)}$12'  # 2.00% row (OMB A-94 real rate used by FAA PGL 22-01; 2.0% in 2026)
    low = f'INDEX($H${T0}:$H${T1},MATCH(MIN($O${T0}:$O${T1}),$O${T0}:$O${T1},0))'
    ws[f'G{VER2}'] = (f'=IF(COUNT($O${T0}:$O${T1})<2,"",'
                 f'IF({low}=INDEX($H${T0}:$H${T1},MATCH(SMALL({S2},COUNTIF({S2},0)+1),{S2},0)),"Same lowest-cost alternative at 2% (OMB A-94 real rate, FAA PGL 22-01)","At 2% (OMB A-94 real rate, FAA PGL 22-01) the lowest-cost alternative changes")'
                 f'&"   |   "&IF({low}=INDEX($H${T0}:$H${T1},MATCH(SMALL({S7},COUNTIF({S7},0)+1),{S7},0)),"same at 7% (pre-2022 AIP rule)","changes at 7% (pre-2022 AIP rule): see chart 2"))')
    # SMALL(range, COUNTIF(range,0)+1) = smallest non-zero value: unused alternative columns hold 0 and must not win
    ws[f'G{VER2}'].font = F_N

    # ---- chart data (column W onward)
    d = lambda k: L(DC + k)
    ws.cell(1, DC, 'CHART DATA (calculated automatically from the table on the left; do not edit)').font = F_H
    ws.column_dimensions[d(0)].width = 24
    for k in range(1, 15): ws.column_dimensions[d(k)].width = 13
    # category block rows 3-9
    ws.cell(3, DC, 'Present worth by category').font = F_H
    hdr(4, DC, ['Category'] + [None] * NALT)
    for i in range(NALT): ws.cell(4, DC + 1 + i, f'=IF($H${T0+i}="","",$H${T0+i})')
    cats = [('Initial construction', 10), ('Maintenance', 11), ('Rehabilitation', 12), ('Lost revenue', 13), ('Salvage', 14)]
    for k, (lab, col) in enumerate(cats):
        ws.cell(5 + k, DC, lab).font = F_DATA
        for i in range(NALT): c = ws.cell(5 + k, DC + 1 + i, f'=IF({L(col)}{T0+i}="",0,{L(col)}{T0+i})'); c.number_format = CUR; c.font = F_DATA
    # sensitivity block rows 10-36 (row 12 = 2.00%, row 32 = 7.00%)
    ws.cell(10, DC, 'Net present worth vs. discount rate').font = F_H
    hdr(11, DC, ['Rate (%)'] + [None] * NALT)
    for i in range(NALT): ws.cell(11, DC + 1 + i, f'=IF({d(1+i)}$4="","",{d(1+i)}$4)')
    for k in range(25):
        r = 12 + k; c = ws.cell(r, DC, 2 + 0.25 * k); c.number_format = '0.00"%"'; c.font = F_DATA
        for i in range(NALT):
            g = f'$G${T0+i}'; C, D = rng(g, 'C'), rng(g, 'D')
            c = ws.cell(r, DC + 1 + i, f'=IF({g}="",0,IFERROR($J${T0+i}+SUMPRODUCT(({C}<={GI}!$D$33)*{D}/(1+${d(0)}{r}/100)^{C}),0))'); c.number_format = CUR; c.font = F_DATA
    SENS0, SENS1 = 12, 36
    # by-year block rows 39-71
    Y0 = 39
    ws.cell(Y0, DC, 'By calendar year: spend (undiscounted), cumulative discounted cost, closure days').font = F_H
    hdr(Y0 + 1, DC, ['Year', 'Calendar year'] + [None] * (3 * NALT))
    for i in range(NALT):
        lab = lambda suffix: f'=IF({d(1+i)}$4="","",{d(1+i)}$4&"{suffix}")'
        ws.cell(Y0 + 1, DC + 2 + i, lab(' spend')); ws.cell(Y0 + 1, DC + 2 + NALT + i, lab(' cumulative PW')); ws.cell(Y0 + 1, DC + 2 + 2 * NALT + i, lab(' closure days'))
    for k in range(31):
        r = Y0 + 2 + k
        ws.cell(r, DC, k).font = F_DATA; c = ws.cell(r, DC + 1, f'={GI}!$D$25+{d(0)}{r}'); c.font = F_DATA
        for i in range(NALT):
            g = f'$G${T0+i}'; B, C, D, E = rng(g, 'B'), rng(g, 'C'), rng(g, 'D'), rng(g, 'E'); yr = f'{d(0)}{r}'
            spend = f'IF({yr}>{GI}!$D$33,0,IF({yr}=0,$J${T0+i},IFERROR(SUMIF({C},{yr},{D}),0)))'
            c = ws.cell(r, DC + 2 + i, f'=IF({g}="",0,{spend})'); c.number_format = CUR; c.font = F_DATA
            pv = f'IF({yr}>{GI}!$D$33,0,IF({yr}=0,$J${T0+i},IFERROR(SUMIF({C},{yr},{E}),0)))'
            prev = '' if k == 0 else f'{d(2+NALT+i)}{r-1}+'
            c = ws.cell(r, DC + 2 + NALT + i, f'=IF({g}="",0,{prev}{pv})'); c.number_format = CUR; c.font = F_DATA
            days = f'IFERROR(SUMIFS({D},{C},{yr},{B},"*Indirect*")/INDIRECT("\'"&{g}&"\'!$F$2"),0)'
            c = ws.cell(r, DC + 2 + 2 * NALT + i, f'=IF({g}="",0,IF({yr}>{GI}!$D$33,0,{days}))'); c.number_format = '0'; c.font = F_DATA
    Y1 = Y0 + 2 + 30

    # ---- charts, two per row under the results table
    def style(ch, title, h=8.0, w=16.5, xt='Alternative', yt='Present worth ($)'):
        ch.title = title; ch.width = w; ch.height = h; ch.legend.position = 'r'; ch.y_axis.numFmt = '$#,##0.0,,"M"'; ch.y_axis.majorGridlines = None
        ch.x_axis.delete = False; ch.y_axis.delete = False
        ch.x_axis.title = xt; ch.y_axis.title = yt
    def color(ch, line=False):
        for s, rgb in zip(ch.series, ALT_COLORS):
            if line: s.graphicalProperties.line.solidFill = rgb; s.graphicalProperties.line.width = 22000; s.marker.symbol = 'none'; s.smooth = False
            else: s.graphicalProperties.solidFill = rgb; s.graphicalProperties.line.solidFill = rgb
    # ---- RealCost-style comparison block (agency cost / user cost / total, present worth and EUAC)
    ws[f'G{CMP}'] = 'COMPARISON  (RealCost layout: agency cost, user cost, total; present worth and equivalent uniform annual cost)'; ws[f'G{CMP}'].font = F_T
    hdr(CMPH, 7, ['Alternative', 'Agency cost PW', 'Agency EUAC', 'User cost PW (lost revenue)', 'User EUAC', 'Total PW', 'Total EUAC', 'vs. lowest ($)', 'vs. lowest (%)', 'Lowest?'])
    CRF = f'(({GI}!$D$34/100)*(1+{GI}!$D$34/100)^{GI}!$D$33/((1+{GI}!$D$34/100)^{GI}!$D$33-1))'
    for i in range(NALT):
        r = CMP0 + i; src = T0 + i; g = f'$G${src}'
        ws.cell(r, 7, f'=IF({g}="","",$H${src})')
        ws.cell(r, 8, f'=IF({g}="","",$J${src}+$K${src}+$L${src}+$N${src})')
        ws.cell(r, 9, f'=IF({g}="","",H{r}*{CRF})')
        ws.cell(r, 10, f'=IF({g}="","",$M${src})')
        ws.cell(r, 11, f'=IF({g}="","",J{r}*{CRF})')
        ws.cell(r, 12, f'=IF({g}="","",$O${src})')
        ws.cell(r, 13, f'=IF({g}="","",L{r}*{CRF})')
        ws.cell(r, 14, f'=IF({g}="","",$P${src})')
        ws.cell(r, 15, f'=IF({g}="","",IFERROR($P${src}/MIN($O${T0}:$O${T1}),0))')
        ws.cell(r, 16, f'=IF({g}="","",IF($O${src}=MIN($O${T0}:$O${T1}),"lowest",""))')
        for c in range(8, 15): ws.cell(r, c).number_format = CUR
        ws.cell(r, 15).number_format = '0.0%'
        for c in range(7, 17): ws.cell(r, c).font = F_B; ws.cell(r, c).border = BOX
    ws.row_dimensions[CMPH].height = 27
    ws[f'G{CMPN}'] = ('Agency cost = initial construction + maintenance + rehabilitation + salvage. User cost = lost airport revenue during runway closures, the airport-side counterpart of the user delay cost in the Caltrans/FHWA RealCost layout. '
                 'EUAC = PW x r(1+r)^P / ((1+r)^P - 1) at the discount rate and analysis period on General Information. "vs. lowest (%)" is the difference as a share of the lowest total PW.')
    ws[f'G{CMPN}'].font = F_N
    ws[f'G{CHT}'] = 'CHARTS'; ws[f'G{CHT}'].font = F_T
    ch = BarChart(); ch.type = 'col'; ch.grouping = 'stacked'; ch.overlap = 100; ch.gapWidth = 60
    ch.add_data(Reference(ws, min_col=DC, max_col=DC + NALT, min_row=5, max_row=9), titles_from_data=True, from_rows=True)
    ch.set_categories(Reference(ws, min_col=DC + 1, max_col=DC + NALT, min_row=4, max_row=4))
    # one ramp, dark to light, so the stack reads in the order the categories are listed
    for s, rgb in zip(ch.series, ['3E4C59', '6B7A88', '98A4AF', 'C4CDD5', 'E4E9ED']): s.graphicalProperties.solidFill = rgb; s.graphicalProperties.line.solidFill = rgb
    ch.x_axis.tickLblPos = 'low'; style(ch, '1. Present worth by category (salvage below zero)', xt='Alternative', yt='Present worth ($)'); ws.add_chart(ch, f'G{C1}')
    ch = LineChart()
    ch.add_data(Reference(ws, min_col=DC + 1, max_col=DC + NALT, min_row=11, max_row=SENS1), titles_from_data=True); ch.set_categories(Reference(ws, min_col=DC, min_row=SENS0, max_row=SENS1))
    color(ch, line=True); ch.x_axis.tickLblSkip = 4; ch.x_axis.numFmt = '0.00"%"'; style(ch, '2. Net present worth vs. discount rate (TDOT 3%, FAA 2%, pre-2022 rule 7%)', xt='Discount rate (%)', yt='Net present worth ($)'); ws.add_chart(ch, f'N{C1}')
    ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'; ch.gapWidth = 40
    ch.add_data(Reference(ws, min_col=DC + 2, max_col=DC + 1 + NALT, min_row=Y0 + 1, max_row=Y1), titles_from_data=True); ch.set_categories(Reference(ws, min_col=DC + 1, min_row=Y0 + 2, max_row=Y1))
    color(ch); ch.x_axis.tickLblSkip = 5; ch.x_axis.tickLblPos = 'low'; style(ch, '3. Expenditure stream by calendar year (undiscounted)', xt='Calendar year', yt='Spend, undiscounted ($)'); ws.add_chart(ch, f'G{C2}')
    ch = LineChart()
    ch.add_data(Reference(ws, min_col=DC + 2 + NALT, max_col=DC + 1 + 2 * NALT, min_row=Y0 + 1, max_row=Y1), titles_from_data=True); ch.set_categories(Reference(ws, min_col=DC + 1, min_row=Y0 + 2, max_row=Y1))
    color(ch, line=True); ch.x_axis.tickLblSkip = 5; style(ch, '4. Cumulative discounted cost', xt='Calendar year', yt='Cumulative discounted cost ($)'); ws.add_chart(ch, f'N{C2}')
    ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'; ch.gapWidth = 40
    ch.add_data(Reference(ws, min_col=DC + 2 + 2 * NALT, max_col=DC + 1 + 3 * NALT, min_row=Y0 + 1, max_row=Y1), titles_from_data=True); ch.set_categories(Reference(ws, min_col=DC + 1, min_row=Y0 + 2, max_row=Y1))
    color(ch); ch.x_axis.tickLblSkip = 5; style(ch, '5. Runway closure days by calendar year', xt='Calendar year', yt='Closure days'); ch.y_axis.numFmt = '0'; ws.add_chart(ch, f'G{C3}')
    # ---- pavement section read back from the pay-item quantities (display only; no cost depends on it)
    PCF = f'$J${SEC_T + 1}'
    AREA = f'{GI}!$D$26'; SHLD = f'{GI}!$D$27'
    ws[f'G{SEC_T}'] = 'PAVEMENT SECTION  (read back from the quantities already entered; nothing here changes a cost)'; ws[f'G{SEC_T}'].font = F_T
    ws[f'G{SEC_T+1}'] = 'Asphalt unit weight, pcf:'; ws[f'G{SEC_T+1}'].font = F_H
    c = ws[f'J{SEC_T+1}']; c.value = 145; c.number_format = '0'; c.font = F_B; c.fill = PatternFill('solid', fgColor='D9D9D9'); c.border = BOX
    ws[f'K{SEC_T+1}'] = ('Only asphalt needs an assumption: 145 pcf reproduces the Murfreesboro section to the hundredth of an inch. '
                 'Volume items convert directly, and concrete carries its thickness in the pay-item name.')
    ws[f'K{SEC_T+1}'].font = F_N
    hdr(SEC_T + 2, 7, ['Alternative', 'Surface course (in)', 'Aggregate base (in)', 'Subbase (in)', 'Treated subgrade',
                'Total section (in)', 'Excavation (in)', 'Excavation check', None])
    hdr(SEC_T + 2, 16, ['Section from the quantities', 'Same quantities over mainline + shoulder'])
    ws.row_dimensions[SEC_T + 2].height = 27
    for i in range(NALT):
        r = SEC_T + 3 + i; g = f'$G${T0+i}'
        B, C, D, E = rng(g, 'B', 13, 22), rng(g, 'C', 13, 22), rng(g, 'D', 13, 22), rng(g, 'E', 13, 22)
        tons = f'SUMIFS({E},{D},"TON")'
        cy = lambda pat: f'SUMIFS({E},{D},"C.Y.",{B},"{pat}")'
        agg = f'({cy("P-2*")}+{cy("P-3*")})'
        desc = f'IFERROR(INDEX({C},MATCH("Concrete Pavement*",{C},0)),"")'
        pcc = f'IFERROR(VALUE(TRIM(SUBSTITUTE(MID({desc},FIND(",",{desc})+1,99),"-inch",""))),0)'
        item = lambda pat: f'IFERROR(SUBSTITUTE(LEFT(INDEX({B},MATCH("{pat}",{B},0)),5),"-",""),"")'
        ws.cell(r, 7, f'=IF({g}="","",$H${T0+i})')
        ws.cell(r, 8, f'=IF({g}="","",IF({tons}>0,{tons}*2666.6667/({PCF}*{AREA}),{pcc}))')
        ws.cell(r, 9, f'=IF({g}="","",36*{agg}/{AREA})')
        ws.cell(r, 10, f'=IF({g}="","",36*{cy("P-154*")}/{AREA})')
        ws.cell(r, 11, f'=IF({g}="","",IF(SUMIFS({E},{D},"S.Y.",{B},"P-155*")+SUMIFS({E},{D},"S.Y.",{B},"P-156*")+SUMIFS({E},{D},"S.Y.",{B},"P-157*")+SUMIFS({E},{D},"S.Y.",{B},"P-158*")>0,"yes","-"))')
        ws.cell(r, 12, f'=IF({g}="","",H{r}+I{r}+J{r})')
        ws.cell(r, 13, f'=IF({g}="","",36*{cy("P-152*")}/{AREA})')
        ws.cell(r, 14, f'=IF({g}="","",IF(M{r}=0,"no excavation item",IF(ABS(M{r}-L{r})<=0.5,"agrees","differs by "&TEXT(M{r}-L{r},"0.0")&" in")))')
        ws.cell(r, 16, f'=IF({g}="","",TEXT(H{r},"0")&"\"\" "&IF({tons}>0,{item("P-4*")},{item("P-501*")})'
                       f'&IF(I{r}>0," on "&TEXT(I{r},"0")&"\"\" "&{item("P-2*")},"")'
                       f'&IF(J{r}>0," on "&TEXT(J{r},"0")&"\"\" P154",""))')
        fs = f'{AREA}/({AREA}+{SHLD})'   # shoulder reading: bound layers thin out over the larger area, concrete keeps its named thickness
        ws.cell(r, 17, f'=IF(OR({g}="",{SHLD}=0),"",TEXT(IF({tons}>0,H{r}*{fs},H{r}),"0")&"\"\" "&IF({tons}>0,{item("P-4*")},{item("P-501*")})'
                       f'&IF(I{r}>0," on "&TEXT(I{r}*{fs},"0")&"\"\" "&{item("P-2*")},"")'
                       f'&IF(J{r}>0," on "&TEXT(J{r}*{fs},"0")&"\"\" P154",""))')
        for c2 in range(8, 14): ws.cell(r, c2).number_format = '0.00'
        for c2 in list(range(7, 15)) + [16, 17]: ws.cell(r, c2).font = F_B; ws.cell(r, c2).border = BOX
    ws[f'G{SEC_T+7}'] = ('Thickness is not an input. Volume items give inches = 36 x C.Y. / mainline S.Y.; asphalt gives inches = 2,666.67 x tons / (pcf x mainline S.Y.); '
                 'items measured by area carry no implied thickness. The last column is the section as the quantities describe it: paste it into the alternative description so the two can never disagree. The read-back divides by the mainline area only; when a shoulder area is entered, the last column spreads the same quantities over mainline plus shoulder, which is the reading to use if the quantities were taken off both.')
    ws[f'G{SEC_T+7}'].font = F_N
    # chart data for the two section charts (columns W onward, below the by-year block)
    SEC = 74
    ws.cell(SEC, DC, 'PAVEMENT SECTION (inches): mainline, then the reading if the quantities also cover the shoulder').font = F_H
    hdr(SEC + 1, DC, ['Layer'] + [None] * NALT)
    for i in range(NALT): ws.cell(SEC + 1, DC + 1 + i, f'=IF({d(1+i)}$4="","",{d(1+i)}$4)')
    layers = [('Subbase (P-154)', 'J'), ('Aggregate base', 'I'), ('Asphalt', 'H'), ('Concrete', 'H')]
    for k, (lab, col) in enumerate(layers):
        ws.cell(SEC + 2 + k, DC, lab).font = F_DATA
        for i in range(NALT):
            r83 = SEC_T + 3 + i; g = f'$G${T0+i}'
            tons = f'SUMIFS({rng(g, "E", 13, 22)},{rng(g, "D", 13, 22)},"TON")'
            v = f'{col}{r83}'
            if lab == 'Asphalt': v = f'IF({tons}>0,H{r83},0)'
            if lab == 'Concrete': v = f'IF({tons}>0,0,H{r83})'
            c = ws.cell(SEC + 2 + k, DC + 1 + i, f'=IF({g}="",0,IFERROR({v},0))'); c.number_format = '0.00'; c.font = F_DATA
    ws.cell(SEC + 7, DC, 'Shoulder reading (blank unless shoulder area is entered)').font = F_H
    hdr(SEC + 8, DC, ['Layer'] + [None] * NALT)
    for i in range(NALT): ws.cell(SEC + 8, DC + 1 + i, f'={d(1+i)}$4')
    for k, lab in enumerate([l for l, _ in layers]):
        ws.cell(SEC + 9 + k, DC, lab).font = F_DATA
        for i in range(NALT):
            src = f'{d(1+i)}{SEC + 2 + k}'
            factor = '1' if lab == 'Concrete' else f'{AREA}/({AREA}+{SHLD})'
            scale = f'IF({SHLD}>0,{factor},0)'
            c = ws.cell(SEC + 9 + k, DC + 1 + i, f'=IFERROR({src}*{scale},0)'); c.number_format = '0.00'; c.font = F_DATA
    LAYER_RGB = ['DCCBA0', 'C8A96E', '3A3A3A', 'C6CBD0']
    for anch, r0, title in [(f'G{CH67}', SEC + 2, '6. Pavement section, mainline (inches, surface on top)'),
                            (f'N{CH67}', SEC + 9, '7. Pavement section with the shoulder (empty unless a shoulder area is entered)')]:
        ch = BarChart(); ch.type = 'col'; ch.grouping = 'stacked'; ch.overlap = 100; ch.gapWidth = 80
        ch.add_data(Reference(ws, min_col=DC, max_col=DC + NALT, min_row=r0, max_row=r0 + 3), titles_from_data=True, from_rows=True)
        ch.set_categories(Reference(ws, min_col=DC + 1, max_col=DC + NALT, min_row=r0 - 1, max_row=r0 - 1))
        for s2, rgb in zip(ch.series, LAYER_RGB): s2.graphicalProperties.solidFill = rgb; s2.graphicalProperties.line.solidFill = rgb
        style(ch, title, xt='Alternative', yt='Thickness (inches)'); ch.y_axis.numFmt = '0'; ch.x_axis.tickLblPos = 'low'
        ws.add_chart(ch, anch)
    # ---- unit cost of every alternative against published Tennessee runway work (chart 8, beside chart 5).
    # A stacked bar with an invisible first series is how a floating band is drawn in Excel: the published
    # range starts at $210 and spans $70, so it floats instead of growing from zero.
    BM = SEC + 15
    ws.cell(BM, DC, 'UNIT COST ($/S.Y.) against published Tennessee runway work (Typical Values, unit costs)').font = F_H
    hdr(BM + 1, DC, ['Series'] + [None] * (NALT + 1))
    for i in range(NALT): ws.cell(BM + 1, DC + 1 + i, f'=IF({d(1+i)}$4="","",{d(1+i)}$4)')
    ws.cell(BM + 1, DC + 1 + NALT, 'TN 2024-25, all-in')
    ws.cell(BM + 2, DC, 'Below the published range').font = F_DATA
    ws.cell(BM + 3, DC, 'Unit cost / published range').font = F_DATA
    for i in range(NALT):
        c = ws.cell(BM + 2, DC + 1 + i, 0); c.number_format = '$#,##0'; c.font = F_DATA
        c = ws.cell(BM + 3, DC + 1 + i, f'=IFERROR(IF($G${T0+i}="",0,$J${T0+i}/{GI}!$D$26),0)'); c.number_format = '$#,##0'; c.font = F_DATA
    c = ws.cell(BM + 2, DC + 1 + NALT, 210); c.number_format = '$#,##0'; c.font = F_DATA
    c = ws.cell(BM + 3, DC + 1 + NALT, 70); c.number_format = '$#,##0'; c.font = F_DATA
    ch = BarChart(); ch.type = 'bar'; ch.grouping = 'stacked'; ch.overlap = 100; ch.gapWidth = 60
    ch.add_data(Reference(ws, min_col=DC, max_col=DC + 1 + NALT, min_row=BM + 2, max_row=BM + 3),
                titles_from_data=True, from_rows=True)
    ch.set_categories(Reference(ws, min_col=DC + 1, max_col=DC + 1 + NALT, min_row=BM + 1, max_row=BM + 1))
    base, span = ch.series
    base.graphicalProperties = GraphicalProperties(noFill=True); base.graphicalProperties.line.noFill = True
    span.graphicalProperties.solidFill = 'BBD4EE'; span.graphicalProperties.line.solidFill = 'BBD4EE'
    span.data_points = [DataPoint(idx=NALT, spPr=GraphicalProperties(solidFill='CBD5DF'))]
    span.dLbls = DataLabelList(showVal=True, showSerName=False, showCatName=False,
                               showLegendKey=False, showPercent=False, showBubbleSize=False)
    span.dLbls.dLblPos = 'inEnd'
    ch.legend = None
    ch.title = '8. Unit cost against published Tennessee runway work'
    ch.width = 16.5; ch.height = 8.0
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.x_axis.title = 'Alternative'; ch.y_axis.title = 'Initial construction / mainline S.Y.'
    ch.y_axis.numFmt = '$#,##0'; ch.y_axis.majorGridlines = None
    ws.add_chart(ch, f'N{C3}')
    ws[f'G{BMN}'] = ('Chart 8 divides each alternative\'s initial construction by the mainline area. The band is the all-in '
                       'range for recent Tennessee runway work, $210 to $280 per S.Y. on Typical Values, which covers pavement, '
                       'lighting and grading; this workbook prices the pavement contract, so every bar should sit below the band. '
                       'Far below it, or above it, is worth a second look at the quantities.')
    ws[f'G{BMN}'].font = F_N
    ws[f'G{NOTE67}'] = ('Charts 6 and 7 stack the layers as they sit in the ground, surface on top. Chart 7 stays empty unless a shoulder area is entered on General Information, '
                  'and then shows the same quantities spread over mainline plus shoulder, which is the lower bound on each thickness.')
    ws[f'G{NOTE67}'].font = F_N
    ws[f'G{VBACH-1}'] = ('Below: "Chart 1", the chart the Alternative Setup form draws and keeps up to date. It shows initial '
                         'construction and net present worth per alternative, which chart 1 above shows with the categories '
                         'broken out. It is left in place so the form keeps working.')
    ws[f'G{VBACH-1}'].font = F_N
    ws[f'G{HOW}'] = "How to read: 1 shows where each alternative's cost sits; 2 whether the lowest-cost alternative holds at other discount rates (FAA now uses the OMB A-94 real rate, about 2%; 7% was the rule before 2022); 3 and 4 when the money is spent and when the higher first cost is paid back; 5 how often and how long the runway closes."; ws[f'G{HOW}'].font = F_N
    # ---- project identity and the locator map. Coordinates and the state outline are embedded (geo_data.py),
    # so the map draws with no internet connection, no add-in and no external reference.
    GI_ = GI
    ws['L1'] = (f'=IF({GI_}!$D$9="","No airport selected yet: pick one on General Information.",'
                f'{GI_}!$D$9&" ("&{GI_}!$D$10&")  |  "&{GI_}!$D$11&", "&{GI_}!$D$13&" region"'
                f'&IF({GI_}!$D$22="",""," |  "&{GI_}!$D$22)&IF({GI_}!$D$23="",""," ("&{GI_}!$D$23&")")'
                f'&"  |  construction "&TEXT({GI_}!$D$25,"0")&"  |  "&TEXT({GI_}!$D$33,"0")&" years at "&TEXT({GI_}!$D$34,"0.##")&"%")')
    ws['L1'].font = F_H

    MAP0 = 130                                   # map data, below every other block and outside the print area
    A0 = MAP0 + 2                                # first airport row
    ws.cell(MAP0, DC, 'MAP DATA (embedded coordinates and state outline; reference only, do not edit)').font = F_H
    hdr(MAP0 + 1, DC, ['Airport', 'FAA ID', 'Longitude', 'Latitude', 'Region', 'Elevation (ft)', None,
                       'This project: longitude', 'Latitude', 'Label', None, 'Outline longitude', 'Outline latitude'])
    REGIONS = ['West', 'Middle', 'East']          # the TDOT Grand Divisions, the grouping Pay_Items is built around
    ordered, bands = [], {}
    for reg in REGIONS:
        start = A0 + len(ordered)
        ordered += [a for a in geo_data.AIRPORTS if a[2] == reg]
        bands[reg] = (start, A0 + len(ordered) - 1)
    ordered += [a for a in geo_data.AIRPORTS if a[2] not in REGIONS]
    for i, (name, ident, region, lon, lat, elev) in enumerate(ordered):
        r = A0 + i
        for k, v in enumerate([name, ident, lon, lat, region, elev]):
            c = ws.cell(r, DC + k, v); c.font = F_DATA
            if k in (2, 3): c.number_format = '0.00000'
    AN = len(ordered)
    ID_ = f'${L(DC+1)}${A0}:${L(DC+1)}${A0+AN-1}'
    LON_ = f'${L(DC+2)}${A0}:${L(DC+2)}${A0+AN-1}'
    LAT_ = f'${L(DC+3)}${A0}:${L(DC+3)}${A0+AN-1}'
    ELV_ = f'${L(DC+5)}${A0}:${L(DC+5)}${A0+AN-1}'
    pick = lambda rng_: f'=IFERROR(IF(INDEX({rng_},MATCH({GI_}!$D$10,{ID_},0))="","",INDEX({rng_},MATCH({GI_}!$D$10,{ID_},0))),"")'
    c = ws.cell(A0, DC + 7, pick(LON_)); c.number_format = '0.00000'; c.font = F_DATA
    c = ws.cell(A0, DC + 8, pick(LAT_)); c.number_format = '0.00000'; c.font = F_DATA
    ws.cell(A0, DC + 9, f'=IF({GI_}!$D$9="","",{GI_}!$D$10)').font = F_DATA   # short label: the block below spells the name out
    br = A0
    for seg in geo_data.BORDER:                  # a blank row between segments leaves a gap instead of a join
        for lon, lat in seg:
            c = ws.cell(br, DC + 11, lon); c.number_format = '0.00000'; c.font = F_DATA
            c = ws.cell(br, DC + 12, lat); c.number_format = '0.00000'; c.font = F_DATA
            br += 1
        br += 1
    BN = br - 1

    mc = ScatterChart(); mc.height = 5.2; mc.width = 11.0; mc.legend = None
    def scatter(col_x, col_y, r1, r2):
        return Series(Reference(ws, min_col=col_y, min_row=r1, max_row=r2), Reference(ws, min_col=col_x, min_row=r1, max_row=r2))
    outline = scatter(DC + 11, DC + 12, A0, BN)
    outline.marker.symbol = 'none'; outline.graphicalProperties.line.solidFill = '9AA0A6'; outline.graphicalProperties.line.width = 9525
    outline.tx = SeriesLabel(v='Tennessee')
    DIV_RGB = {'West': '7FA8CF', 'Middle': '2A78D6', 'East': '1D3F6E'}
    divisions = []
    for reg in REGIONS:
        r1, r2 = bands[reg]
        d_ = scatter(DC + 2, DC + 3, r1, r2)
        d_.marker.symbol = 'circle'; d_.marker.size = 4
        d_.marker.graphicalProperties.solidFill = DIV_RGB[reg]; d_.marker.graphicalProperties.line.solidFill = DIV_RGB[reg]
        d_.graphicalProperties.line.noFill = True; d_.tx = SeriesLabel(v=reg)
        divisions.append(d_)
    here = scatter(DC + 7, DC + 8, A0, A0)
    here.marker.symbol = 'circle'; here.marker.size = 9
    here.marker.graphicalProperties.solidFill = 'C1440E'; here.marker.graphicalProperties.line.solidFill = '7F2D09'
    here.graphicalProperties.line.noFill = True
    here.tx = SeriesLabel(strRef=StrRef(f"Summary!${L(DC+9)}${A0}"))
    here.dLbls = DataLabelList(showSerName=True, showVal=False, showCatName=False, showLegendKey=False, dLblPos='r')
    for sr in [outline] + divisions + [here]: mc.series.append(sr)
    mc.x_axis.scaling.min, mc.x_axis.scaling.max = -90.6, -81.4
    mc.y_axis.scaling.min, mc.y_axis.scaling.max = 34.8, 36.9
    for ax, t in ((mc.x_axis, 'Longitude'), (mc.y_axis, 'Latitude')):
        ax.delete = False; ax.title = t; ax.majorTickMark = 'none'; ax.minorTickMark = 'none'
        ax.majorGridlines = None; ax.numFmt = ';;;'      # the axes are named but degree values would only be noise
    for col in range(19, 23):                       # S:V band, the card header
        c = ws.cell(2, col); c.fill = PatternFill('solid', fgColor='1D2733')
        c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    ws['S2'] = 'PROJECT LOCATION'
    ws.add_chart(mc, 'S3')
    for k, (reg, rgb) in enumerate(list(DIV_RGB.items()) + [('This project', 'C1440E')]):
        c = ws.cell(17, 19 + k, '\u25a0 ' + reg)     # legend in cells: the plot keeps its full height
        c.font = Font(name='Arial', size=9, bold=True, color=rgb)

    ws['S19'] = 'PROJECT'; ws['S19'].font = F_H
    county = f'IFERROR(VLOOKUP({GI_}!$D$9,{GI_}!$L$10:$Q$88,4,FALSE),"")'
    rowsS = [('Airport', f'=IF({GI_}!$D$9="","",{GI_}!$D$9&" ("&{GI_}!$D$10&")")'),
             ('City / county', f'=IF({GI_}!$D$9="","",{GI_}!$D$11&IF({county}="",""," / "&{county}&" County"))'),
             ('TDOT Grand Division', f'=IF({GI_}!$D$9="","",{GI_}!$D$13)'),
             ('Coordinates', f'=IF({L(DC+8)}{A0}="","not in the reference list",TEXT({L(DC+8)}{A0},"0.0000")&"\u00b0 N, "&TEXT(-{L(DC+7)}{A0},"0.0000")&"\u00b0 W")'),
             ('Elevation', f'=IFERROR(IF(INDEX({ELV_},MATCH({GI_}!$D$10,{ID_},0))="","",TEXT(INDEX({ELV_},MATCH({GI_}!$D$10,{ID_},0)),"#,##0")&" ft"),"")'),
             ('Branch / project', f'=IF({GI_}!$D$22="","",{GI_}!$D$22&IF({GI_}!$D$23=""," ",", "&{GI_}!$D$23))'),
             ('Mainline area', f'=IF({GI_}!$D$26="","",TEXT({GI_}!$D$26,"#,##0")&" S.Y."&IF({GI_}!$D$27>0," + "&TEXT({GI_}!$D$27,"#,##0")&" S.Y. shoulder",""))')]
    for k, (lab, f) in enumerate(rowsS):
        ws.cell(20 + k, 19, lab).font = F_B
        ws.cell(20 + k, 20, f).font = F_B
    ws['S27'] = 'Runway width, ft'; ws['S27'].font = F_B
    c = ws['T27']; c.value = 100; c.number_format = '0'; c.font = F_B
    c.fill = PatternFill('solid', fgColor='D9D9D9'); c.border = BOX
    ws['S28'] = (f'=IF({GI_}!$D$9="","","Pricing basis: every unit cost comes from the single Pay_Items \'Unit Cost\' column. '
                 f'The Middle, West and East average-cost columns beside it are empty, so this '
                 f'"&{GI_}!$D$13&" division project is priced on the same statewide numbers as any other.")')
    ws['S28'].font = Font(name='Arial', size=9, color='7F6000')
    ws['S28'].fill = PatternFill('solid', fgColor='FFF3CD')
    ws['S29'] = 'Coordinates are embedded (Method sheet); nothing here goes online. Five private fields have none and plot no dot.'
    ws['S30'] = 'Google Earth: import LCCA_KML_Export.bas once (Alt+F11, File, Import File), then Alt+F8 and run ExportLCCAKML.'
    ws['S29'].font = F_N
    ws['S30'].font = F_N
    # the chart-data block starts at column W, so these notes are merged and wrapped instead of spilling into it
    for row, ht in ((28, 44), (29, 26), (30, 26)):
        ws.merge_cells(start_row=row, start_column=19, end_row=row, end_column=22)
        ws.cell(row, 19).alignment = Alignment(wrap_text=True, vertical='top')
        ws.row_dimensions[row].height = ht

    dxf_low = DifferentialStyle(fill=PatternFill(bgColor='FFDDEBF7'), font=Font(bold=True, color='FF1F3864'))
    # bars inside the net-present-worth column, so the table reads without going to a chart
    ws.conditional_formatting.add(f'O{T0}:O{T1}',
                                  DataBarRule(start_type='num', start_value=0, end_type='max',
                                              color='FFBBD4EE', showValue=True, minLength=None, maxLength=None))
    dxf_flag = DifferentialStyle(fill=PatternFill(bgColor='FFFFF3CD'), font=Font(bold=True, color='FF7F6000'))
    ws.conditional_formatting.add(f'K{KPI0}:N{KPI0+2}', Rule(type='expression', formula=[KPI_FLAG], stopIfTrue=False, dxf=dxf_flag))
    for sqref, f in [(f'G{T0}:R{T1}', f'AND($G{T0}<>"",COUNT($O${T0}:$O${T1})>0,$O{T0}=MIN($O${T0}:$O${T1}))'),
                     (f'G{CMP0}:P{CMP1}', f'AND($G{CMP0}<>"",COUNT($L${CMP0}:$L${CMP1})>0,$L{CMP0}=MIN($L${CMP0}:$L${CMP1}))')]:
        rule = Rule(type='expression', formula=[f], stopIfTrue=False, dxf=dxf_low)
        ws.conditional_formatting.add(sqref, rule)
    ws.freeze_panes = f'A{T0}'
    wb.save(path)

# ---------------------------------------------------------------- transplant
def _children_xml(styles_xml, tag):
    doc = minidom.parseString(styles_xml)
    nodes = doc.getElementsByTagName(tag)
    if not nodes: return []
    return [n.toxml() for n in nodes[0].childNodes if n.nodeType == n.ELEMENT_NODE]

def transplant(work, sheet_part, drawing_part, chart_start, sheet_index):
    """work: unpacked .xlsm dir. sheet_part e.g. 'xl/worksheets/sheet14.xml'. drawing_part e.g. 'xl/drawings/drawing10.xml'.
    chart_start: first free chart index (e.g. 7 when chart1..6 exist)."""
    tmp = tempfile.mkdtemp(); scratch = os.path.join(tmp, 's.xlsx'); build_scratch(scratch)
    sd = os.path.join(tmp, 'u'); zipfile.ZipFile(scratch).extractall(sd)
    rd = lambda p: open(os.path.join(work, p), encoding='utf-8').read()
    wr = lambda p, s: open(os.path.join(work, p), 'w', encoding='utf-8').write(s)
    srd = lambda p: open(os.path.join(sd, p), encoding='utf-8').read()

    # --- styles: append scratch numFmts/fonts/fills/borders/cellXfs to the package styles (parsed, not regex-sliced)
    bs = rd('xl/styles.xml'); ss = srd('xl/styles.xml')
    def append(xml, tag, new_items):
        m = re.search(r'<%s count="(\d+)"' % tag, xml); cnt = int(m.group(1))
        xml = re.sub(r'<%s count="\d+"' % tag, '<%s count="%d"' % (tag, cnt + len(new_items)), xml, count=1)
        end = xml.index('</%s>' % tag); return xml[:end] + ''.join(new_items) + xml[end:], cnt
    nfs = _children_xml(ss, 'numFmts')
    base_nf_ids = [int(x) for x in re.findall(r'<numFmt numFmtId="(\d+)"', bs)]
    next_nf = max(base_nf_ids + [163]) + 1
    nf_map = {}; new_nfs = []
    for nf in nfs:
        old = int(re.search(r'numFmtId="(\d+)"', nf).group(1)); nf_map[old] = next_nf
        new_nfs.append(re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % next_nf, nf)); next_nf += 1
    if new_nfs:
        if '<numFmts' in bs: bs, _ = append(bs, 'numFmts', new_nfs)
        else: bs = bs.replace('<fonts', '<numFmts count="%d">%s</numFmts><fonts' % (len(new_nfs), ''.join(new_nfs)), 1)
    bs, font_off = append(bs, 'fonts', _children_xml(ss, 'fonts'))
    bs, fill_off = append(bs, 'fills', _children_xml(ss, 'fills'))
    bs, border_off = append(bs, 'borders', _children_xml(ss, 'borders'))
    new_xfs = []
    for xf in _children_xml(ss, 'cellXfs'):
        xf = re.sub(r'fontId="(\d+)"', lambda m: 'fontId="%d"' % (int(m.group(1)) + font_off), xf)
        xf = re.sub(r'fillId="(\d+)"', lambda m: 'fillId="%d"' % (int(m.group(1)) + fill_off), xf)
        xf = re.sub(r'borderId="(\d+)"', lambda m: 'borderId="%d"' % (int(m.group(1)) + border_off), xf)
        xf = re.sub(r'numFmtId="(\d+)"', lambda m: 'numFmtId="%d"' % nf_map.get(int(m.group(1)), int(m.group(1))), xf)
        xf = re.sub(r'xfId="\d+"', 'xfId="0"', xf); new_xfs.append(xf)
    bs, xf_off = append(bs, 'cellXfs', new_xfs)
    dxfs = _children_xml(ss, 'dxfs')
    dxf_off = 0
    if dxfs:
        if '<dxfs' in bs: bs, dxf_off = append(bs, 'dxfs', dxfs)
        else: bs = bs.replace('</styleSheet>', '<dxfs count="%d">%s</dxfs></styleSheet>' % (len(dxfs), ''.join(dxfs)))
    minidom.parseString(bs)  # must stay well-formed
    wr('xl/styles.xml', bs)

    # --- sheet XML: remap styles, inline any shared strings, keep sheetPr/codeName and the drawing rel
    sx = srd('xl/worksheets/sheet1.xml'); orig = rd(sheet_part)
    ssp = os.path.join(sd, 'xl/sharedStrings.xml')
    strings = [html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S))) for si in re.findall(r'<si>(.*?)</si>', open(ssp, encoding='utf-8').read(), re.S)] if os.path.exists(ssp) else []
    sx = re.sub(r'<c r="([A-Z]+\d+)"([^>]*?) t="s"([^>]*)><v>(\d+)</v></c>',
                lambda m: '<c r="%s"%s t="inlineStr"%s><is><t xml:space="preserve">%s</t></is></c>' % (m.group(1), m.group(2), m.group(3), html.escape(strings[int(m.group(4))], quote=False)), sx)
    sx = re.sub(r' s="(\d+)"', lambda m: ' s="%d"' % (int(m.group(1)) + xf_off), sx)
    sx = re.sub(r'dxfId="(\d+)"', lambda m: 'dxfId="%d"' % (int(m.group(1)) + dxf_off), sx)
    sheetpr = re.search(r'<sheetPr[^>]*>.*?</sheetPr>|<sheetPr[^>]*/>', orig, re.S)
    sx = re.sub(r'<sheetPr[^>]*>.*?</sheetPr>|<sheetPr[^>]*/>', '', sx, flags=re.S)
    if sheetpr: sx = sx.replace('<dimension', sheetpr.group(0) + '<dimension', 1)
    # carry over the form-written rows A4:E7 (present in populated workbooks; empty in the template)
    for rr in range(4, 8):
        om = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rr, orig, re.S)
        if not om: continue
        keep = ''.join(c for c in re.findall(r'<c r="[A-E]%d"(?:\s[^>]*)?(?:/>|>.*?</c>)' % rr, om.group(1), re.S))
        if not keep: continue
        nm = re.search(r'(<row r="%d"[^>]*>)(.*?)(</row>)' % rr, sx, re.S)
        if nm: sx = sx[:nm.start()] + nm.group(1) + keep + nm.group(2) + nm.group(3) + sx[nm.end():]
    sx = re.sub(r'<drawing [^>]*/>', '', sx)
    drawing_rel = re.search(r'<drawing r:id="[^"]+"/>', orig).group(0)
    sx = sx.replace('</worksheet>', drawing_rel + '</worksheet>')
    sx = sx.replace('</sheetPr>', '<pageSetUpPr fitToPage="1"/></sheetPr>', 1) if '</sheetPr>' in sx else sx.replace('<sheetPr', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetPr', 1)
    sx = re.sub(r'<pageMargins[^>]*/>', lambda m: m.group(0) + '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>', sx, count=1)
    wbx = rd('xl/workbook.xml')
    if '_xlnm.Print_Area" localSheetId="%d"' % sheet_index not in wbx:
        pa = '<definedName name="_xlnm.Print_Area" localSheetId="%d">Summary!$A$1:$V$141</definedName>' % sheet_index
        wbx = wbx.replace('</definedNames>', pa + '</definedNames>') if '</definedNames>' in wbx else wbx.replace('</sheets>', '</sheets><definedNames>' + pa + '</definedNames>', 1)
        wr('xl/workbook.xml', wbx)
    if 'xmlns:r=' not in sx[:600]: sx = sx.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    minidom.parseString(sx)
    wr(sheet_part, sx)

    # --- charts: copy parts, register content types
    ct = rd('[Content_Types].xml')
    scharts = sorted(f for f in os.listdir(os.path.join(sd, 'xl/charts')) if f.startswith('chart'))
    chart_map = {}
    for i, f in enumerate(scharts):
        new = 'chart%d.xml' % (chart_start + i); chart_map[f] = new
        shutil.copy(os.path.join(sd, 'xl/charts', f), os.path.join(work, 'xl/charts', new))
        ct = ct.replace('</Types>', '<Override PartName="/xl/charts/%s" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>' % new)
    wr('[Content_Types].xml', ct)

    # --- drawing: append the scratch anchors to the existing drawing, remap rIds and shape ids
    sdraw = srd('xl/drawings/drawing1.xml'); srels = srd('xl/drawings/_rels/drawing1.xml.rels')
    sdraw = re.sub(r'<(/?)([A-Za-z]+)(?=[\s>/])', r'<\1xdr:\2', sdraw)   # default-namespace elements -> xdr: prefix
    sdraw = re.sub(r'name="Chart (\d+)"', lambda m: 'name="Summary Chart %s"' % m.group(1), sdraw)
    rid2chart = {}
    for rel in re.findall(r'<Relationship\b[^>]*/>', srels):
        i = re.search(r'\bId="([^"]+)"', rel); tg = re.search(r'\bTarget="[^"]*?(chart\d+\.xml)"', rel)
        if i and tg: rid2chart[i.group(1)] = tg.group(1)
    rels_part = os.path.dirname(drawing_part) + '/_rels/' + os.path.basename(drawing_part) + '.rels'
    bdraw = rd(drawing_part); brels = rd(rels_part)
    # park the form-managed "Chart 1" below the dashboard: it duplicates chart 1 with less detail, but the
    # Alternative Setup macro still maintains it, so it is moved rather than removed
    bdraw = re.sub(r'(<xdr:twoCellAnchor>)<xdr:from>.*?</xdr:from><xdr:to>.*?</xdr:to>',
                   r'\1<xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                   r'<xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                   % (VBACH - 1, VBACH + 16), bdraw, count=1, flags=re.S)
    next_rid = max(int(x) for x in re.findall(r'Id="rId(\d+)"', brels)) + 1
    anchors = re.findall(r'<xdr:(?:one|two)CellAnchor\b.*?</xdr:(?:one|two)CellAnchor>', sdraw, re.S)
    out = []; new_rels = []
    for n, a in enumerate(anchors):
        def rid(m):
            nonlocal next_rid
            tgt = chart_map[rid2chart[m.group(1)]]; new_id = 'rId%d' % next_rid; next_rid += 1
            new_rels.append('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/%s"/>' % (new_id, tgt))
            return 'r:id="%s"' % new_id
        a = re.sub(r'r:id="(rId\d+)"', rid, a)
        a = re.sub(r'(<xdr:cNvPr id=")\d+(")', lambda m: '%s%d%s' % (m.group(1), 101 + n, m.group(2)), a)
        if 'xmlns:c=' not in a: a = a.replace('<c:chart ', '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ', 1)
        if 'xmlns:a=' not in bdraw[:800] and 'xmlns:a=' not in a: a = a.replace('<xdr:graphicFrame', '<xdr:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"', 1)
        out.append(a)
    bdraw = bdraw.replace('</xdr:wsDr>', ''.join(out) + '</xdr:wsDr>')
    brels = brels.replace('</Relationships>', ''.join(new_rels) + '</Relationships>')
    minidom.parseString(bdraw)
    wr(drawing_part, bdraw); wr(rels_part, brels)
    shutil.rmtree(tmp)
    return len(scharts)

if __name__ == '__main__':
    build_scratch('/tmp/summary_scratch.xlsx'); print('scratch written')


# ---------------------------------------------------------------------------------------------
# Plain reference sheet (no charts): appended as the last sheet so no localSheetId shifts.
def _merge_styles(work, sd):
    """Append the scratch workbook's styles to the package styles; return the cellXfs offset."""
    rd = lambda p: open(os.path.join(work, p), encoding='utf-8').read()
    bs = rd('xl/styles.xml'); ss = open(os.path.join(sd, 'xl/styles.xml'), encoding='utf-8').read()
    def append(xml, tag, new_items):
        m = re.search(r'<%s count="(\d+)"' % tag, xml); cnt = int(m.group(1))
        xml = re.sub(r'<%s count="\d+"' % tag, '<%s count="%d"' % (tag, cnt + len(new_items)), xml, count=1)
        end = xml.index('</%s>' % tag); return xml[:end] + ''.join(new_items) + xml[end:], cnt
    nfs = _children_xml(ss, 'numFmts')
    base_nf_ids = [int(x) for x in re.findall(r'<numFmt numFmtId="(\d+)"', bs)]
    next_nf = max(base_nf_ids + [163]) + 1; nf_map = {}; new_nfs = []
    for nf in nfs:
        old = int(re.search(r'numFmtId="(\d+)"', nf).group(1)); nf_map[old] = next_nf
        new_nfs.append(re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % next_nf, nf)); next_nf += 1
    if new_nfs:
        if '<numFmts' in bs: bs, _ = append(bs, 'numFmts', new_nfs)
        else: bs = bs.replace('<fonts', '<numFmts count="%d">%s</numFmts><fonts' % (len(new_nfs), ''.join(new_nfs)), 1)
    bs, font_off = append(bs, 'fonts', _children_xml(ss, 'fonts'))
    bs, fill_off = append(bs, 'fills', _children_xml(ss, 'fills'))
    bs, border_off = append(bs, 'borders', _children_xml(ss, 'borders'))
    new_xfs = []
    for xf in _children_xml(ss, 'cellXfs'):
        xf = re.sub(r'fontId="(\d+)"', lambda m: 'fontId="%d"' % (int(m.group(1)) + font_off), xf)
        xf = re.sub(r'fillId="(\d+)"', lambda m: 'fillId="%d"' % (int(m.group(1)) + fill_off), xf)
        xf = re.sub(r'borderId="(\d+)"', lambda m: 'borderId="%d"' % (int(m.group(1)) + border_off), xf)
        xf = re.sub(r'numFmtId="(\d+)"', lambda m: 'numFmtId="%d"' % nf_map.get(int(m.group(1)), int(m.group(1))), xf)
        xf = re.sub(r'xfId="\d+"', 'xfId="0"', xf); new_xfs.append(xf)
    bs, xf_off = append(bs, 'cellXfs', new_xfs)
    minidom.parseString(bs)
    open(os.path.join(work, 'xl/styles.xml'), 'w', encoding='utf-8').write(bs)
    return xf_off


def add_plain_sheet(work, build_fn, sheet_name):
    """Build a sheet with openpyxl (build_fn(path) must create a workbook whose first sheet is the content)
    and append it to the unpacked package as the last worksheet. Returns the new sheet part name."""
    tmp = tempfile.mkdtemp(); scratch = os.path.join(tmp, 's.xlsx'); build_fn(scratch)
    sd = os.path.join(tmp, 'u'); zipfile.ZipFile(scratch).extractall(sd)
    rd = lambda p: open(os.path.join(work, p), encoding='utf-8').read()
    wr = lambda p, s: open(os.path.join(work, p), 'w', encoding='utf-8').write(s)
    xf_off = _merge_styles(work, sd)
    sx = open(os.path.join(sd, 'xl/worksheets/sheet1.xml'), encoding='utf-8').read()
    ssp = os.path.join(sd, 'xl/sharedStrings.xml')
    strings = [html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S))) for si in re.findall(r'<si>(.*?)</si>', open(ssp, encoding='utf-8').read(), re.S)] if os.path.exists(ssp) else []
    sx = re.sub(r'<c r="([A-Z]+\d+)"([^>]*?) t="s"([^>]*)><v>(\d+)</v></c>',
                lambda m: '<c r="%s"%s t="inlineStr"%s><is><t xml:space="preserve">%s</t></is></c>' % (m.group(1), m.group(2), m.group(3), html.escape(strings[int(m.group(4))], quote=False)), sx)
    sx = re.sub(r' s="(\d+)"', lambda m: ' s="%d"' % (int(m.group(1)) + xf_off), sx)
    sx = re.sub(r'dxfId="(\d+)"', lambda m: 'dxfId="%d"' % (int(m.group(1)) + dxf_off), sx)
    sx = re.sub(r'<pageSetup [^>]*/>', '', sx)
    sx = re.sub(r'<pageMargins[^>]*/>', lambda m: m.group(0) + '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>', sx, count=1)
    # exactly one sheetPr, carrying pageSetUpPr fitToPage (Excel rejects a second sheetPr element)
    sx = re.sub(r'<pageSetUpPr[^>]*/>', '', sx)
    if re.search(r'<sheetPr[^>]*/>', sx): sx = re.sub(r'<sheetPr([^>]*)/>', r'<sheetPr\1><pageSetUpPr fitToPage="1"/></sheetPr>', sx, count=1)
    elif '<sheetPr' in sx: sx = sx.replace('</sheetPr>', '<pageSetUpPr fitToPage="1"/></sheetPr>', 1)
    else: sx = sx.replace('<dimension', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension', 1)
    assert sx.count('<sheetPr') == 1
    minidom.parseString(sx)
    existing = [int(m) for m in re.findall(r'worksheets/sheet(\d+)\.xml', rd('xl/_rels/workbook.xml.rels'))]
    n = max(existing) + 1; part = 'xl/worksheets/sheet%d.xml' % n
    wr(part, sx)
    wbx = rd('xl/workbook.xml'); rels = rd('xl/_rels/workbook.xml.rels')
    rid = 'rId%d' % (max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels)) + 1)
    sid = max(int(x) for x in re.findall(r'sheetId="(\d+)"', wbx)) + 1
    wbx = wbx.replace('</sheets>', '<sheet name="%s" sheetId="%d" r:id="%s"/></sheets>' % (html.escape(sheet_name, quote=True), sid, rid), 1)
    rels = rels.replace('</Relationships>', '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/></Relationships>' % (rid, n))
    ct = rd('[Content_Types].xml').replace('</Types>', '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % part)
    wr('xl/workbook.xml', wbx); wr('xl/_rels/workbook.xml.rels', rels); wr('[Content_Types].xml', ct)
    return part


# ---------------------------------------------------------------------------------------------
# Style helpers for the design pass: append fonts/fills/borders/cellXfs/dxfs to the package styles.
def _append(xml, tag, items):
    """Append children to a counted styles element; returns (xml, index of the first new child)."""
    m = re.search(r'<%s count="(\d+)"' % tag, xml)
    if not m:  # element absent (e.g. no dxfs yet): create it before </styleSheet>
        xml = xml.replace('</styleSheet>', '<%s count="%d">%s</%s></styleSheet>' % (tag, len(items), ''.join(items), tag))
        return xml, 0
    cnt = int(m.group(1))
    xml = re.sub(r'<%s count="\d+"' % tag, '<%s count="%d"' % (tag, cnt + len(items)), xml, count=1)
    if '</%s>' % tag in xml:
        end = xml.index('</%s>' % tag); return xml[:end] + ''.join(items) + xml[end:], cnt
    xml = xml.replace('<%s count="%d"/>' % (tag, cnt + len(items)), '<%s count="%d">%s</%s>' % (tag, cnt + len(items), ''.join(items), tag), 1)
    return xml, cnt


def add_style(work, font=None, fill=None, border=None, alignment=None, numfmt=None):
    """Append one cellXfs entry (with any new font/fill/border/number format) to the package styles;
    returns its index."""
    p = os.path.join(work, 'xl/styles.xml'); s = open(p, encoding='utf-8').read()
    attrs = ['xfId="0"']
    if numfmt is None:
        attrs.insert(0, 'numFmtId="0"')
    else:
        used = [int(i) for i in re.findall(r'<numFmt numFmtId="(\d+)"', s)]
        nid = max(used + [163]) + 1
        nf = '<numFmt numFmtId="%d" formatCode="%s"/>' % (nid, numfmt)
        if '<numFmts' in s: s, _ = _append(s, 'numFmts', [nf])
        else: s = s.replace('<fonts', '<numFmts count="1">%s</numFmts><fonts' % nf, 1)
        attrs.insert(0, 'numFmtId="%d"' % nid); attrs.append('applyNumberFormat="1"')
    if font is not None:
        s, off = _append(s, 'fonts', [font]); attrs.append('fontId="%d"' % off); attrs.append('applyFont="1"')
    if fill is not None:
        s, off = _append(s, 'fills', [fill]); attrs.append('fillId="%d"' % off); attrs.append('applyFill="1"')
    if border is not None:
        s, off = _append(s, 'borders', [border]); attrs.append('borderId="%d"' % off); attrs.append('applyBorder="1"')
    body = ''
    if alignment is not None:
        attrs.append('applyAlignment="1"'); body = '<alignment %s/>' % alignment
    xf = '<xf %s>%s</xf>' % (' '.join(attrs), body) if body else '<xf %s/>' % ' '.join(attrs)
    s, idx = _append(s, 'cellXfs', [xf])
    minidom.parseString(s); open(p, 'w', encoding='utf-8').write(s)
    return idx


def add_dxf(work, dxf):
    """Append one differential format (conditional-formatting style); returns its dxfId."""
    p = os.path.join(work, 'xl/styles.xml'); s = open(p, encoding='utf-8').read()
    s, idx = _append(s, 'dxfs', [dxf])
    minidom.parseString(s); open(p, 'w', encoding='utf-8').write(s)
    return idx


FONT = lambda sz=10, b=False, i=False, color='FF1D2733': '<font>%s%s<sz val="%s"/><color rgb="%s"/><name val="Arial"/><family val="2"/></font>' % ('<b/>' if b else '', '<i/>' if i else '', sz, color)
FILL = lambda rgb: '<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>' % rgb
BORDER_BOX = '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>'
BORDER_BOTTOM = '<border><left/><right/><top/><bottom style="thin"><color rgb="FF2A78D6"/></bottom><diagonal/></border>'
