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
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter as L

GI = "'General Information'"
HMA_RGB, PCC_RGB, HMA2, PCC2 = '2A78D6', 'EB6834', '6DA7EC', 'F39C7A'
ALT_COLORS = [HMA_RGB, PCC_RGB, HMA2, PCC2]
NALT = 4          # rows 4..7, as the existing Summary/VBA assume
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
    button('A1', '◄ General Information', "'General Information'!D9")
    button('B1', 'Instructions', 'Instructions!B9', FILL_BTN2)
    ws.row_dimensions[1].height = 22
    hdr(3, 1, ['Alternative', 'Name', 'Initial Construction', 'Present Worth', 'Alternative Description'])
    ws.column_dimensions['A'].width = 22; ws.column_dimensions['B'].width = 32; ws.column_dimensions['C'].width = 17; ws.column_dimensions['D'].width = 17; ws.column_dimensions['E'].width = 34
    ws.column_dimensions['F'].width = 3
    for c in 'GHIJKLMNOPQRSTUV': ws.column_dimensions[c].width = 14
    ws.column_dimensions['G'].width = 20; ws.column_dimensions['H'].width = 24; ws.column_dimensions['P'].width = 30; ws.column_dimensions['Q'].width = 30

    # ---- 1. results table
    ws['G1'] = 'LCCA SUMMARY'; ws['G1'].font = F_T
    ws['G2'] = 'Everything from this column onward calculates from the alternative worksheets and updates by itself. Columns A:E and "Chart 1" are written by the Alternative Setup form as before.'; ws['G2'].font = F_N
    hdr(3, 7, ['Worksheet', 'Alternative', 'Type', 'Initial construction', 'Maintenance PW', 'Rehabilitation PW', 'Lost revenue PW', 'Salvage PW', 'Net present worth', 'vs. lowest NPW', 'Closure days in period', 'Runway availability'])
    for i in range(NALT):
        r = 4 + i; dbr = 4 + i; g = f'$G{r}'
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
        ws.cell(r, 16, f'=IF({g}="","",O{r}-MIN($O$4:$O$7))')
        dcol = L(DC + 2 + 2 * NALT + i)  # this alternative's closure-days column in the by-year block (period-aware)
        ws.cell(r, 17, f'=IF({g}="","",IF(IFERROR(INDIRECT("\'"&{g}&"\'!$F$2"),0)>0,SUM(${dcol}$41:${dcol}$71),IFERROR(SUM(INDIRECT("\'"&{g}&"\'!$F$4:$F$10")),0)))')
        ws.cell(r, 18, f'=IF({g}="","",1-Q{r}/({GI}!$D$33*365))')
        for c in range(10, 17): ws.cell(r, c).number_format = CUR
        ws.cell(r, 18).number_format = '0.00%'
        for c in range(7, 19): ws.cell(r, c).font = F_B; ws.cell(r, c).border = BOX
    S7 = f'${L(DC+1)}${12+20}:${L(DC+NALT)}${12+20}'  # 7.00% row of the sensitivity data block
    ws['G9'] = (f'=IF(COUNT($O$4:$O$7)=0,"No alternatives yet. Go to General Information and click Alternative Setup.",'
                f'"Lowest present worth: "&INDEX($H$4:$H$7,MATCH(MIN($O$4:$O$7),$O$4:$O$7,0))&IF(COUNT($O$4:$O$7)<2,"   |   only one alternative so far",IF(SMALL($O$4:$O$7,2)=MIN($O$4:$O$7),"   |   tied with the next alternative","   |   margin to next: "&TEXT(SMALL($O$4:$O$7,2)-MIN($O$4:$O$7),"$#,##0")))'
                f'&"   |   "&{GI}!$D$34&"% over "&{GI}!$D$33&" years   |   lost revenue: "&{GI}!$D$38)')
    ws['G9'].font = F_H; ws['G9'].fill = FILL_VERDICT
    for c in range(8, 19): ws.cell(9, c).fill = FILL_VERDICT
    S2 = f'${L(DC+1)}$12:${L(DC+NALT)}$12'  # 2.00% row (OMB A-94 real rate used by FAA PGL 22-01; 2.0% in 2026)
    low = 'INDEX($H$4:$H$7,MATCH(MIN($O$4:$O$7),$O$4:$O$7,0))'
    ws['G10'] = (f'=IF(COUNT($O$4:$O$7)<2,"",'
                 f'IF({low}=INDEX($H$4:$H$7,MATCH(SMALL({S2},COUNTIF({S2},0)+1),{S2},0)),"Same lowest-cost alternative at 2% (OMB A-94 real rate, FAA PGL 22-01)","At 2% (OMB A-94 real rate, FAA PGL 22-01) the lowest-cost alternative changes")'
                 f'&"   |   "&IF({low}=INDEX($H$4:$H$7,MATCH(SMALL({S7},COUNTIF({S7},0)+1),{S7},0)),"same at 7% (pre-2022 AIP rule)","changes at 7% (pre-2022 AIP rule): see chart 2"))')
    # SMALL(range, COUNTIF(range,0)+1) = smallest non-zero value: unused alternative columns hold 0 and must not win
    ws['G10'].font = F_N

    # ---- chart data (column W onward)
    d = lambda k: L(DC + k)
    ws.cell(1, DC, 'CHART DATA (calculated automatically from the table on the left; do not edit)').font = F_H
    ws.column_dimensions[d(0)].width = 24
    for k in range(1, 15): ws.column_dimensions[d(k)].width = 13
    # category block rows 3-9
    ws.cell(3, DC, 'Present worth by category').font = F_H
    hdr(4, DC, ['Category'] + [None] * NALT)
    for i in range(NALT): ws.cell(4, DC + 1 + i, f'=IF($H${4+i}="","Alt "&{i+1},$H${4+i})')
    cats = [('Initial construction', 10), ('Maintenance', 11), ('Rehabilitation', 12), ('Lost revenue', 13), ('Salvage', 14)]
    for k, (lab, col) in enumerate(cats):
        ws.cell(5 + k, DC, lab).font = F_DATA
        for i in range(NALT): c = ws.cell(5 + k, DC + 1 + i, f'=IF({L(col)}{4+i}="",0,{L(col)}{4+i})'); c.number_format = CUR; c.font = F_DATA
    # sensitivity block rows 10-36 (row 12 = 2.00%, row 32 = 7.00%)
    ws.cell(10, DC, 'Net present worth vs. discount rate').font = F_H
    hdr(11, DC, ['Rate (%)'] + [None] * NALT)
    for i in range(NALT): ws.cell(11, DC + 1 + i, f'={d(1+i)}$4')
    for k in range(25):
        r = 12 + k; c = ws.cell(r, DC, 2 + 0.25 * k); c.number_format = '0.00"%"'; c.font = F_DATA
        for i in range(NALT):
            g = f'$G${4+i}'; C, D = rng(g, 'C'), rng(g, 'D')
            c = ws.cell(r, DC + 1 + i, f'=IF({g}="",0,IFERROR($J${4+i}+SUMPRODUCT(({C}<={GI}!$D$33)*{D}/(1+${d(0)}{r}/100)^{C}),0))'); c.number_format = CUR; c.font = F_DATA
    SENS0, SENS1 = 12, 36
    # by-year block rows 39-71
    Y0 = 39
    ws.cell(Y0, DC, 'By calendar year: spend (undiscounted), cumulative discounted cost, closure days').font = F_H
    hdr(Y0 + 1, DC, ['Year', 'Calendar year'] + [None] * (3 * NALT))
    for i in range(NALT):
        ws.cell(Y0 + 1, DC + 2 + i, f'={d(1+i)}$4&" spend"'); ws.cell(Y0 + 1, DC + 2 + NALT + i, f'={d(1+i)}$4&" cumulative PW"'); ws.cell(Y0 + 1, DC + 2 + 2 * NALT + i, f'={d(1+i)}$4&" closure days"')
    for k in range(31):
        r = Y0 + 2 + k
        ws.cell(r, DC, k).font = F_DATA; c = ws.cell(r, DC + 1, f'={GI}!$D$25+{d(0)}{r}'); c.font = F_DATA
        for i in range(NALT):
            g = f'$G${4+i}'; B, C, D, E = rng(g, 'B'), rng(g, 'C'), rng(g, 'D'), rng(g, 'E'); yr = f'{d(0)}{r}'
            spend = f'IF({yr}>{GI}!$D$33,0,IF({yr}=0,$J${4+i},IFERROR(SUMIF({C},{yr},{D}),0)))'
            c = ws.cell(r, DC + 2 + i, f'=IF({g}="",0,{spend})'); c.number_format = CUR; c.font = F_DATA
            pv = f'IF({yr}>{GI}!$D$33,0,IF({yr}=0,$J${4+i},IFERROR(SUMIF({C},{yr},{E}),0)))'
            prev = '' if k == 0 else f'{d(2+NALT+i)}{r-1}+'
            c = ws.cell(r, DC + 2 + NALT + i, f'=IF({g}="",0,{prev}{pv})'); c.number_format = CUR; c.font = F_DATA
            days = f'IFERROR(SUMIFS({D},{C},{yr},{B},"*Indirect*")/INDIRECT("\'"&{g}&"\'!$F$2"),0)'
            c = ws.cell(r, DC + 2 + 2 * NALT + i, f'=IF({g}="",0,IF({yr}>{GI}!$D$33,0,{days}))'); c.number_format = '0'; c.font = F_DATA
    Y1 = Y0 + 2 + 30

    # ---- charts, two per row under the results table
    def style(ch, title, h=8.0, w=15.5, xt='Alternative', yt='Present worth ($)'):
        ch.title = title; ch.width = w; ch.height = h; ch.legend.position = 'b'; ch.y_axis.numFmt = '$#,##0,,"M"'; ch.y_axis.majorGridlines = None
        ch.x_axis.delete = False; ch.y_axis.delete = False
        ch.x_axis.title = xt; ch.y_axis.title = yt
    def colour(ch, line=False):
        for s, rgb in zip(ch.series, ALT_COLORS):
            if line: s.graphicalProperties.line.solidFill = rgb; s.graphicalProperties.line.width = 22000; s.marker.symbol = 'none'; s.smooth = False
            else: s.graphicalProperties.solidFill = rgb; s.graphicalProperties.line.solidFill = rgb
    # ---- RealCost-style comparison block (agency cost / user cost / total, present worth and EUAC)
    ws['G12'] = 'COMPARISON  (RealCost layout: agency cost, user cost, total; present worth and equivalent uniform annual cost)'; ws['G12'].font = F_T
    hdr(13, 7, ['Alternative', 'Agency cost PW', 'Agency EUAC', 'User cost PW (lost revenue)', 'User EUAC', 'Total PW', 'Total EUAC', 'vs. lowest ($)', 'vs. lowest (%)', 'Lowest?'])
    CRF = f'(({GI}!$D$34/100)*(1+{GI}!$D$34/100)^{GI}!$D$33/((1+{GI}!$D$34/100)^{GI}!$D$33-1))'
    for i in range(NALT):
        r = 14 + i; src = 4 + i; g = f'$G${src}'
        ws.cell(r, 7, f'=IF({g}="","",$H${src})')
        ws.cell(r, 8, f'=IF({g}="","",$J${src}+$K${src}+$L${src}+$N${src})')
        ws.cell(r, 9, f'=IF({g}="","",H{r}*{CRF})')
        ws.cell(r, 10, f'=IF({g}="","",$M${src})')
        ws.cell(r, 11, f'=IF({g}="","",J{r}*{CRF})')
        ws.cell(r, 12, f'=IF({g}="","",$O${src})')
        ws.cell(r, 13, f'=IF({g}="","",L{r}*{CRF})')
        ws.cell(r, 14, f'=IF({g}="","",$P${src})')
        ws.cell(r, 15, f'=IF({g}="","",IFERROR($P${src}/MIN($O$4:$O$7),0))')
        ws.cell(r, 16, f'=IF({g}="","",IF($O${src}=MIN($O$4:$O$7),"lowest",""))')
        for c in range(8, 15): ws.cell(r, c).number_format = CUR
        ws.cell(r, 15).number_format = '0.0%'
        for c in range(7, 17): ws.cell(r, c).font = F_B; ws.cell(r, c).border = BOX
    ws.row_dimensions[13].height = 27
    ws['G18'] = ('Agency cost = initial construction + maintenance + rehabilitation + salvage. User cost = lost airport revenue during runway closures, the airport-side counterpart of the user delay cost in the Caltrans/FHWA RealCost layout. '
                 'EUAC = PW x r(1+r)^P / ((1+r)^P - 1) at the discount rate and analysis period on General Information. "vs. lowest (%)" is the difference as a share of the lowest total PW.')
    ws['G18'].font = F_N
    ws['G20'] = 'CHARTS'; ws['G20'].font = F_T
    ch = BarChart(); ch.type = 'col'; ch.grouping = 'stacked'; ch.overlap = 100; ch.gapWidth = 60
    ch.add_data(Reference(ws, min_col=DC, max_col=DC + NALT, min_row=5, max_row=9), titles_from_data=True, from_rows=True)
    ch.set_categories(Reference(ws, min_col=DC + 1, max_col=DC + NALT, min_row=4, max_row=4))
    for s, rgb in zip(ch.series, ['4A4A4A', '8C8C8C', '646464', 'B8B6AE', 'DCDCDC']): s.graphicalProperties.solidFill = rgb; s.graphicalProperties.line.solidFill = rgb
    ch.x_axis.tickLblPos = 'low'; style(ch, '1. Present worth by category (salvage below zero)', xt='Alternative', yt='Present worth ($)'); ws.add_chart(ch, 'G22')
    ch = LineChart()
    ch.add_data(Reference(ws, min_col=DC + 1, max_col=DC + NALT, min_row=11, max_row=SENS1), titles_from_data=True); ch.set_categories(Reference(ws, min_col=DC, min_row=SENS0, max_row=SENS1))
    colour(ch, line=True); ch.x_axis.tickLblSkip = 4; ch.x_axis.numFmt = '0.00"%"'; style(ch, '2. Net present worth vs. discount rate (TDOT 3%, FAA 2%, pre-2022 rule 7%)', xt='Discount rate (%)', yt='Net present worth ($)'); ws.add_chart(ch, 'N22')
    ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'; ch.gapWidth = 40
    ch.add_data(Reference(ws, min_col=DC + 2, max_col=DC + 1 + NALT, min_row=Y0 + 1, max_row=Y1), titles_from_data=True); ch.set_categories(Reference(ws, min_col=DC + 1, min_row=Y0 + 2, max_row=Y1))
    colour(ch); ch.x_axis.tickLblSkip = 5; ch.x_axis.tickLblPos = 'low'; style(ch, '3. Expenditure stream by calendar year (undiscounted)', xt='Calendar year', yt='Spend, undiscounted ($)'); ws.add_chart(ch, 'G40')
    ch = LineChart()
    ch.add_data(Reference(ws, min_col=DC + 2 + NALT, max_col=DC + 1 + 2 * NALT, min_row=Y0 + 1, max_row=Y1), titles_from_data=True); ch.set_categories(Reference(ws, min_col=DC + 1, min_row=Y0 + 2, max_row=Y1))
    colour(ch, line=True); ch.x_axis.tickLblSkip = 5; style(ch, '4. Cumulative discounted cost', xt='Calendar year', yt='Cumulative discounted cost ($)'); ws.add_chart(ch, 'N40')
    ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'; ch.gapWidth = 40
    ch.add_data(Reference(ws, min_col=DC + 2 + 2 * NALT, max_col=DC + 1 + 3 * NALT, min_row=Y0 + 1, max_row=Y1), titles_from_data=True); ch.set_categories(Reference(ws, min_col=DC + 1, min_row=Y0 + 2, max_row=Y1))
    colour(ch); ch.x_axis.tickLblSkip = 5; style(ch, '5. Runway closure days by calendar year', xt='Calendar year', yt='Closure days'); ch.y_axis.numFmt = '0'; ws.add_chart(ch, 'G58')
    # ---- pavement section read back from the pay-item quantities (display only; no cost depends on it)
    PCF = '$J$81'
    AREA = f'{GI}!$D$26'; SHLD = f'{GI}!$D$27'
    ws['G80'] = 'PAVEMENT SECTION  (read back from the quantities already entered; nothing here changes a cost)'; ws['G80'].font = F_T
    ws['G81'] = 'Asphalt unit weight, pcf:'; ws['G81'].font = F_H
    c = ws['J81']; c.value = 145; c.number_format = '0'; c.font = F_B; c.fill = PatternFill('solid', fgColor='D9D9D9'); c.border = BOX
    ws['K81'] = ('Only asphalt needs an assumption: 145 pcf reproduces the Murfreesboro section to the hundredth of an inch. '
                 'Volume items convert directly, and concrete carries its thickness in the pay-item name.')
    ws['K81'].font = F_N
    hdr(82, 7, ['Alternative', 'Surface course (in)', 'Aggregate base (in)', 'Subbase (in)', 'Treated subgrade',
                'Total section (in)', 'Excavation (in)', 'Excavation check', None])
    hdr(82, 16, ['Section from the quantities', 'Same quantities over mainline + shoulder'])
    ws.row_dimensions[82].height = 27
    for i in range(NALT):
        r = 83 + i; g = f'$G${4+i}'
        B, C, D, E = rng(g, 'B', 13, 22), rng(g, 'C', 13, 22), rng(g, 'D', 13, 22), rng(g, 'E', 13, 22)
        tons = f'SUMIFS({E},{D},"TON")'
        cy = lambda pat: f'SUMIFS({E},{D},"C.Y.",{B},"{pat}")'
        agg = f'({cy("P-2*")}+{cy("P-3*")})'
        desc = f'IFERROR(INDEX({C},MATCH("Concrete Pavement*",{C},0)),"")'
        pcc = f'IFERROR(VALUE(TRIM(SUBSTITUTE(MID({desc},FIND(",",{desc})+1,99),"-inch",""))),0)'
        item = lambda pat: f'IFERROR(SUBSTITUTE(LEFT(INDEX({B},MATCH("{pat}",{B},0)),5),"-",""),"")'
        ws.cell(r, 7, f'=IF({g}="","",$H${4+i})')
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
    ws['G87'] = ('Thickness is not an input. Volume items give inches = 36 x C.Y. / mainline S.Y.; asphalt gives inches = 2,666.67 x tons / (pcf x mainline S.Y.); '
                 'items measured by area carry no implied thickness. The last column is the section as the quantities describe it: paste it into the alternative description so the two can never disagree. The read-back divides by the mainline area only; when a shoulder area is entered, the last column spreads the same quantities over mainline plus shoulder, which is the reading to use if the quantities were taken off both.')
    ws['G87'].font = F_N
    # chart data for the two section charts (columns W onward, below the by-year block)
    SEC = 74
    ws.cell(SEC, DC, 'PAVEMENT SECTION (inches): mainline, then the reading if the quantities also cover the shoulder').font = F_H
    hdr(SEC + 1, DC, ['Layer'] + [None] * NALT)
    for i in range(NALT): ws.cell(SEC + 1, DC + 1 + i, f'={d(1+i)}$4')
    layers = [('Subbase (P-154)', 'J'), ('Aggregate base', 'I'), ('Asphalt', 'H'), ('Concrete', 'H')]
    for k, (lab, col) in enumerate(layers):
        ws.cell(SEC + 2 + k, DC, lab).font = F_DATA
        for i in range(NALT):
            r83 = 83 + i; g = f'$G${4+i}'
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
    for anch, r0, title in [('G90', SEC + 2, '6. Pavement section, mainline (inches, surface on top)'),
                            ('N90', SEC + 9, '7. Pavement section with the shoulder (empty unless a shoulder area is entered)')]:
        ch = BarChart(); ch.type = 'col'; ch.grouping = 'stacked'; ch.overlap = 100; ch.gapWidth = 80
        ch.add_data(Reference(ws, min_col=DC, max_col=DC + NALT, min_row=r0, max_row=r0 + 3), titles_from_data=True, from_rows=True)
        ch.set_categories(Reference(ws, min_col=DC + 1, max_col=DC + NALT, min_row=r0 - 1, max_row=r0 - 1))
        for s2, rgb in zip(ch.series, LAYER_RGB): s2.graphicalProperties.solidFill = rgb; s2.graphicalProperties.line.solidFill = rgb
        style(ch, title, xt='Alternative', yt='Thickness (inches)'); ch.y_axis.numFmt = '0'; ch.x_axis.tickLblPos = 'low'
        ws.add_chart(ch, anch)
    ws['G108'] = ('Charts 6 and 7 stack the layers as they sit in the ground, surface on top. Chart 7 stays empty unless a shoulder area is entered on General Information, '
                  'and then shows the same quantities spread over mainline plus shoulder, which is the lower bound on each thickness.')
    ws['G108'].font = F_N
    ws['G77'] = "How to read: 1 shows where each alternative's cost sits; 2 whether the lowest-cost alternative holds at other discount rates (FAA now uses the OMB A-94 real rate, about 2%; 7% was the rule before 2022); 3 and 4 when the money is spent and when the higher first cost is paid back; 5 how often and how long the runway closes."; ws['G77'].font = F_N
    dxf_low = DifferentialStyle(fill=PatternFill(bgColor='FFDDEBF7'), font=Font(bold=True, color='FF1F3864'))
    for sqref, f in [('G4:R7', 'AND($G4<>"",COUNT($O$4:$O$7)>0,$O4=MIN($O$4:$O$7))'),
                     ('G14:P17', 'AND($G14<>"",COUNT($L$14:$L$17)>0,$L14=MIN($L$14:$L$17))')]:
        rule = Rule(type='expression', formula=[f], stopIfTrue=False, dxf=dxf_low)
        ws.conditional_formatting.add(sqref, rule)
    ws.freeze_panes = 'A4'
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
        pa = '<definedName name="_xlnm.Print_Area" localSheetId="%d">Summary!$A$1:$U$112</definedName>' % sheet_index
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
    # move the form-managed "Chart 1" out of the results table area, next to chart 5
    bdraw = re.sub(r'(<xdr:twoCellAnchor>)<xdr:from>.*?</xdr:from><xdr:to>.*?</xdr:to>',
                   r'\1<xdr:from><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>57</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>20</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>74</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>', bdraw, count=1, flags=re.S)
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


def add_style(work, font=None, fill=None, border=None, alignment=None):
    """Append one cellXfs entry (with any new font/fill/border) to the package styles; returns its index."""
    p = os.path.join(work, 'xl/styles.xml'); s = open(p, encoding='utf-8').read()
    attrs = ['numFmtId="0"', 'xfId="0"']
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
BORDER_BOTTOM = '<border><left/><right/><top/><bottom style="thin"><color rgb="FF2A78D6"/></bottom><diagonal/></border>'
