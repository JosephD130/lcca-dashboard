#!/usr/bin/env python3
"""Builds TDOT_LCCA_Decision_Workbook.xlsx: one live, scenario-capable dashboard sheet per
project (MBT, SRB) plus a Monte Carlo sheet. All results are formulas driven by the parameter
block at the top of each sheet; the Monte Carlo sheet holds simulation output (values)."""
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, ScatterChart, BubbleChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.properties import PageSetupProperties

OUT = '/tmp/claude-0/-home-user-lcca-dashboard/444ca227-15d7-51d4-a51e-a0ae3dbe81ed/scratchpad/TDOT_LCCA_Decision_Workbook.xlsx'

# ----------------------------------------------------------------- data from the workbooks (2026-05-21 runs)
PROJECTS = {
 'MBT': dict(title='Murfreesboro Municipal (MBT) - Runway 18-36 Reconstruction', daily=8841.38, area=52777.7, markings=5433, year0=2027,
   hma=dict(name='Alt 1 - New HMA', desc='5" P-401 on 16" P-209', initial=5306576.53, salv_frac=0.125, salv_basis='rehab',
     events=[('Maintenance 1',4,482829.92,6),('Maintenance 2',8,482829.92,6),('Maintenance 3',12,547129.30,8),('Maintenance 4',16,561004.48,8),
             ('Rehabilitation 1',20,2302557.72,15),('Maintenance 5',24,482829.92,6),('Maintenance 6',28,577147.05,8)]),
   pcc=dict(name='Alt 2 - New PCC', desc='8" P-501 on 6" P-209', initial=8410244.97, salv_frac=0.25, salv_basis='initial',
     events=[('Maintenance 1',19,306563.68,9),('Rehabilitation 1',27,428405.86,18)])),
 'SRB': dict(title='Upper Cumberland Regional (SRB) - Runway Reconstruction', daily=4344.12, area=74488.89, markings=22052, year0=2027,
   hma=dict(name='Alt 2 - New HMA', desc='4" P-401 + 5" P-403 on 6" P-209', initial=9009181.19, salv_frac=0.125, salv_basis='rehab',
     events=[('Maintenance 1',4,714698.49,8),('Maintenance 2',8,714698.49,8),('Maintenance 3',12,805448.74,10),('Maintenance 4',16,825031.76,10),
             ('Rehabilitation 1',20,3270106.16,21),('Maintenance 5',24,714698.49,8),('Maintenance 6',28,847814.90,10)]),
   pcc=dict(name='Alt 1 - New PCC', desc='11" P-501 on 6" P-209 + P-403 base', initial=15445398.25, salv_frac=0.25, salv_basis='initial',
     events=[('Maintenance 1',19,432674.93,11),('Rehabilitation 1',27,604639.40,20)])),
}
PROJECTS['TEMPLATE']=dict(title='NEW PROJECT (template) - copy this sheet, rename it, fill the yellow cells', daily=0, area=0, markings=0, year0=2027,
   hma=dict(name='Alt 1 - New HMA', desc='fill in', initial=0, salv_frac=0.125, salv_basis='rehab',
     events=[('Maintenance 1',4,0,0),('Maintenance 2',8,0,0),('Maintenance 3',12,0,0),('Maintenance 4',16,0,0),('Rehabilitation 1',20,0,0),('Maintenance 5',24,0,0),('Maintenance 6',28,0,0)]),
   pcc=dict(name='Alt 2 - New PCC', desc='fill in', initial=0, salv_frac=0.25, salv_basis='initial',
     events=[('Maintenance 1',19,0,0),('Rehabilitation 1',27,0,0)]))

def salvage_base(alt):
    basis = alt['initial'] if alt['salv_basis']=='initial' else [c for n,y,c,d in alt['events'] if 'Rehab' in n][0]
    return -basis*alt['salv_frac']

# ----------------------------------------------------------------- python engine (mirror of the sheet formulas, used to cross-check and for Monte Carlo)
def npw(alt, rate, P, daily, cm=1, dm=1, sm=1, rf=1, salv=None):
    r=rate/100; tot=alt['initial']*cm
    for n,y,c,d in alt['events']:
        if y<=P: tot += (c*cm + d*dm*daily*rf)/(1+r)**y
    s = salvage_base(alt) if salv is None else salv
    tot += s*cm*sm/(1+r)**P
    return tot

# ----------------------------------------------------------------- styles
F_TITLE=Font(name='Arial',size=14,bold=True); F_H=Font(name='Arial',size=10,bold=True); F_B=Font(name='Arial',size=10)
F_IN=Font(name='Arial',size=10,color='0000FF'); F_NOTE=Font(name='Arial',size=9,italic=True,color='595959')
FILL_IN=PatternFill('solid',fgColor='FFFF99'); FILL_H=PatternFill('solid',fgColor='D9D9D9'); FILL_HMA=PatternFill('solid',fgColor='DCE9F9'); FILL_PCC=PatternFill('solid',fgColor='FBE3D6')
CUR='$#,##0;($#,##0);-'; PCT='0.0%'; thin=Side(style='thin',color='BFBFBF'); BOX=Border(top=thin,bottom=thin,left=thin,right=thin)
HMA_RGB='2A78D6'; PCC_RGB='EB6834'; HMA2='6DA7EC'; PCC2='F39C7A'; GRAY='B8B6AE'

def hdr(ws, row, col, labels, fill=FILL_H):
    for i,l in enumerate(labels):
        c=ws.cell(row,col+i,l); c.font=F_H; c.fill=fill; c.alignment=Alignment(wrap_text=True,vertical='center'); c.border=BOX

def style_chart(ch, title, ytitle=None, xtitle=None, w=18, h=9):
    ch.title=title; ch.width=w; ch.height=h; ch.legend.position='b'
    if ytitle: ch.y_axis.title=ytitle
    if xtitle: ch.x_axis.title=xtitle
    ch.y_axis.numFmt='$#,##0,,"M"'; ch.y_axis.majorGridlines=None
    ch.x_axis.delete=False; ch.y_axis.delete=False

def color_series(s, rgb, line=False, width=None):
    if line:
        s.graphicalProperties.line.solidFill=rgb; s.graphicalProperties.line.width=(width or 22000); s.smooth=False; s.marker.symbol='none'
    else:
        s.graphicalProperties.solidFill=rgb; s.graphicalProperties.line.solidFill=rgb

# ----------------------------------------------------------------- project sheet
def build_project(wb, key, pr):
    ws=wb.create_sheet(key)
    ws.sheet_view.showGridLines=False
    ws.column_dimensions['A'].width=30
    for c in 'BCDEFGHIJKLMNOPQR': ws.column_dimensions[c].width=14
    ws['A1']=pr['title']; ws['A1'].font=F_TITLE
    ws['A2']='Life-cycle cost comparison with runway-closure lost revenue. Yellow cells are inputs; everything else recalculates.'; ws['A2'].font=F_NOTE

    # ---- parameters
    params=[('Discount rate (%)',3,'TDOT default 3; FAA AIP handbook requires 7 for AIP-funded LCCA'),
            ('Analysis period (years)',30,'TDOT policy 30; FAA AIP caps project life at 20'),
            ('Construction year',pr['year0'],''),
            ('Airport daily revenue ($/day)',pr['daily'],'Sum of the seven RevenueData columns for this airport (workbook method)'),
            ('Share of daily revenue actually lost during a closure',1.0,'1.0 = workbook method (gross fuel sales and tenant rent count as lost). Try 0.3 to 0.5 for margin-based revenue'),
            ('Closure-days multiplier',1.0,'Scales every closure duration (production-rate defaults from the workbook)'),
            ('HMA cost multiplier',1.0,'Bid-price uncertainty on all HMA direct costs'),
            ('PCC cost multiplier',1.0,'Bid-price uncertainty on all PCC direct costs'),
            ('Salvage multiplier',1.0,'0 = no salvage credit; 1 = workbook policy (PCC 25% of initial, HMA 12.5% of one mill-and-overlay)'),
            ('Mainline area (SY)',pr['area'],''),
            ('HMA salvage fraction',pr['hma']['salv_frac'],'TDOT policy: 2 of 16 years of the mill-and-overlay = 0.125'),
            ('HMA salvage basis','Rehabilitation' if pr['hma']['salv_basis']=='rehab' else 'Initial','Initial or Rehabilitation (which cost the fraction applies to)'),
            ('PCC salvage fraction',pr['pcc']['salv_frac'],'TDOT policy: 10 of 40 years of initial construction = 0.25'),
            ('PCC salvage basis','Rehabilitation' if pr['pcc']['salv_basis']=='rehab' else 'Initial','Initial or Rehabilitation')]
    hdr(ws,4,1,['Parameter','Value','Note'])
    for i,(l,v,n) in enumerate(params):
        r=5+i; ws.cell(r,1,l).font=F_B; c=ws.cell(r,2,v); c.font=F_IN; c.fill=FILL_IN; c.border=BOX; ws.cell(r,3,n).font=F_NOTE
    ws['B5'].number_format='0.00'; ws['B8'].number_format='$#,##0.00'; ws['B9'].number_format='0%'; ws['B14'].number_format='#,##0'
    ws['B15'].number_format='0.0%'; ws['B17'].number_format='0.0%'
    from openpyxl.worksheet.datavalidation import DataValidation
    dv=DataValidation(type='list',formula1='"Initial,Rehabilitation"',allow_blank=False); ws.add_data_validation(dv); dv.add('B16'); dv.add('B18')
    ws['A20']='ACTIVITY TABLES: type the base direct cost and closure days for each policy activity (yellow). Years follow the TDOT Maintenance Policies. Copy them from the framework workbook alternative sheets (column D of the NPW table and cells F4:F10).'; ws['A20'].font=F_NOTE
    RATE,PER,Y0,DAILY,RF,DM,CM_H,CM_P,SM,AREA='$B$5','$B$6','$B$7','$B$8','$B$9','$B$10','$B$11','$B$12','$B$13','$B$14'

    # ---- activity tables (salvage kept outside the table so scenarios can move its year)
    tables={}
    def activity_table(alt, col0, cm, fill, sf, sb):
        c=lambda k: L(col0+k)
        r0=24
        ws.cell(22,col0,alt['name']+'  ('+alt['desc']+')').font=F_H
        hdr(ws,23,col0,['Activity','Year','Direct cost (base $)','Closure days (base)','Direct cost (scenario)','Lost revenue (scenario)','Present worth'],fill)
        rows=[('Initial construction',0,alt['initial'],0)]+[(n,y,cc,d) for n,y,cc,d in alt['events']]
        for i,(n,y,cc,d) in enumerate(rows):
            r=r0+i
            ws.cell(r,col0,n).font=F_B; ws.cell(r,col0+1,y).font=F_B
            for k,v in ((2,cc),(3,d)):
                cell=ws.cell(r,col0+k,v); cell.font=F_IN; cell.fill=FILL_IN; cell.border=BOX
            ws.cell(r,col0+2).number_format=CUR
            ws.cell(r,col0+4,f'={c(2)}{r}*{cm}').number_format=CUR
            ws.cell(r,col0+5,f'={c(3)}{r}*{DM}*{DAILY}*{RF}').number_format=CUR
            ws.cell(r,col0+6,f'=IF({c(1)}{r}>{PER},0,({c(4)}{r}+{c(5)}{r})/(1+{RATE}/100)^{c(1)}{r})').number_format=CUR
        rl=r0+len(rows)-1
        rs=rl+1
        ws.cell(rs,col0,'Salvage value (at end of period)').font=F_B
        ws.cell(rs,col0+1,f'={PER}')
        ws.cell(rs,col0+2,f'=-{sf}*IF({sb}="Initial",{c(2)}{r0},SUMPRODUCT(ISNUMBER(SEARCH("Rehab",{c(0)}{r0+1}:{c(0)}{rl}))*{c(2)}{r0+1}:{c(2)}{rl}))').number_format=CUR
        ws.cell(rs,col0+4,f'={c(2)}{rs}*{cm}*{SM}').number_format=CUR
        ws.cell(rs,col0+6,f'={c(4)}{rs}/(1+{RATE}/100)^{c(1)}{rs}').number_format=CUR
        ws.cell(rs+1,col0,'Net present worth').font=F_H
        ws.cell(rs+1,col0+6,f'=SUM({c(6)}{r0}:{c(6)}{rs})').number_format=CUR; ws.cell(rs+1,col0+6).font=F_H
        ws.cell(rs+2,col0,'Salvage = fraction x basis from the parameter block (TDOT Maintenance Policies). Base costs and closure days are inputs.').font=F_NOTE
        return dict(r0=r0,rl=rl,rs=rs,rn=rs+1,col0=col0,c=c,cm=cm)
    tables['H']=activity_table(pr['hma'],1,CM_H,FILL_HMA,'$B$15','$B$16')
    tables['P']=activity_table(pr['pcc'],10,CM_P,FILL_PCC,'$B$17','$B$18')
    rowNext=max(tables['H']['rn'],tables['P']['rn'])+4

    # generic NPW formula with explicit parameters (cell refs or literals)
    def NPW(t, rate=RATE, P=PER, cm=None, dm=DM, sm=SM, rf=RF, daily=DAILY):
        c=t['c']; cm=cm or t['cm']; r0,rl,rs=t['r0'],t['rl'],t['rs']
        yr=f'{c(1)}{r0}:{c(1)}{rl}'; D=f'{c(2)}{r0}:{c(2)}{rl}'; E=f'{c(3)}{r0}:{c(3)}{rl}'
        return (f'SUMPRODUCT(({yr}<={P})*({D}*{cm}+{E}*{dm}*{daily}*{rf})/(1+{rate}/100)^{yr})'
                f'+{c(2)}{rs}*{cm}*{sm}/(1+{rate}/100)^{P}')
    def CAT(t, cat):  # PW by category (scenario columns)
        c=t['c']; r0,rl,rs=t['r0'],t['rl'],t['rs']
        if cat=='initial': return f'{c(6)}{r0}'
        if cat=='salvage': return f'{c(6)}{rs}'
        if cat=='lost':    return f'SUMPRODUCT(({c(1)}{r0}:{c(1)}{rl}<={PER})*{c(5)}{r0}:{c(5)}{rl}/(1+{RATE}/100)^{c(1)}{r0}:{c(1)}{rl})'
        key={'maint':'Maintenance','rehab':'Rehabilitation'}[cat]
        return f'SUMPRODUCT(({c(1)}{r0+1}:{c(1)}{rl}<={PER})*ISNUMBER(SEARCH("{key}",{c(0)}{r0+1}:{c(0)}{rl}))*{c(4)}{r0+1}:{c(4)}{rl}/(1+{RATE}/100)^{c(1)}{r0+1}:{c(1)}{rl})'
    def DAYS(t): c=t['c']; return f'SUMPRODUCT(({c(1)}{t["r0"]}:{c(1)}{t["rl"]}<={PER})*{c(3)}{t["r0"]}:{c(3)}{t["rl"]})*{DM}'

    # ---- results block
    R=rowNext
    ws.cell(R,1,'RESULTS').font=F_TITLE
    hdr(ws,R+1,1,['Measure','HMA','PCC','PCC - HMA'])
    measures=[('Initial construction (scenario)',lambda t:f'={t["c"](4)}{t["r0"]}',CUR),
              ('Maintenance PW',lambda t:'='+CAT(t,'maint'),CUR),('Rehabilitation PW',lambda t:'='+CAT(t,'rehab'),CUR),
              ('Lost revenue PW',lambda t:'='+CAT(t,'lost'),CUR),('Salvage PW',lambda t:'='+CAT(t,'salvage'),CUR),
              ('Net present worth',lambda t:f'={t["c"](6)}{t["rn"]}',CUR),
              ('Equivalent uniform annual cost',lambda t:f'={t["c"](6)}{t["rn"]}*({RATE}/100)*(1+{RATE}/100)^{PER}/((1+{RATE}/100)^{PER}-1)',CUR),
              ('NPW per SY of mainline',lambda t:f'=IFERROR({t["c"](6)}{t["rn"]}/{AREA},0)','$#,##0.00'),
              ('Closure days in period',lambda t:'='+DAYS(t),'0'),
              ('Runway availability over period',lambda t:f'=1-({DAYS(t)})/({PER}*365)','0.00%'),
              ('Lost revenue, undiscounted',lambda t:f'=SUMPRODUCT(({t["c"](1)}{t["r0"]}:{t["c"](1)}{t["rl"]}<={PER})*{t["c"](5)}{t["r0"]}:{t["c"](5)}{t["rl"]})',CUR)]
    for i,(lab,fn,fmt) in enumerate(measures):
        r=R+2+i; ws.cell(r,1,lab).font=F_B
        ws.cell(r,2,fn(tables['H'])).number_format=fmt; ws.cell(r,3,fn(tables['P'])).number_format=fmt
        ws.cell(r,4,f'=C{r}-B{r}').number_format=fmt
        for cc in range(1,5): ws.cell(r,cc).border=BOX
    rNPW=R+2+5
    rv=R+2+len(measures)
    ws.cell(rv,1,'Lower-cost alternative').font=F_H
    ws.cell(rv,2,f'=IF(C{rNPW}<B{rNPW},"PCC","HMA")').font=F_H
    ws.cell(rv,3,f'=ABS(D{rNPW})').number_format=CUR
    ws.cell(rv,4,f'=IFERROR(ABS(D{rNPW})/MAX(B{rNPW},C{rNPW}),0)').number_format=PCT
    ws.cell(rv,5,f'=IF(D{rv}<0.1,"Margin inside estimating noise (<10%): treat as a tie on cost; decide on closures and constructability","Margin outside estimating noise")').font=F_NOTE
    ws.cell(rv+1,1,'Same winner at 7% (FAA AIP)?').font=F_H
    ws.cell(rv+1,2,f'=IF(({NPW(tables["P"],rate=7)})<({NPW(tables["H"],rate=7)}),"PCC","HMA")')
    ws.cell(rv+1,3,f'=IF(B{rv+1}=B{rv},"Yes","No - decision depends on the discount rate")').font=F_H

    # ---- break-even block
    BE=rv+4
    ws.cell(BE,1,'BREAK-EVEN VALUES (where PCC and HMA present worth are equal)').font=F_TITLE
    hdr(ws,BE+1,1,['Input','Break-even value','Current value','Reading'])
    tH,tP=tables['H'],tables['P']
    pvdays=lambda t:f'SUMPRODUCT(({t["c"](1)}{t["r0"]}:{t["c"](1)}{t["rl"]}<={PER})*{t["c"](3)}{t["r0"]}:{t["c"](3)}{t["rl"]}*{DM}*{RF}/(1+{RATE}/100)^{t["c"](1)}{t["r0"]}:{t["c"](1)}{t["rl"]})'
    pvdirect=lambda t:f'(SUMPRODUCT(({t["c"](1)}{t["r0"]}:{t["c"](1)}{t["rl"]}<={PER})*{t["c"](4)}{t["r0"]}:{t["c"](4)}{t["rl"]}/(1+{RATE}/100)^{t["c"](1)}{t["r0"]}:{t["c"](1)}{t["rl"]})+{t["c"](6)}{t["rs"]})'
    be_d=f'-(({pvdirect(tP)})-({pvdirect(tH)}))/(({pvdays(tP)})-({pvdays(tH)}))'
    be=[('Airport daily revenue ($/day)',f'=IFERROR(IF({be_d}<=0,"none (no positive value ties them)",{be_d}),"n/a")','$#,##0',f'={DAILY}','Daily revenue at which lost-revenue exposure alone flips the decision'),
        ('PCC cost multiplier',f'=IFERROR(({tH["c"](6)}{tH["rn"]}-{CAT(tP,"lost")})/(({pvdirect(tP)})/{CM_P}),"n/a")','0.00',f'={CM_P}','PCC bid level (x base estimate) at which the two tie'),
        ('HMA cost multiplier',f'=IFERROR(({tP["c"](6)}{tP["rn"]}-{CAT(tH,"lost")})/(({pvdirect(tH)})/{CM_H}),"n/a")','0.00',f'={CM_H}','HMA bid level (x base estimate) at which the two tie'),
        ('PCC salvage credit (% of initial)',f'=IFERROR(({tP["c"](6)}{tP["rn"]}-{tP["c"](6)}{tP["rs"]}-{tH["c"](6)}{tH["rn"]})/({tP["c"](4)}{tP["r0"]}/(1+{RATE}/100)^{PER}),"n/a")','0.0%',f'=IFERROR(-{tP["c"](2)}{tP["rs"]}*{SM}/{tP["c"](2)}{tP["r0"]},0)','Salvage fraction PCC needs to tie; compare with the 25% policy')]
    for i,(lab,f,fmt,cur,note) in enumerate(be):
        r=BE+2+i; ws.cell(r,1,lab).font=F_B; ws.cell(r,2,f).number_format=fmt; ws.cell(r,3,cur).number_format=fmt; ws.cell(r,4,note).font=F_NOTE
        for cc in range(1,4): ws.cell(r,cc).border=BOX
    rBE_rate=BE+2+len(be)
    ws.cell(rBE_rate,1,'Discount rate (%)').font=F_B; ws.cell(rBE_rate,3,f'={RATE}').number_format='0.00'; ws.cell(rBE_rate,4,'First rate in the 2-8% sweep where the winner changes (from the sensitivity table)').font=F_NOTE

    # ---- sensitivity table (rate sweep)
    S0=rBE_rate+3
    ws.cell(S0,1,'DISCOUNT-RATE SENSITIVITY').font=F_TITLE
    hdr(ws,S0+1,1,['Rate (%)','NPW HMA','NPW PCC','PCC - HMA','Winner'])
    rates=[2+0.25*i for i in range(25)]
    for i,rt in enumerate(rates):
        r=S0+2+i; ws.cell(r,1,rt).number_format='0.00'
        ws.cell(r,2,'='+NPW(tH,rate=f'$A{r}')).number_format=CUR; ws.cell(r,3,'='+NPW(tP,rate=f'$A{r}')).number_format=CUR
        ws.cell(r,4,f'=C{r}-B{r}').number_format=CUR; ws.cell(r,5,f'=IF(D{r}<0,"PCC","HMA")')
    S1=S0+1+len(rates)
    ws.cell(rBE_rate,2,f'=IFERROR(INDEX($A${S0+2}:$A${S1},MATCH(TRUE,INDEX($E${S0+2}:$E${S1}<>$E${S0+2},0),0)),"none in 2-8%")').number_format='0.00'
    ws.conditional_formatting.add(f'D{S0+2}:D{S1}',ColorScaleRule(start_type='min',start_color='F8CDB9',mid_type='num',mid_value=0,mid_color='FFFFFF',end_type='max',end_color='B7D3F6'))

    # ---- by-year table
    Y0=S1+3
    ws.cell(Y0,1,'EXPENDITURE BY CALENDAR YEAR').font=F_TITLE
    hdr(ws,Y0+1,1,['Year','Calendar year','HMA direct','HMA lost revenue','PCC direct','PCC lost revenue','HMA cumulative PW','PCC cumulative PW','HMA closure days','PCC closure days'])
    for k in range(31):
        r=Y0+2+k; ws.cell(r,1,k); ws.cell(r,2,f'=$B$7+A{r}')
        for j,(t,colk) in enumerate([(tH,4),(tH,5),(tP,4),(tP,5)]):
            c=t['c']; rng=f'{c(1)}{t["r0"]}:{c(1)}{t["rl"]}'
            f=f'=IF(A{r}>{PER},0,SUMIF({rng},A{r},{c(colk)}{t["r0"]}:{c(colk)}{t["rl"]})' + (f'+IF(A{r}={PER},{c(4)}{t["rs"]},0))' if colk==4 else ')')
            ws.cell(r,3+j,f).number_format=CUR
        for j,t in enumerate([tH,tP]):
            c=t['c']; rng=f'{c(1)}{t["r0"]}:{c(1)}{t["rl"]}'
            pv=f'IF(A{r}>{PER},0,SUMIF({rng},A{r},{c(6)}{t["r0"]}:{c(6)}{t["rl"]})+IF(A{r}={PER},{c(6)}{t["rs"]},0))'
            ws.cell(r,7+j,f'={pv}' if k==0 else f'={L(7+j)}{r-1}+{pv}').number_format=CUR
            ws.cell(r,9+j,f'=IF(A{r}>{PER},0,SUMIF({rng},A{r},{c(3)}{t["r0"]}:{c(3)}{t["rl"]})*{DM})')
    Y1=Y0+2+30
    ws.cell(Y1+1,1,'Crossover year (cumulative PW lines cross)').font=F_B
    ws.cell(Y1+1,2,f'=IFERROR(INDEX($B${Y0+3}:$B${Y1},MATCH(TRUE,INDEX(SIGN($H${Y0+3}:$H${Y1}-$G${Y0+3}:$G${Y1})<>SIGN($H${Y0+2}-$G${Y0+2}),0),0)),"no crossover")')

    # ---- tornado
    T0=Y1+4
    ws.cell(T0,1,'WHAT COULD FLIP THE ANSWER (one input at a time, PCC - HMA present worth)').font=F_TITLE
    hdr(ws,T0+1,1,['Input moved','Low case','High case','Low PCC-HMA','High PCC-HMA','Swing','Winner changes?'])
    base=f'$D${rNPW}'
    torn=[('Discount rate','2%','8%',f'({NPW(tP,rate=2)})-({NPW(tH,rate=2)})',f'({NPW(tP,rate=8)})-({NPW(tH,rate=8)})'),
          ('PCC salvage credit','0','policy',f'({NPW(tP,sm=0)})-({NPW(tH)})',base),
          ('HMA salvage credit','0','policy',f'({NPW(tP)})-({NPW(tH,sm=0)})',base),
          ('PCC direct costs','-20%','+20%',f'({NPW(tP,cm=f"({CM_P}*0.8)")})-({NPW(tH)})',f'({NPW(tP,cm=f"({CM_P}*1.2)")})-({NPW(tH)})'),
          ('HMA direct costs','-20%','+20%',f'({NPW(tP)})-({NPW(tH,cm=f"({CM_H}*0.8)")})',f'({NPW(tP)})-({NPW(tH,cm=f"({CM_H}*1.2)")})'),
          ('Closure durations','x0.5','x2',f'({NPW(tP,dm=f"({DM}*0.5)")})-({NPW(tH,dm=f"({DM}*0.5)")})',f'({NPW(tP,dm=f"({DM}*2)")})-({NPW(tH,dm=f"({DM}*2)")})'),
          ('Lost revenue','excluded','included',f'({NPW(tP,rf=0)})-({NPW(tH,rf=0)})',base),
          ('Analysis period','20 yr (FAA)','30 yr',f'({NPW(tP,P=20)})-({NPW(tH,P=20)})',base)]
    for i,(lab,lo,hi,flo,fhi) in enumerate(torn):
        r=T0+2+i; ws.cell(r,1,lab).font=F_B; ws.cell(r,2,lo); ws.cell(r,3,hi)
        ws.cell(r,4,'='+flo).number_format=CUR; ws.cell(r,5,'='+fhi).number_format=CUR
        ws.cell(r,6,f'=ABS(E{r}-D{r})').number_format=CUR; ws.cell(r,7,f'=IF(SIGN(D{r})<>SIGN(E{r}),"YES","no")')
    T1=T0+1+len(torn)
    ws.cell(T1+1,1,f'Base case PCC - HMA').font=F_B; ws.cell(T1+1,4,f'={base}').number_format=CUR
    ws.conditional_formatting.add(f'G{T0+2}:G{T1}',CellIsRule(operator='equal',formula=['"YES"'],font=Font(color='C00000',bold=True)))

    # ---- scenario scorecard
    C0=T1+4
    ws.cell(C0,1,'SCENARIO SCORECARD').font=F_TITLE
    hdr(ws,C0+1,1,['Scenario','Rate %','Period','HMA cost x','PCC cost x','Closure x','Salvage x','Revenue share lost','NPW HMA','NPW PCC','PCC - HMA','Winner'])
    scen=[('Base: TDOT policy (3%, 30 yr)',3,30,1,1,1,1,1),('FAA AIP funded (7%, 20 yr)',7,20,1,1,1,1,1),('TDOT 3% but 20-yr life',3,20,1,1,1,1,1),
          ('High rate 5%',5,30,1,1,1,1,1),('No lost-revenue term',3,30,1,1,1,1,0),('Revenue loss at 40% of gross',3,30,1,1,1,1,0.4),
          ('No salvage credit',3,30,1,1,1,0,1),('PCC bids +20%',3,30,1,1.2,1,1,1),('HMA bids +20%',3,30,1.2,1,1,1,1),
          ('Closures twice as long',3,30,1,1,2,1,1),('Closures half as long',3,30,1,1,0.5,1,1),('Stress: 7%, no salvage, no lost revenue',7,30,1,1,1,0,0)]
    for i,(lab,*v) in enumerate(scen):
        r=C0+2+i; ws.cell(r,1,lab).font=F_B
        for j,x in enumerate(v): ws.cell(r,2+j,x)
        ws.cell(r,9,'='+NPW(tH,rate=f'$B{r}',P=f'$C{r}',cm=f'$D{r}',dm=f'$F{r}',sm=f'$G{r}',rf=f'$H{r}')).number_format=CUR
        ws.cell(r,10,'='+NPW(tP,rate=f'$B{r}',P=f'$C{r}',cm=f'$E{r}',dm=f'$F{r}',sm=f'$G{r}',rf=f'$H{r}')).number_format=CUR
        ws.cell(r,11,f'=J{r}-I{r}').number_format=CUR; ws.cell(r,12,f'=IF(K{r}<0,"PCC","HMA")')
    C1=C0+1+len(scen)
    ws.cell(C1+1,1,'Note: scenario rows use their own parameters, not the yellow inputs. A 20-year period applies the same salvage fraction at year 20 (conservative simplification).').font=F_NOTE
    ws.conditional_formatting.add(f'K{C0+2}:K{C1}',ColorScaleRule(start_type='min',start_color='F8CDB9',mid_type='num',mid_value=0,mid_color='FFFFFF',end_type='max',end_color='B7D3F6'))

    # ---- decision matrix rate x salvage
    M0=C1+4
    ws.cell(M0,1,'DECISION MAP: PCC - HMA present worth by discount rate and salvage credit').font=F_TITLE
    ws.cell(M0+1,1,'Salvage multiplier \\ Rate (%)').font=F_H
    mrates=[2,3,4,5,6,7,8]; msalv=[0,0.25,0.5,0.75,1.0]
    for j,rt in enumerate(mrates): c=ws.cell(M0+1,2+j,rt); c.font=F_H; c.fill=FILL_H
    for i,sv in enumerate(msalv):
        r=M0+2+i; c=ws.cell(r,1,sv); c.font=F_H; c.number_format='0%'
        for j,rt in enumerate(mrates):
            ws.cell(r,2+j,f'=({NPW(tP,rate=rt,sm=sv)})-({NPW(tH,rate=rt,sm=sv)})').number_format='$#,##0,"k"'
    M1=M0+1+len(msalv)
    ws.conditional_formatting.add(f'B{M0+2}:H{M1}',ColorScaleRule(start_type='min',start_color='F39C7A',mid_type='num',mid_value=0,mid_color='FFFFFF',end_type='max',end_color='6DA7EC'))
    ws.cell(M1+1,1,'Orange = PCC lower, blue = HMA lower. White cells are within a rounding error of a tie.').font=F_NOTE

    # ---- category helper for stacked chart
    K0=M1+4
    ws.cell(K0,1,'PW by category (chart data)').font=F_H
    hdr(ws,K0+1,1,['Alternative','Initial','Maintenance','Rehabilitation','Lost revenue','Salvage'])
    for i,(nm,t) in enumerate([('HMA',tH),('PCC',tP)]):
        r=K0+2+i; ws.cell(r,1,nm)
        for j,cat in enumerate(['initial','maint','rehab','lost','salvage']):
            ws.cell(r,2+j,f'={L(2+i)}{R+2+j}').number_format=CUR

    # ---- closure timeline helper (x=calendar year, y=lane, size=days)
    Z0=K0+5
    ws.cell(Z0,1,'Closure events (chart data)').font=F_H
    hdr(ws,Z0+1,1,['Activity','Calendar year','Lane','Days','Lost revenue'])
    r=Z0+2
    for lane,(nm,t,alt) in enumerate([('HMA',tH,pr['hma']),('PCC',tP,pr['pcc'])],start=1):
        for k,(n,y,cc,d) in enumerate(alt['events']):
            rr=t['r0']+1+k; c=t['c']
            ws.cell(r,1,f'{nm}: {n}'); ws.cell(r,2,f'=IF({c(1)}{rr}<={PER},$B$7+{c(1)}{rr},NA())')
            ws.cell(r,3,lane); ws.cell(r,4,f'={c(3)}{rr}*{DM}'); ws.cell(r,5,f'={c(5)}{rr}').number_format=CUR
            r+=1
    Z1=r-1

    # ================================================================= charts (anchored in column T onward)
    anchor_col='T'; top=4; step=20
    def place(ch, idx):
        ws.add_chart(ch, f'{anchor_col}{top+idx*step}')

    # 1 PW by category, stacked
    ch=BarChart(); ch.type='col'; ch.grouping='stacked'; ch.overlap=100; ch.gapWidth=60
    data=Reference(ws,min_col=2,max_col=6,min_row=K0+1,max_row=K0+3); cats=Reference(ws,min_col=1,min_row=K0+2,max_row=K0+3)
    ch.add_data(data,titles_from_data=True); ch.set_categories(cats)
    for s,rgb in zip(ch.series,['4A4A4A','8C8C8C','646464',GRAY,'DCDCDC']): color_series(s,rgb)
    style_chart(ch,'Present worth by category (salvage shown below zero)'); place(ch,0)

    # 2 expenditure stream by year (direct + lost revenue, per alt) -> clustered columns of totals
    ch=BarChart(); ch.type='col'; ch.grouping='clustered'; ch.gapWidth=40
    ws.cell(Y0+1,13,'HMA total'); ws.cell(Y0+1,14,'PCC total')
    for k in range(31):
        r=Y0+2+k; ws.cell(r,13,f'=C{r}+D{r}').number_format=CUR; ws.cell(r,14,f'=E{r}+F{r}').number_format=CUR
    data=Reference(ws,min_col=13,max_col=14,min_row=Y0+1,max_row=Y1); cats=Reference(ws,min_col=2,min_row=Y0+2,max_row=Y1)
    ch.add_data(data,titles_from_data=True); ch.set_categories(cats)
    color_series(ch.series[0],HMA_RGB); color_series(ch.series[1],PCC_RGB)
    ch.x_axis.tickLblSkip=5; ch.x_axis.tickMarkSkip=5
    style_chart(ch,'Expenditure stream by calendar year (undiscounted, incl. lost revenue)'); place(ch,1)

    # 3 cumulative discounted cost
    ch=LineChart()
    data=Reference(ws,min_col=7,max_col=8,min_row=Y0+1,max_row=Y1); cats=Reference(ws,min_col=2,min_row=Y0+2,max_row=Y1)
    ch.add_data(data,titles_from_data=True); ch.set_categories(cats)
    color_series(ch.series[0],HMA_RGB,line=True); color_series(ch.series[1],PCC_RGB,line=True)
    ch.x_axis.tickLblSkip=5
    style_chart(ch,'Cumulative discounted cost (crossover = payback year of the higher first cost)'); place(ch,2)

    # 4 discount-rate sensitivity (scatter with lines)
    ch=ScatterChart(); ch.style=13
    xs=Reference(ws,min_col=1,min_row=S0+2,max_row=S1)
    for col,rgb in [(2,HMA_RGB),(3,PCC_RGB)]:
        ys=Reference(ws,min_col=col,min_row=S0+1,max_row=S1); s=Series(ys,xs,title_from_data=True); color_series(s,rgb,line=True); ch.series.append(s)
    ch.x_axis.numFmt='0"%"'; ch.x_axis.scaling.min=2; ch.x_axis.scaling.max=8
    style_chart(ch,'Net present worth vs. discount rate (TDOT 3%, FAA AIP 7%)',xtitle='Discount rate'); place(ch,3)

    # 5 tornado (horizontal bars, low and high cases)
    ch=BarChart(); ch.type='bar'; ch.grouping='clustered'; ch.gapWidth=50
    data=Reference(ws,min_col=4,max_col=5,min_row=T0+1,max_row=T1); cats=Reference(ws,min_col=1,min_row=T0+2,max_row=T1)
    ch.add_data(data,titles_from_data=True); ch.set_categories(cats)
    color_series(ch.series[0],'9E9E9E'); color_series(ch.series[1],'4A4A4A')
    ch.x_axis.tickLblPos='low'; ch.x_axis.scaling.orientation='maxMin'
    style_chart(ch,'PCC minus HMA present worth when one input moves (bars crossing zero flip the winner)'); ch.y_axis.numFmt='$#,##0,,"M"'; place(ch,4)

    # 6 closure timeline bubble
    ch=BubbleChart(); ch.style=18
    for lane,rgb in [(1,HMA_RGB),(2,PCC_RGB)]:
        rows=[r for r in range(Z0+2,Z1+1) if ws.cell(r,3).value==lane]
        xs=Reference(ws,min_col=2,min_row=rows[0],max_row=rows[-1]); ys=Reference(ws,min_col=3,min_row=rows[0],max_row=rows[-1]); zs=Reference(ws,min_col=4,min_row=rows[0],max_row=rows[-1])
        s=Series(values=ys,xvalues=xs,zvalues=zs,title='HMA closures' if lane==1 else 'PCC closures'); s.graphicalProperties.solidFill=rgb; ch.series.append(s)
    ch.y_axis.scaling.min=0; ch.y_axis.scaling.max=3; ch.y_axis.numFmt='0'; ch.x_axis.numFmt='0'
    ch.x_axis.scaling.min=pr['year0']; ch.x_axis.scaling.max=pr['year0']+30
    style_chart(ch,'Runway closure timeline (bubble size = days closed; lane 1 HMA, lane 2 PCC)',xtitle='Calendar year'); ch.y_axis.numFmt='0'; place(ch,5)

    # 7 closure days per year stacked-by-alt? use clustered columns of closure days
    ch=BarChart(); ch.type='col'; ch.grouping='clustered'; ch.gapWidth=40
    data=Reference(ws,min_col=9,max_col=10,min_row=Y0+1,max_row=Y1); cats=Reference(ws,min_col=2,min_row=Y0+2,max_row=Y1)
    ch.add_data(data,titles_from_data=True); ch.set_categories(cats)
    color_series(ch.series[0],HMA_RGB); color_series(ch.series[1],PCC_RGB); ch.x_axis.tickLblSkip=5
    style_chart(ch,'Runway closure days by calendar year'); ch.y_axis.numFmt='0'; place(ch,6)

    # 8 scenario scorecard chart
    ch=BarChart(); ch.type='bar'; ch.grouping='clustered'; ch.gapWidth=40
    data=Reference(ws,min_col=9,max_col=10,min_row=C0+1,max_row=C1); cats=Reference(ws,min_col=1,min_row=C0+2,max_row=C1)
    ch.add_data(data,titles_from_data=True); ch.set_categories(cats)
    color_series(ch.series[0],HMA_RGB); color_series(ch.series[1],PCC_RGB)
    ch.x_axis.scaling.orientation='maxMin'
    style_chart(ch,'Net present worth by scenario',h=12); place(ch,7)

    # ================================================================= live Monte Carlo (columns AK onward)
    N_MC=1000; mc0=37  # column AK
    mcC=lambda k: L(mc0+k)
    ws.cell(4,mc0,'MONTE CARLO (live): uncertain inputs drawn jointly from triangular ranges; press F9 to redraw').font=F_TITLE
    hdr(ws,5,mc0,['Uncertain input','Min','Most likely','Max'])
    dists=[('Discount rate (%)',2,3,7),('HMA cost multiplier',0.8,1,1.25),('PCC cost multiplier',0.8,1,1.25),('Closure-days multiplier',0.5,1,2),('Salvage multiplier',0.5,1,1),('Share of revenue lost',0.3,0.65,1)]
    for i,(lab,a,b,cc) in enumerate(dists):
        r=6+i; ws.cell(r,mc0,lab).font=F_B
        for j,v in enumerate((a,b,cc)):
            cell=ws.cell(r,mc0+1+j,v); cell.font=F_IN; cell.fill=FILL_IN; cell.border=BOX
    d0=14  # first draw row
    hdr(ws,d0-1,mc0,['Draw','Rate %','HMA cost x','PCC cost x','Closure x','Salvage x','Revenue share','NPW HMA','NPW PCC','PCC - HMA','u1','u2','u3','u4','u5','u6'])
    def tri(u,row):
        a,b,cc=f'${mcC(1)}${row}',f'${mcC(2)}${row}',f'${mcC(3)}${row}'
        return f'IFERROR(IF({u}<({b}-{a})/({cc}-{a}),{a}+SQRT({u}*({cc}-{a})*({b}-{a})),{cc}-SQRT((1-{u})*({cc}-{a})*({cc}-{b}))),{b})'
    for k in range(N_MC):
        r=d0+k; ws.cell(r,mc0,k+1)
        for j in range(6): ws.cell(r,mc0+10+j,'=RAND()')
        for j in range(6): ws.cell(r,mc0+1+j,'='+tri(f'{mcC(10+j)}{r}',6+j))
        drow=dict(rate=f'{mcC(1)}{r}',dm=f'{mcC(4)}{r}',sm=f'{mcC(5)}{r}',rf=f'{mcC(6)}{r}')
        ws.cell(r,mc0+7,'='+NPW(tH,cm=f'{mcC(2)}{r}',**drow)); ws.cell(r,mc0+8,'='+NPW(tP,cm=f'{mcC(3)}{r}',**drow))
        ws.cell(r,mc0+9,f'={mcC(8)}{r}-{mcC(7)}{r}')
    d1=d0+N_MC-1
    diff=f'${mcC(9)}${d0}:${mcC(9)}${d1}'
    sc=mc0+17  # summary columns (BB..)
    ws.cell(5,sc,'Simulation summary').font=F_H
    summ=[('Draws',f'=COUNT({diff})','0'),('Probability PCC is lower cost',f'=COUNTIF({diff},"<0")/COUNT({diff})','0%'),
          ('Mean PCC - HMA',f'=AVERAGE({diff})',CUR),('P10',f'=PERCENTILE({diff},0.1)',CUR),('P50 (median)',f'=PERCENTILE({diff},0.5)',CUR),('P90',f'=PERCENTILE({diff},0.9)',CUR)]
    for i,(lab,f,fmt) in enumerate(summ):
        ws.cell(6+i,sc,lab).font=F_B; ws.cell(6+i,sc+1,f).number_format=fmt
    ws.cell(13,sc,'Histogram of PCC - HMA').font=F_H
    hdr(ws,14,sc,['Bin upper edge','Draws','Cumulative'])
    nb=20; lo=f'PERCENTILE({diff},0.01)'; hi=f'PERCENTILE({diff},0.99)'
    for i in range(nb):
        r=15+i
        ws.cell(r,sc,f'={lo}+({hi}-{lo})*{i+1}/{nb}').number_format='$#,##0,"k"'
        prev=f'{L(sc)}{r-1}' if i>0 else None
        ws.cell(r,sc+1,f'=COUNTIF({diff},"<="&{L(sc)}{r})' + (f'-COUNTIF({diff},"<="&{prev})' if prev else ''))
        ws.cell(r,sc+2,f'=COUNTIF({diff},"<="&{L(sc)}{r})/COUNT({diff})').number_format='0%'
    # RESULTS link
    ws.cell(rv+2,1,'Probability PCC is lower (live simulation, F9 redraws)').font=F_H
    ws.cell(rv+2,2,f'={L(sc+1)}7').number_format='0%'; ws.cell(rv+2,2).font=F_H
    ws.cell(rv+2,3,f'="P10 "&TEXT({L(sc+1)}9,"$#,##0,,")&"M   P90 "&TEXT({L(sc+1)}11,"$#,##0,,")&"M   (PCC - HMA)"').font=F_NOTE
    ch=BarChart(); ch.type='col'; ch.grouping='clustered'; ch.gapWidth=10
    data=Reference(ws,min_col=sc+1,max_col=sc+1,min_row=14,max_row=14+nb); cats=Reference(ws,min_col=sc,min_row=15,max_row=14+nb)
    ch.add_data(data,titles_from_data=True); ch.set_categories(cats); color_series(ch.series[0],'4A4A4A'); ch.legend=None
    ch.title='Monte Carlo: PCC minus HMA present worth (left of zero = PCC cheaper)'; ch.width=18; ch.height=9; ch.y_axis.title='Draws'; ch.y_axis.numFmt='0'; ch.x_axis.numFmt='$#,##0,"k"'; ch.x_axis.tickLblSkip=2; ch.x_axis.delete=False; ch.y_axis.delete=False; ch.y_axis.majorGridlines=None
    place(ch,8)
    ws.freeze_panes='A4'
    return dict(rNPW=rNPW, R=R)

# ----------------------------------------------------------------- Monte Carlo (values)
def build_mc(wb):
    ws=wb.create_sheet('MonteCarlo')
    ws.column_dimensions['A'].width=34
    for c in 'BCDEFGHIJ': ws.column_dimensions[c].width=14
    ws['A1']='Monte Carlo simulation of PCC - HMA present worth'; ws['A1'].font=F_TITLE
    ws['A2']=('5,000 draws per project, seed 42, computed in Python with the same engine as the project sheets (values, not formulas). '
              'Inputs varied jointly: discount rate ~ triangular(2, 3, 7)%; HMA and PCC cost multipliers ~ triangular(0.8, 1.0, 1.25) independently; '
              'closure-days multiplier ~ triangular(0.5, 1, 2); salvage multiplier ~ triangular(0.5, 1, 1); revenue share lost ~ uniform(0.3, 1.0).'); ws['A2'].font=F_NOTE
    ws['A2'].alignment=Alignment(wrap_text=True); ws.merge_cells('A2:J3'); ws.row_dimensions[2].height=30
    rng=np.random.default_rng(42); N=5000
    r0=5; out={}
    for key,pr in PROJECTS.items():
        rate=rng.triangular(2,3,7,N); cmh=rng.triangular(0.8,1,1.25,N); cmp_=rng.triangular(0.8,1,1.25,N); dm=rng.triangular(0.5,1,2,N); sm=rng.triangular(0.5,1,1,N); rf=rng.uniform(0.3,1,N)
        H=np.array([npw(pr['hma'],rate[i],30,pr['daily'],cmh[i],dm[i],sm[i],rf[i]) for i in range(N)])
        P=np.array([npw(pr['pcc'],rate[i],30,pr['daily'],cmp_[i],dm[i],sm[i],rf[i]) for i in range(N)])
        D=P-H; out[key]=(H,P,D)
    hdr(ws,r0,1,['Statistic','MBT HMA NPW','MBT PCC NPW','MBT PCC-HMA','SRB HMA NPW','SRB PCC NPW','SRB PCC-HMA'])
    stats=[('Mean',np.mean),('P10',lambda a:np.percentile(a,10)),('P50 (median)',np.median),('P90',lambda a:np.percentile(a,90)),('Std. deviation',np.std)]
    for i,(lab,fn) in enumerate(stats):
        r=r0+1+i; ws.cell(r,1,lab).font=F_B
        for j,key in enumerate(['MBT','SRB']):
            for k in range(3): ws.cell(r,2+3*j+k,float(fn(out[key][k]))).number_format=CUR
    r=r0+1+len(stats)
    ws.cell(r,1,'Probability PCC is lower cost').font=F_H
    for j,key in enumerate(['MBT','SRB']):
        ws.cell(r,4+3*j,float((out[key][2]<0).mean())).number_format='0%'; ws.cell(r,4+3*j).font=F_H
    # histograms of D
    H0=r+3
    ws.cell(H0,1,'Distribution of PCC - HMA present worth (bin upper edge, count of draws)').font=F_H
    hdr(ws,H0+1,1,['Bin upper edge ($)','MBT draws','SRB draws','MBT cumulative %','SRB cumulative %'])
    allD=np.concatenate([out['MBT'][2],out['SRB'][2]]); lo,hi=np.percentile(allD,0.5),np.percentile(allD,99.5)
    edges=np.linspace(lo,hi,31)
    cm_=np.histogram(out['MBT'][2],edges)[0]; cs_=np.histogram(out['SRB'][2],edges)[0]
    for i in range(30):
        rr=H0+2+i; ws.cell(rr,1,float(edges[i+1])).number_format='$#,##0,"k"'
        ws.cell(rr,2,int(cm_[i])); ws.cell(rr,3,int(cs_[i]))
        ws.cell(rr,4,float(cm_[:i+1].sum()/N)).number_format='0%'; ws.cell(rr,5,float(cs_[:i+1].sum()/N)).number_format='0%'
    H1=H0+1+30
    ch=BarChart(); ch.type='col'; ch.grouping='clustered'; ch.gapWidth=10
    data=Reference(ws,min_col=2,max_col=3,min_row=H0+1,max_row=H1); cats=Reference(ws,min_col=1,min_row=H0+2,max_row=H1)
    ch.add_data(data,titles_from_data=True); ch.set_categories(cats); color_series(ch.series[0],'4A4A4A'); color_series(ch.series[1],'B8B6AE')
    ch.title='PCC - HMA present worth: 5,000 joint-uncertainty draws (left of zero = PCC cheaper)'; ch.width=22; ch.height=10; ch.y_axis.title='Draws'; ch.x_axis.numFmt='$#,##0,"k"'; ch.x_axis.tickLblSkip=3; ch.legend.position='b'
    ch.x_axis.delete=False; ch.y_axis.delete=False
    ws.add_chart(ch,'H5')
    ch=LineChart()
    data=Reference(ws,min_col=4,max_col=5,min_row=H0+1,max_row=H1); ch.add_data(data,titles_from_data=True); ch.set_categories(cats)
    color_series(ch.series[0],'4A4A4A',line=True); color_series(ch.series[1],'B8B6AE',line=True)
    ch.title='Cumulative probability that PCC - HMA is below the value'; ch.width=22; ch.height=10; ch.y_axis.numFmt='0%'; ch.x_axis.numFmt='$#,##0,"k"'; ch.x_axis.tickLblSkip=3; ch.legend.position='b'
    ch.x_axis.delete=False; ch.y_axis.delete=False
    ws.add_chart(ch,'H26')
    return out

# ----------------------------------------------------------------- README
def build_readme(wb):
    ws=wb.active; ws.title='README'; ws.column_dimensions['A'].width=120
    lines=[('TDOT Aeronautics LCCA - Decision Workbook (MBT and SRB runs)',F_TITLE),
    ('Purpose: show the two 2026-05-21 analyses side by side with the sensitivity, break-even and closure views the framework workbook does not produce. Companion to TDOA_LCCA_Framework v1.2.0; base costs, closure days and policies are copied from those files.',F_B),
    ('',F_B),('How to use',F_H),
    ('1. Open the MBT or SRB sheet. Yellow cells (B5:B14) are the only inputs. Change the discount rate, period, revenue share lost, closure multiplier, bid multipliers or salvage multiplier and every table and chart updates.',F_B),
    ('2. RESULTS gives present worth by category, EUAC, cost per SY, closure days and runway availability, the winner, the margin, and whether the winner holds at the FAA AIP 7% rate.',F_B),
    ('3. BREAK-EVEN VALUES answer "how wrong would an input have to be" without a goal seek: daily revenue, PCC and HMA bid levels, PCC salvage fraction and discount rate at which the two alternatives tie.',F_B),
    ('4. SENSITIVITY, TORNADO, SCENARIO SCORECARD and the DECISION MAP are pre-run: the scorecard rows carry their own parameters (TDOT policy, FAA AIP 7%/20 yr, no lost revenue, 40% revenue loss, no salvage, bids +20%, closures x2 and x0.5, and a stress case).',F_B),
    ('5. Charts start in column T of each project sheet: PW by category, expenditure stream, cumulative discounted cost, rate sensitivity, tornado, closure timeline, closure days by year, NPW by scenario.',F_B),
    ('6. A live Monte Carlo block (column AK onward on each project sheet) draws 1,000 joint scenarios from triangular ranges you can edit; RESULTS shows the probability that PCC is the lower-cost alternative. Press F9 to redraw; results move by a percent or two between draws.',F_B),
    ('7. To analyse a new project: right-click the TEMPLATE sheet tab > Move or Copy > Create a copy, rename it, then fill the yellow cells: parameters B5:B18, and for each alternative the base direct cost and closure days of every policy activity (columns C:D and L:M of the activity tables). Take them from the framework workbook: General Information D25 (year) and D39 (daily revenue), each Alt sheet column D of the NPW table and cells F4:F10. Every table and chart updates.',F_B),
    ('8. The workbook compares one HMA and one PCC alternative per sheet, which is how the TDOT framework is used. Activity rows follow the TDOT Maintenance Policies (six maintenance events and one rehabilitation for HMA, one maintenance and one rehabilitation for PCC); the years are editable if the policies change.',F_B),
    ('',F_B),('Method notes',F_H),
    ('Present worth = sum over activities within the period of (direct cost x cost multiplier + closure days x closure multiplier x daily revenue x revenue share lost) / (1 + r)^year, plus salvage x cost multiplier x salvage multiplier / (1 + r)^period. Initial construction is year 0. This reproduces the framework workbook NPW exactly at its own settings (3%, 30 years, multipliers = 1).',F_B),
    ('Salvage follows TDOT Maintenance Policies: PCC 25% of initial construction; HMA 12.5% of one mill-and-overlay (2 of 16 years). When the period is shortened the same fraction is applied at the end of the shorter period, which is a conservative simplification.',F_B),
    ('Closure days are the production-rate defaults from the framework (surface treatment 15,000 SY/day; mill and overlay 3,800 SY/day; PCC joint and slab rates). Initial construction closures are not modelled, as in the framework.',F_B),
    ('Lost revenue uses the framework method (sum of the seven RevenueData categories). The "revenue share lost" input lets you test margin-based assumptions.',F_B),
    ('Sources: TDOA_LCCA_Framework_v1.1.2_MBT_20260521_LostRevenue_1.xlsm and _SRB_20260521_LostRevenue.xlsm (ARA); TDOT Maintenance Policies tables 1 and 2; FAA AIP Handbook Appendix U (7% rate, 20-year life for AIP-funded LCCA).',F_B)]
    for i,(t,f) in enumerate(lines):
        c=ws.cell(i+1,1,t); c.font=f; c.alignment=Alignment(wrap_text=True,vertical='top')

def main():
    wb=Workbook(); build_readme(wb)
    info={k:build_project(wb,k,pr) for k,pr in PROJECTS.items()}
    wb.save(OUT); print('wrote',OUT)
    for k,pr in PROJECTS.items():
        if k=='TEMPLATE': continue
        print(k,'engine NPW HMA',round(npw(pr['hma'],3,30,pr['daily']),2),'PCC',round(npw(pr['pcc'],3,30,pr['daily']),2))

if __name__=='__main__': main()
