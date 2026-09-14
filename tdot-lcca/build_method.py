"""'Method' reference sheet for the TDOT LCCA framework: every calculation the workbook performs, written
once, with the formula as it stands in the cell, the rule in plain English and where the rule comes from.
Reference only: nothing on this sheet feeds a calculation. Built with openpyxl and appended to the xlsm by
build_summary.add_plain_sheet."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

F_T = Font(name='Arial', size=12, bold=True); F_H = Font(name='Arial', size=10, bold=True); F_B = Font(name='Arial', size=10)
F_N = Font(name='Arial', size=9, italic=True, color='595959'); F_BTN = Font(name='Arial', size=10, bold=True, color='FFFFFF')
F_F = Font(name='Courier New', size=9)
FILL = PatternFill('solid', fgColor='D9D9D9'); FILL_BTN = PatternFill('solid', fgColor='1D2733'); FILL_SEC = PatternFill('solid', fgColor='EAF2FB')
FILL_FLAG = PatternFill('solid', fgColor='FFF3CD')
thin = Side(style='thin', color='BFBFBF'); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
WRAP = Alignment(wrap_text=True, vertical='top')

GI = "'General Information'"
MP = "'Maintenance Policies'"
HDR = ['Step / quantity', 'Where it lives', 'Formula as it stands', 'What it means', 'Source of the rule']

# (title, rows, note). Each row: [step, where, formula, meaning, source]
SECTIONS = [

 ('1. The answer the workbook produces', [
  ['Net present worth (NPW)', 'Alt sheet E52 (HMA), E43 (PCC)', '=SUM(E36:E51)',
   'Today\'s value of everything the alternative costs over the analysis period: initial construction, every '
   'maintenance and rehabilitation event, the airport revenue lost while the runway is closed, less the salvage credit. '
   'The alternative with the lowest NPW is the economic choice.', 'FHWA LCCA Primer; TDOT framework'],
  ['Where it is reported', 'Summary O4:O7 and D4:D7', '=IFERROR(INDEX(INDIRECT("\'"&$G4&"\'!$E$1:$E$70"),MATCH("Net Present Worth",...)),0)',
   'The Summary reads the NPW out of each alternative worksheet by name, so it follows whatever alternatives exist.',
   'Summary, v1.2.0'],
  ['Everything else is presentation', 'Summary G3:R7, G12:O17, charts 1 to 7', '',
   'The results table, the comparison block, the five cost charts and the two section charts all restate the same '
   'numbers. No chart or block feeds a cost back into the calculation.', 'Summary, v1.2.0'],
 ], 'Read the steps below in order: they follow the money from the quantities you type to the NPW.'),

 ('2. Step 1: initial construction cost', [
  ['Item cost', 'Alt sheet G13:G22', '=IF(B13<>0,F13*E13,0)',
   'The quantity you type in column E times the unit cost the item carries on Pay_Items. Ten item lines are available.',
   'Alternative worksheet'],
  ['Unit cost, item number, unit', 'Alt sheet B, D and F columns',
   '=IF(C13="","N/A",IF(INDEX(UnitCostGrid,MATCH(C13,PayItemKeys,0),$I$11)="",INDEX(UnitCostGrid,MATCH(C13,PayItemKeys,0),1),'
   'INDEX(UnitCostGrid,MATCH(C13,PayItemKeys,0),$I$11)))',
   'Unit costs are never typed on an alternative sheet. Picking a description in column C pulls the item number, the '
   'unit and the price from Pay_Items, so a price change for every future project is one edit on Pay_Items. '
   'UnitCostGrid is Table2[[Unit Cost]:[East]], the four price columns; I11 turns the division picked in C11 into '
   'which of them to read.', 'Pay_Items (Table2), columns F to I'],
  ['Which division the alternative is priced from', 'Alt sheet C11 and I11',
   '=IFERROR(MATCH($C$11,PriceSources,0),1)',
   'C11 is a list: Regular, Middle, West or East. Regular is the statewide Unit Cost column and is what a new '
   'alternative ships with. Any item the chosen division leaves blank is priced from Unit Cost instead, so filling '
   'one regional cost prices that item regionally and changes nothing else. An empty C11 reads as Regular.',
   'v1.2.0, Pay_Items columns F to I'],
  ['Pay-item subtotal', 'Alt sheet G24', '=SUM(G13:G22)',
   'The ten item lines. (In v1.1.2 this summed G14:G22 and silently dropped item 1.)', 'v1.2.0 fix 1'],
  ['Mobilization', 'Alt sheet G25', "=(%s!$D$36/100)*G24" % GI,
   'A percentage of the pay-item subtotal. Default 10 percent, set on General Information D36.', 'General Information D36'],
  ['Engineering', 'Alt sheet G26', "=(%s!$D$37/100)*G24" % GI,
   'A percentage of the pay-item subtotal. Default 5 percent, set on General Information D37. Applied to initial '
   'construction from v1.2.0; before that it was charged only on later work.', 'General Information D37; v1.2.0 fix 2'],
  ['Initial construction total', 'Alt sheet G27', '=SUM(G24:G26)',
   'Subtotal plus mobilization plus engineering. This is the "Initial construction" column on the Summary and the '
   'year-0 row of the life-cycle table.', 'Alternative worksheet'],
 ], None),

 ('3. Step 2: when the later work happens', [
  ['Year of each event', 'Alt sheet C37:C51', "=%s!E10" % MP,
   'Years come from the maintenance policy tables, not from the project. New HMA: maintenance at 4, 8, 12 and 16, '
   'mill-and-overlay rehabilitation at 20, maintenance at 24 and 28, salvage at 30. New PCC: maintenance at 19, '
   'rehabilitation at 27, salvage at 30.', 'Maintenance Policies, Tables 1 and 2'],
  ['The year is an offset', 'Alt sheet column C', '',
   'Column C counts years after construction, not calendar years. Calendar years appear only on the Summary charts, '
   'where the construction year is added back.', 'Alternative worksheet'],
  ['Events beyond the analysis period', 'Alt sheet column E', "=IF(C37>%s!$D$33,0,...)" % GI,
   'An event later than the analysis period is dropped to zero. A 20-year run therefore drops HMA maintenance 5 and 6 '
   'and the PCC rehabilitation. Added in v1.2.0; before that every event was counted whatever the period.',
   'General Information D33; v1.2.0 fix 3'],
 ], 'Changing a policy year on Maintenance Policies moves the event everywhere: the cost, the chart and the closure.'),

 ('4. Step 3: what each maintenance and rehabilitation event costs', [
  ['HMA maintenance 1 and 2 (years 4, 8)', 'Alt sheet J13:P19 and R13:X19 (hidden to the right)', "qty = %s!D10*$C$5 ; cost = unit cost x qty" % MP,
   'Surface treatment over 100 percent of the mainline area, the old markings removed and 100 percent of the markings '
   'repainted. The treatment product is the one chosen in C7.', 'Maintenance Policies rows 10 to 13'],
  ['HMA maintenance 3 and 4 (years 12, 16)', 'Alt sheet Z13:AD20 and AH13:AN21', "qty = %s!D14*($C$5*9)" % MP,
   'Patching 0.75 and 1.00 percent of the area (converted to square feet, hence x9), crack sealing 3.25 and 3.50 '
   'percent of the area as linear feet, then a full surface treatment and markings.', 'Maintenance Policies rows 14 to 21'],
  ['HMA rehabilitation 1 (year 20)', 'Alt sheet AP13:AV22', "qty = %s!D22*$C$5" % MP,
   '4-inch cold mill, tack coat and asphalt overlay over 100 percent of the area, with pre-overlay crack repair on 1.5 '
   'percent and patching on 0.5 percent, plus markings.', 'Maintenance Policies rows 22 to 25'],
  ['PCC maintenance 1 (year 19)', 'Alt sheet J13:P20', "qty = K8*%s!D37 and %s!D38*(C5*9)" % (MP, MP),
   'Joint resealing over the whole estimated joint length, crack sealing 0.25 percent of the area as linear feet, '
   'partial-depth patching 0.13 percent and full-depth patching 0.30 percent.', 'Maintenance Policies rows 37 to 40'],
  ['PCC rehabilitation 1 (year 27)', 'Alt sheet R13:X21', "qty = %s!D41*C5 (slabs) and %s!D45*K8 (joints)" % (MP, MP),
   'Slab replacement on 2 percent of the area, full-depth patching 0.30 percent, partial-depth 0.13 percent, crack '
   'sealing 0.30 percent as linear feet and joint sealing over the full joint length.', 'Maintenance Policies rows 41 to 45'],
  ['Joint length (PCC only)', 'Alt sheet K7 and K8', '=C5/((12.5*12.5)/9) then =K7*25+((K7/8)*12.5)',
   'Slabs are assumed 12.5 x 12.5 ft; the joint length follows from the panel count. Nothing asks the user for a joint '
   'layout.', 'Alternative worksheet (PCC)'],
  ['Mobilization and engineering on each event', 'Alt sheet P17, P18 and the matching cells in each block', "=(%s!$D$36/100)*P16" % GI,
   'Every maintenance and rehabilitation event carries the same mobilization and engineering percentages as initial '
   'construction.', 'General Information D36, D37'],
 ], 'The quantities of later work are computed from the policy rates and the project area. They are not typed anywhere.'),

 ('5. Step 4: how long the runway closes', [
  ['HMA surface treatment (F4, F5, F9)', 'Alt sheet F4:F10 (grey, editable)', '=ROUNDUP((C5/15000),0)+ROUNDUP((C8/20000),0)+ROUNDUP((C8/50000),0)',
   'Production rates: 15,000 S.Y. a day of surface treatment, plus the two marking operations at 20,000 and 50,000 '
   'S.F. a day.', 'v1.2.0 defaults, taken from the MBT and SRB runs'],
  ['HMA with patching and crack sealing (F6, F7, F10)', 'Alt sheet F6, F7, F10', "... +ROUNDUP((C5*%s!D15/10000),0)+ROUNDUP((C5*%s!D14/5000),0)" % (MP, MP),
   'Adds crack sealing at 10,000 L.F. a day and patching at 5,000 S.Y. a day, on the policy quantities for that event.',
   'v1.2.0 defaults'],
  ['HMA mill and overlay (F8)', 'Alt sheet F8', '=ROUNDUP((C5/3800),0)+ROUNDUP((C8/50000),0)',
   'Mill and overlay at 3,800 S.Y. a day, plus markings.', 'v1.2.0 defaults'],
  ['PCC (F4, F5)', 'Alt sheet F4, F5', "=ROUNDUP(C5*550/(50*100/9)/10000,0)+... +7 on F5",
   'Joint work at 10,000 L.F. a day, patching and sealing at 5,000 a day, slab replacement at 1,000 S.Y. a day, plus '
   'seven days of cure before the runway reopens.', 'v1.2.0 defaults'],
  ['These are inputs', 'Alt sheet F4:F10', '',
   'The cells are grey because a project can override them. Type a number and the formula is replaced for that '
   'alternative only.', 'Alternative worksheet'],
 ], 'Closure days matter only when lost revenue is switched on, but they are always reported on the Summary.'),

 ('6. Step 5: lost airport revenue (the "indirect cost")', [
  ['Switch', 'General Information D38', '',
   'Yes adds a lost-revenue row beside every maintenance and rehabilitation event. No leaves the analysis as agency '
   'cost only.', 'General Information D38'],
  ['Daily revenue', 'Alt sheet F2', "=IF(%s!$D$38=\"Yes\",SUMIF(RevenueData!$A:$A,%s!$D$10,RevenueData!$B:$B)+... ,0)" % (GI, GI),
   'Adds the airport\'s revenue lines on the hidden RevenueData sheet (fuel, tenant, other). Only 17 of the 79 airports '
   'have revenue data; text entries such as " NA " count as zero.', 'RevenueData (hidden); v1.2.0 fix 5'],
  ['Warning when there is no data', 'Alt sheet G2', "=IF(AND(%s!$D$38=\"Yes\",F2=0),\"No revenue data for this airport...\",\"\")" % GI,
   'An airport outside the 17 returns $0 lost revenue with a visible warning, instead of the not-available error it produced before.',
   'v1.2.0 fix 5'],
  ['Cost of one closure', 'Alt sheet D38, D40, D42 ... (the "Indirect Cost" rows)', '=F2*F4',
   'Daily revenue times the closure days for that event, then discounted like any other cost.', 'Alternative worksheet'],
 ], 'This is the airport-side counterpart of the road user delay cost in the Caltrans and FHWA RealCost method.'),

 ('7. Step 6: salvage value', [
  ['New HMA', 'Alt sheet D51', "=-%s!D32*AV22" % MP,
   '12.5 percent of the mill-and-overlay cost, credited at the end of the analysis period. The policy sheet describes '
   'it as two years of remaining life on a 16-year overlay.', 'Maintenance Policies row 32'],
  ['New PCC', 'Alt sheet D42', "=-%s!D46*G27" % MP,
   '25 percent of the whole initial construction cost, credited at the end of the analysis period. The policy sheet '
   'describes it as ten years of remaining life.', 'Maintenance Policies row 46'],
  ['When it is credited', 'Alt sheet C51 (HMA), C42 (PCC)', "=%s!D33" % GI,
   'Always at the last year of the analysis period, so it is discounted the hardest of any line.',
   'General Information D33'],
 ], 'The two rules are not symmetric: concrete recovers a share of everything spent up front, asphalt a share of one '
    'rehabilitation. See section 11.'),

 ('8. Step 7: discounting and the total', [
  ['Present worth of one event', 'Alt sheet column E', "=IF(C37>%s!$D$33,0,D37/(1+(%s!$D$34/100))^C37)" % (GI, GI),
   'Cost divided by (1 + rate) raised to the number of years after construction. Constant-dollar costs with a real '
   'discount rate: no inflation is added anywhere in the workbook.', 'General Information D33, D34'],
  ['Discount rate', 'General Information D34', '',
   'TDOT uses 3 percent. FAA PGL 22-01 (June 2022) replaced the fixed 7 percent with the OMB Circular A-94 Appendix C '
   'real rate, 2.0 percent for 2026. Chart 2 on the Summary shows the answer from 2 to 8 percent.',
   'TDOT policy; FAA PGL 22-01; OMB A-94'],
  ['Analysis period', 'General Information D33', '',
   'TDOT uses 30 years. FAA leaves the period to the engineer. Events beyond it are dropped, salvage moves with it.',
   'TDOT policy'],
  ['Net present worth', 'Alt sheet E52 (HMA), E43 (PCC)', '=SUM(E36:E51)',
   'The sum of the discounted column, initial construction included.', 'Alternative worksheet'],
 ], None),

 ('9. What the Summary adds on top', [
  ['Results table', 'Summary G3:R7', '=IFERROR(SUMIFS(INDIRECT("\'"&$G4&"\'!$E$37:$E$52"),...),0)',
   'Splits the same NPW into initial construction, maintenance, rehabilitation, lost revenue and salvage by matching '
   'the activity labels, so the five columns always add back to the NPW. Also difference to the lowest, closure days '
   'inside the period and runway availability.', 'Summary, v1.2.0'],
  ['Which alternatives exist', 'Summary G4:I7', '=INDEX(Database!$D:$D,4)',
   'Names and worksheet names come from the hidden Database sheet that the Alternative Setup form writes. INDEX on '
   'whole columns survives the row deletion the form performs when an alternative is removed.', 'Database (hidden)'],
  ['Verdict', 'Summary G9, G10', '',
   'Names the lowest-NPW alternative, the margin to the next, the rate and period, and whether the same alternative '
   'still wins at 2 and at 7 percent.', 'Summary, v1.2.0'],
  ['Comparison block (RealCost layout)', 'Summary G13:O17', "=H14*((%s!$D$34/100)*(1+%s!$D$34/100)^%s!$D$33)/((1+%s!$D$34/100)^%s!$D$33-1)" % (GI, GI, GI, GI, GI),
   'Agency cost (initial + maintenance + rehabilitation + salvage) and user cost (lost revenue) as present worth and '
   'as equivalent uniform annual cost. EUAC = PW x capital recovery factor.', 'Caltrans / FHWA RealCost layout'],
  ['Sensitivity to the discount rate', 'Summary W10:AA36, chart 2', '=SUMPRODUCT((year<=period)*cost/(1+r)^year)',
   'Recomputes every alternative at 25 rates from 2 to 8 percent without touching the alternative sheets.',
   'Summary, v1.2.0'],
  ['By calendar year', 'Summary W39:AJ71, charts 3, 4 and 5', '=SUMIF(year applied, ...)',
   'Spend by year undiscounted, cumulative discounted cost (the payback crossing) and closure days by year.',
   'Summary, v1.2.0'],
  ['Pavement section read-back', 'Summary G80:Q86, charts 6 and 7', '=36*C.Y./mainline S.Y. ; =2666.6667*tons/(pcf*mainline S.Y.)',
   'Derives each layer thickness from the quantities already entered, totals the section, cross-checks it against the '
   'excavation quantity and writes the section out as a string for the alternative description. Display only.',
   'Summary, v1.2.0'],
 ], None),

 ('10. Current settings in this file (live)', [
  ['Analysis period, years', 'General Information D33', "=%s!$D$33" % GI, 'The period every alternative is measured over.', ''],
  ['Discount rate, percent', 'General Information D34', "=%s!$D$34" % GI, 'The real rate used for every present worth above.', ''],
  ['Capital recovery factor', 'computed here', "=IF(%s!$D$34=0,0,(%s!$D$34/100)*(1+%s!$D$34/100)^%s!$D$33/((1+%s!$D$34/100)^%s!$D$33-1))" % ((GI,)*6),
   'Multiply any present worth by this to get the equivalent uniform annual cost.', ''],
  ['Present worth of $1 spent at year 20', 'computed here', "=1/(1+%s!$D$34/100)^20" % GI,
   'What a dollar of rehabilitation at year 20 is worth today at the current rate.', ''],
  ['Lost revenue counted', 'General Information D38', "=%s!$D$38" % GI, 'Yes adds the closure cost rows.', ''],
  ['Alternatives registered', 'Database A4:A7', '=COUNTA(Database!$A$4:$A$7)', 'How many alternatives the form has created.', ''],
 ], 'These cells read the workbook as it stands right now, so they change with the inputs.'),

 ('11. Assumptions and policy choices worth knowing', [
  ['Engineering on initial construction', 'Alt sheet G26', '',
   'v1.2.0 charges engineering on initial construction as well as on later work. It raises every initial cost by the '
   'engineering percentage against v1.1.2 results, and it affects PCC more than HMA because PCC salvage is a share of '
   'initial cost.', 'v1.2.0 fix 2'],
  ['Salvage is not symmetric', 'Maintenance Policies D32 and D46', '',
   'Concrete recovers 25 percent of total initial cost, asphalt 12.5 percent of one overlay. Concrete therefore '
   'carries a large year-30 credit that moves with the discount rate in step with its costs, which is why the concrete '
   'curves in chart 2 are nearly flat while the asphalt curve falls. In some projects this line alone decides the '
   'answer. The method is the remaining-life rule in AAPTP 06-06 and FAA guidance; the two fractions are this '
   'workbook\'s own. PCC checks out: 10 of 40 years left at year 30 is 25 percent. HMA does not. Table 1 places the '
   'mill and overlay at year 20, so a 16-year overlay life leaves 6 of 16 years at year 30, which is 37.5 percent, '
   'not 12.5. Two of 16 is what an overlay placed at year 16 would leave, the rehab year in Table 3.',
   'Pre-existing policy; flagged for decision'],
  ['Lost revenue is gross revenue', 'RevenueData; Alt sheet F2', '',
   'The whole of the airport\'s fuel, tenant and other revenue is treated as lost for every closure day. A percentage '
   'lost by category would be more defensible.', 'Pre-existing; flagged for decision'],
  ['PCC joint length is estimated two ways', 'Alt sheet K8 versus F4 and F5', '',
   'The cost side assumes 12.5 x 12.5 ft slabs; the closure-day formula assumes 550 L.F. of joint per 50 x 100 ft '
   'panel. The closure formula therefore produces about a third less joint length than the cost formula, which '
   'understates PCC closure days and the lost revenue that follows from them.', 'Pre-existing; worth a decision'],
  ['Closure production rates are defaults', 'Alt sheet F4:F10', '',
   'They reproduce the durations used in the completed Murfreesboro and Upper Cumberland runs. A project with unusual '
   'phasing should override them.', 'v1.2.0 defaults'],
  ['Asphalt unit weight for the section read-back', 'Summary J81', '',
   '145 pcf, the value that reproduces the Murfreesboro section exactly. It affects the displayed thickness only, '
   'never a cost.', 'Summary, v1.2.0'],
  ['The division columns are wired but empty', 'Pay_Items columns G, H, I', '',
   'Middle, West and East are read from v1.2.0: each alternative picks one in C11 and prices from it. They ship '
   'empty, and a blank cell falls back to the statewide Unit Cost, so until Aeronautics fills them every division '
   'still prices the same. The locator map on the Summary colors the airports by division, and General Information '
   'D13 names the one the selected airport sits in.', 'v1.2.0; the numbers are still a decision'],
  ['Unit costs date from the 2022 framework', 'Pay_Items column F', '',
   'The Typical Values sheet compares them with recent bid prices. The asphalt defaults look low against 2025 '
   'southeastern bids.', 'Typical Values, section 3'],
  ['No inflation, no risk analysis', 'throughout', '',
   'Costs are constant dollars discounted at a real rate, and every input is a single value. There is no probabilistic '
   'treatment of cost or timing.', 'FHWA LCCA Primer, deterministic method'],
 ], None),
]


def build(path):
    wb = Workbook(); ws = wb.active; ws.title = 'Method'
    for col, w in zip('ABCDE', [34, 30, 56, 66, 30]): ws.column_dimensions[col].width = w
    c = ws['A1']; c.value = '=HYPERLINK("#\'General Information\'!D9","◄ General Information")'; c.font = F_BTN; c.fill = FILL_BTN; c.alignment = Alignment(horizontal='center', vertical='center'); c.border = BOX
    c = ws['B1']; c.value = '=HYPERLINK("#Summary!G1","Summary")'; c.font = F_BTN; c.fill = PatternFill('solid', fgColor='2A78D6'); c.alignment = Alignment(horizontal='center', vertical='center'); c.border = BOX
    ws.row_dimensions[1].height = 22
    ws['A2'] = 'Reference sheet: nothing here feeds the calculation.'; ws['A2'].font = F_N
    ws['A3'] = 'METHOD: WHAT THIS WORKBOOK CALCULATES, AND WHY'; ws['A3'].font = F_T
    ws['A4'] = ('Every calculation in the framework, written once: the formula as it stands in the cell, the rule in plain English, and where the rule comes from. '
                'Sections 2 to 8 follow one alternative from the quantities you type to its net present worth; section 9 covers what the Summary adds; section 10 reads the current '
                'settings live; section 11 lists the assumptions a reviewer will ask about. To see a formula in the workbook itself, go to the cell named in column B, or press '
                'Ctrl + ` on any sheet to show all of its formulas at once.')
    ws['A4'].font = F_N; ws['A4'].alignment = WRAP
    ws.merge_cells('A4:E4'); ws.row_dimensions[4].height = 42
    ws.page_setup.orientation = 'landscape'
    r = 6
    for title, rows, note in SECTIONS:
        ws.cell(r, 1, title).font = F_H; ws.cell(r, 1).fill = FILL_SEC
        for cc in range(2, 6): ws.cell(r, cc).fill = FILL_SEC
        r += 1
        for i, h in enumerate(HDR):
            c = ws.cell(r, 1 + i, h); c.font = F_H; c.fill = FILL; c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical='center')
        r += 1
        flag = title.startswith('11.')
        live = title.startswith('10.')
        fmts = ['0', '0.00', '0.0000', '0.0000', 'General', '0']
        for k, row in enumerate(rows):
            for i, v in enumerate(row):
                c = ws.cell(r, 1 + i, v)
                if i == 2 and isinstance(v, str) and v.startswith('=') and not live:
                    c.data_type = 's'        # show the formula as text, do not evaluate it
                c.font = F_F if i == 2 else F_B; c.border = BOX; c.alignment = WRAP
                if flag and i == 0: c.fill = FILL_FLAG
            if live: ws.cell(r, 3).number_format = fmts[k] if k < len(fmts) else 'General'
            r += 1
        if note:
            c = ws.cell(r, 1, note); c.font = F_N; c.alignment = WRAP
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5); ws.row_dimensions[r].height = 26; r += 1
        r += 1
    ws.freeze_panes = 'A5'
    wb.save(path)
