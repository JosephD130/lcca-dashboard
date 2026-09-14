# TDOA LCCA Framework v1.2.0 — change notes

Base: `TDOA_LCCA_Framework_v1.1.2_ARA_Task2_08182026.xlsm`
Output: `TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm` (self-contained, nothing to import)

All workbook edits were made directly in the sheet XML so the VBA project, the 192 ActiveX
pay-item comboboxes, the tables and data validations are untouched. The new Summary sheet is
formulas and charts only; no VBA was added or changed.

## 1. Calculation fixes (these change answers)

| # | Where | Was | Now |
|---|---|---|---|
| 1 | TMP(NewHMA), TMP(NewPCC), TMP(HMARehab) G24 | `SUM(G14:G23)` skipped pay item 1 | `SUM(G13:G22)` |
| 2 | All templates, row 26 | blank | Engineering % applied to initial construction (`G26 = D37% x G24`); was only applied to M&R activities |
| 3 | All templates, discounted-cost column E | discounted every activity regardless of analysis period | `IF(year > Analysis Period, 0, ...)` so a 20-year run drops years 24 and 28 |
| 4 | General Information D33 | 10 (Overview says 30) | 30 |
| 5 | Indirect templates F2 | `SUM(XLOOKUP(...))`, #N/A for any airport outside the 17 | SUMIF chain (works in any Excel, ignores text cells); returns 0 with a visible warning in G2 when the airport has no revenue data |
| 6 | General Information D39 | XLOOKUP | same SUMIF chain |
| 7 | TMP(NewPCC) and TMP(NewPCC)_IndirectCost B14:B22 | `VLOOKUP(C14,...)` with no blank guard, #N/A on empty rows | `IF(C14="",0,VLOOKUP(...))` like row 13 |
| 8 | Pay_Items D20, D28, D30, D32 | four items all described "Separation Geotextile", so VLOOKUP by description always returned the first | suffixed with the spec: (P-154), (P-208), (P-209), (P-219) |
| 9 | Indirect templates F4:F10 | closure days = 0, users re-typed production-rate formulas in each file | pre-filled with the production-rate defaults used in the MBT and SRB runs (15,000 SY/day surface treatment, 3,800 SY/day mill & overlay, PCC joint/slab rates); still gray input cells, override as needed |
| 10 | General Information D9 dropdown | `$L$9:$L$84`, last four airports unreachable | `$L$10:$L$88` |
| 11 | TMP(HMARehab) C42, D44 | C42 added the construction year to the policy year (discounted to ~0); D44 pointed at an empty cell | fixed, in case the rehab option is re-enabled |
| 12 | Maintenance Policies D32 (New HMA salvage) | fixed `0.125`, described as "2 years of mill and overlay" against a table that places that overlay at year 20 | `=IF(E22>D33,0,MAX(0,MIN(1,(E22+F32-D33)/F32)))`, remaining life over a 16-year overlay life. 37.5% at the default 30-year period |
| 13 | Maintenance Policies D46 (New PCC salvage) | fixed `0.25`, correct at exactly a 30-year period and nowhere else | `=MAX(0,MIN(1,(0+F46-D33)/F46))`, remaining life over a 40-year life. Still 25% at 30 years |

### Salvage, fixes 12 and 13 in detail

The workbook says what rule it uses on each salvage row: remaining life over expected life, times the
cost of the asset being salvaged. Both terms of "remaining" were already live. The credit is taken at
the analysis period (`General Information` D33, an input) and the HMA overlay happens in the year
`Maintenance Policies` E22 names (a policy input). Only the fraction was frozen, so the sheet could be
right at one combination of the two and was not right even at the default: it carried two of sixteen
years left on an overlay its own table places at year 20, which at thirty years has six left. The
concrete figure had the same shape and happened to be correct at exactly thirty years.

Both are now computed. The expected life each one divides by has moved out of the formula into a new
**Asset life (yrs)** column beside the salvage row, 16 for the overlay and 40 for concrete, where it can
be read and changed. The sentence in column C is built from the numbers, so it restates itself instead
of going stale: at thirty years it reads "Salvage value: 6 of 16 years left on the mill and overlay
placed in year 20 (37.5% of its cost) at the 30-year analysis period". The Year Applied shown on all
four salvage rows now reads the analysis period rather than a fixed 30, which is where the credit was
always actually taken.

| Analysis period | HMA credit | PCC credit |
|---|---|---|
| 20 years (the overlay is laid in the salvage year) | 100% | 50% |
| 25 years | 68.8% | 37.5% |
| 30 years (default) | 37.5% | 25% |
| 36 years (the overlay is fully consumed) | 0% | 10% |

A period that ends before the overlay is ever laid credits nothing for HMA, because there is no overlay
to salvage. Tables 3 and 4 keep their zero: "need for reconstruction" is a stated policy, not a
remaining-life calculation, and it is left alone.

**This moves published results.** At Murfreesboro, 3% over 30 years, the HMA alternative goes from
$8,809,266.42 to $8,572,110.60 and concrete is unchanged at $8,028,734.49, so the margin narrows from
$780,532 to $543,376. Concrete is still the lower of the two. What remains a policy choice, and is
recorded as one on the Method sheet, is the asymmetry of the *basis*: concrete salvages a share of its
whole initial construction while asphalt salvages a share of one overlay. That was not changed here.

## 2. Chart data (alternative worksheets)

Columns K:M (K:N on indirect templates) were hard-positioned: Maintenance 1 was placed in row
37+4 because the policy says year 4, so changing a policy year silently misplaced the bar and the
axis read 1 to 31. They are now year-indexed with SUMIF/SUMIFS keyed on `Year Applied`, respect the
analysis period, and the chart has calendar years on the category axis. The indirect-cost chart is
now a stacked column of Direct + Lost revenue (undiscounted); the mixed discounted series was
removed. Column N (total discounted by year) is kept for the Summary.

## 3. Summary worksheet (formulas, no macro)

Columns A:E are still written by the Alternative Setup form exactly as in v1.1.2, and the original
"Alternatives Comparison" chart is kept (moved to the bottom right). Everything from column G
onward is ordinary worksheet formulas that read the hidden Database sheet (alternative names,
types and worksheet names) and reach into each alternative worksheet with INDIRECT, so any project
populates it with no further steps:

- Results table G3:R7: worksheet, alternative, type, initial construction, maintenance PW,
  rehabilitation PW, lost revenue PW, salvage PW, net present worth, difference to the lowest,
  closure days in the analysis period, runway availability.
- Verdict line G9: lowest-cost alternative and its margin to the next (or "only one alternative so
  far" / "tied with the next alternative"); G10 flags whether the same alternative is lowest at the
  FAA AIP rate of 7%.
- The names come from the hidden Database sheet through INDEX on whole columns, because the VBA
  deletes a Database row when an alternative is removed; a fixed cell reference would become #REF!.
- Closure days count only activities inside the analysis period (from the by-year block), so a
  20-year run drops the late maintenance closures; with lost revenue off it falls back to F4:F10.
- Five charts: 1 present worth by category (stacked, salvage below zero); 2 NPW versus discount
  rate 2% to 8% in 0.25 steps; 3 expenditure stream by calendar year (undiscounted); 4 cumulative
  discounted cost by year; 5 runway closure days by calendar year.
- Comparison block G12:P17 in the Caltrans/FHWA RealCost layout: agency cost (initial + M&R + salvage) and
  user cost (lost airport revenue) as present worth and equivalent uniform annual cost, total, difference to
  the lowest in dollars and percent. The decision workbook carries the same block beside RESULTS.
- G10 now checks the winner at 2% (OMB A-94 real rate, which FAA PGL 22-01 of June 2022 substituted for the
  fixed 7% rule) and at 7% (the pre-2022 AIP rule).
- Pavement section read-back G80:P86, with charts 6 and 7: the thickness of each layer, derived from the
  pay-item quantities already entered against the mainline area, the total section, a cross-check against
  the excavation quantity, and the section written out as a string to paste into the alternative
  description. Asphalt is the only layer needing an assumption, and its unit weight sits in J81 (145 pcf,
  the value that reproduces the Murfreesboro section exactly). Chart 7 fills in only when a shoulder area
  is entered and shows the same quantities spread over mainline plus shoulder.
- Chart data lives in columns W onward, greyed and labelled "calculated automatically; do not edit".
- Navigation: dark HYPERLINK button cells "General Information" and "Instructions" at the top of
  Summary, and a "View Summary" button under Alternative Setup on General Information (row 44).
  They need no macro, so they work when ActiveX is blocked.
- Print area A1:U70, landscape, one page wide. Empty alternative rows show blank; with no
  alternatives the verdict line reads "No alternatives yet".

## 3b. Typical Values sheet and input hints

A new last sheet, "Typical Values", is reference only (nothing feeds the calculation): what TDOT Aeronautics
manages (78 public-use airports, 6 commercial / 72 GA, 69 NPIAS, ~70 in the APTech pavement network), typical
TN GA runway geometry with a seven-airport sample and areas, the Pay_Items defaults shown live beside 2025
southeastern bid prices, the closure production rates built into F4:F10, economic parameters (TDOT 3%/30 yr;
FAA PGL 22-01 pointing to OMB A-94 real rates, 2.0% in 2026; the legacy 7%/20 yr; Caltrans 4%), maintenance
timing against published service lives, and a live daily-revenue table for the 17 airports. Every row has its
source URL; see verification/RESEARCH_SOURCES.md for what could and could not be confirmed. General Information
gets a "Typical Values" button (D46) and short hints in column F beside D25:D38.

## 3c. Look and first-run usability

Nothing on this list changes a number. The layout, fonts and grey input cells are unchanged, so the
workbook still reads as the same tool.

- General Information: a "How to use this workbook" card beside the TDOT logo with the three steps,
  and a live status line under it that lists whatever required input is still empty and turns green
  when they are all filled. The three section labels sit on a light band. The salvage note in D35 is
  styled as a note rather than an input. The sheet now prints as one landscape page (A1:J48).
- Every alternative worksheet gets a navigation bar in row 1: "General Information" and "Summary"
  buttons and a one-line reminder of which cells are inputs. It is on the hidden templates, so each
  new alternative the form creates carries it.
- Summary: the lowest-cost row is highlighted in the results table and in the comparison block by
  conditional formatting, so the answer is visible without reading the numbers.
- Tab colors group the sheets: navy for General Information, blue for Summary, light blue for the
  alternative worksheets, grey for reference sheets.

## 3d. Summary as a dashboard, and the setup flow

The Summary now opens with the answer rather than with a table, and every sheet says where it sits
in the sequence a user actually follows.

- Six KPI tiles across rows 3 to 9: lowest present worth (with the section it buys), margin to the
  next alternative, equivalent annual cost, initial construction, unit cost per square yard, and
  rate sensitivity. The rate tile tests the lowest-cost alternative at all 25 rates in the
  sensitivity block, 2% to 8%, not just at the two ends, and it is green only while it actually
  says the winner holds. Every tile is a formula over cells that already existed; nothing new is
  calculated. The margin tile turns amber when the two best alternatives are within five percent of
  each other, which is the point at which a reviewer should treat them as tied.
- Data bars inside the net-present-worth column, so the results table reads without going to a chart.
- Chart 8, new: each alternative's initial construction over the mainline area, against the $210 to
  $280 per square yard all-in range for recent Tennessee runway work cited on Typical Values. The
  band is drawn from two cells, so it can be updated without touching the chart. This workbook prices
  the pavement contract and the published range covers pavement, lighting and grading, so every bar
  should sit below the band; far below it, or above it, is worth a second look at the quantities.
- Chart 1's five categories use one light-to-dark ramp instead of the mixed greys, so the stack reads
  in the order the categories are listed.
- The chart the Alternative Setup form maintains is parked below the dashboard with a line saying what
  it is. It duplicates chart 1 with less detail, but the macro still updates it, so it is moved rather
  than removed.

The project identity line moved from L1 to G2, a row of its own under the band, and the note about
the hidden columns dropped to row 10 just above the results table.

**The Summary row map moved.** Anything outside the workbook that reads Summary cells by address has
to move with it: the results table is now rows 12 to 15 (was 4 to 7), the verdict lines are G17 and
G18 (were G9 and G10), the comparison block starts at row 22 (was 14), the pavement-section block at
row 90 (was 80) and its asphalt unit weight at J91 (was J81), and the map-data block at row 130 (was
115). The KML macro, the verification scripts and the worked examples in this delivery were all
updated; a private copy of the old macro would need the same edit.

### The setup flow

- The "How to use this workbook" card on General Information now lists five steps and names the sheet
  for each: Overview and Instructions, General Information, Pay_Items, Alternative Setup, Summary.
- Every sheet in that sequence says which step it is, in the same place: Overview and Instructions
  carry "STEP 1 of 5", Pay_Items "STEP 3 of 5", each alternative worksheet "STEP 4 of 5" and the
  Summary "STEP 5 of 5".
- Overview, Instructions, Pay_Items and Maintenance Policies gained the navigation row the alternative
  worksheets and the reference sheets already had.
- General Information gained two more buttons, Pay_Items and Maintenance Policies, so every sheet in
  the flow is one click away. It still prints as one landscape page.

### Pay_Items

- Row 1 is now a navigation band with the step line. The header row stays frozen and repeats on every
  printed page.
- The Unit Cost column is marked as the input it is: grey fill, currency format, a box. Nothing about
  the values changed, and Table2 (C2:I59) with all 56 pay items is untouched.
- Middle, West and East are marked as fillable too, in a lighter grey, now that they are read.
- The part headings in column A read as bands down the left of the table.
- Four notes under the table say what the sheet drives, how the division columns are used, that 29 of
  the 56 pay items carry no unit cost at all and will price at $0 if used, and where the sources are.

### Maintenance Policies

- Navigation band, a title, and four lines beside the logo: that the sheet is **live**, which table
  drives which alternative type, what the Rate column means, and where closure days actually come from.
- Correction to the v1.2.0 line first written here. It said the sheet was reference only, that the
  schedules live in the hidden templates and that editing a number here changes nothing. That is the
  opposite of the truth and it is now fixed. Every Rate cell in column D and every Year Applied cell in
  column E is read by the alternative worksheets: D10 sets a maintenance quantity, E22 sets the year
  the HMA overlay happens, D32 sets the HMA salvage credit, D46 the PCC one. A user who believed the
  old note could have changed a policy number expecting nothing to happen and silently moved every
  result. The verifier now checks both that the note says "live" and that the templates really do read
  both columns, so the claim cannot drift from the file again.
- The four table headers sit on a navy band and stay visible as the sheet scrolls.


## 3e. The TDOT mark

The logo was on four of the nine sheets a user opens, each anchored and sized a little differently,
and all four were drawn 8 to 11 percent taller than the artwork: the file is 723 by 316, an aspect of
2.288, and the placements ran 2.06 to 2.12, so the lettering was stretched and the TN square was not
square.

- Row 1 is now a band on every sheet a user can reach, thirteen of them, 40 points tall. It already
  existed everywhere and already carried the navigation buttons, so the band costs only its extra
  height. Forty points is the floor at which "Department of Transportation" stays legible; the mark is
  1.27 in wide there.
- The mark sits at the right-hand edge of each sheet's content, at the artwork's own aspect, where it
  cannot collide with the buttons or the step line.
- Row 1 repeats at the top of every printed page (Print_Titles), so the mark prints on every sheet of
  paper rather than only the first. Pay_Items repeats rows 1 and 2, the band and its table header.
- Three sheets had no drawing part at all and now have one: Pay_Items, Typical Values and Method.
- Two knock-on print fixes. Overview and Instructions now scale to one page wide; their text box is
  wider than a portrait page and had been clipped on the right since v1.1.2. Each alternative worksheet
  now prints its cost and present-worth tables (A1:G56) rather than spreading its chart-data columns
  across half a dozen pages, which took the worked example from 41 printed pages to 30.
- The step lines moved out of the band to make room: each sheet carries a short "STEP n of 5" chip
  beside its buttons and the sentence sits on the row below. On Overview and Instructions the
  Aeronautics address block moved to the left margin under the band, where the logo used to sit.


## 3f. The pay item picker and the division an alternative is priced from

### The picker reads now

Each pay item picker is a Forms 2.0 combo box listing two columns, `Pay_Items!C3:D59`: the pay item
number and its description. A Forms combo box sizes its drop-down list to the control unless
ListWidth says otherwise, and with no explicit column widths it splits that evenly, so each column
got about 1.45 in. Ten characters of pay item number fit in that; a 76-character description did not,
which is why most of the list read as truncated fragments.

- Every one of the 105 pickers now carries a ListWidth of 9 in, so an even split gives each column
  4.5 in, enough for the longest description in the catalogue. The number stays in the list.
- The picker column (C) on every alternative template goes from 41.7 to 53.7 characters wide, and each
  combo box is re-anchored to end exactly at that column's edge, so what you picked reads in the cell
  as well as in the list. Columns A and B give up the width, so the table prints the same.
- This is a change inside the persisted ActiveX streams ([MS-OFORMS] MorphDataControl): the property
  mask gains fListWidth, the DataBlock gains the four bytes that property needs, and cbMorphData grows
  to match. All 105 were re-parsed afterwards with an independent implementation of the format and
  each one consumes exactly to the end of its stream.

### Which division the alternative is priced from

Pay_Items has carried Middle, West and East columns beside Unit Cost since v1.1.2 and nothing read
them. They are read from v1.2.0.

- Every alternative worksheet has a **Price from** list in C11, one row above the pay item table:
  Regular, Middle, West or East. A new alternative ships set to Regular, which is the statewide Unit
  Cost column, so nothing changes until someone picks otherwise.
- All 156 unit cost lookups on the five templates now read the chosen column. Any item the chosen
  division leaves blank is priced from Unit Cost instead, so filling one regional cost prices that one
  item regionally and changes nothing else. An empty C11 reads as Regular rather than producing an error.
- The note beside the list names the division the selected airport sits in, read from General
  Information D13, so the person choosing can see which column matches the project.
- Three workbook names carry the lookup: `PriceSources` (the four labels), `UnitCostGrid`
  (`Table2[[Unit Cost]:[East]]`) and `PayItemKeys` (`Table2[Pay Item Description]`). I11 on each sheet
  turns the choice into a column number; column I is a hidden spacer on every template.
- Every result in this release is unchanged as a consequence: both worked examples and the Murfreesboro
  regression recompute to the same numbers they did before, because the regional columns ship empty and
  Regular is the same column the old formula read.

### Two things fixed alongside

- TMP(HMARehab) was the one template whose working columns (I to BL) were never hidden, so an HMA
  overlay alternative opened showing its internal cost blocks and chart data. They are hidden now, and
  that sheet's chart is set to plot hidden cells so it still draws.
- General Information's **Airport Owner** row read `Table17` column 6, which is empty for all 79
  airports, so it always showed blank. It now reads column 4 and is labelled **County**, which is
  populated for every airport. State Region below it has not moved, so the Summary still finds it.


## 4. Housekeeping

- Instructions text box: ActiveX "blocked content" steps added (Trust Center > ActiveX Settings,
  then restart Excel; contact IT if greyed out); step 1 now says D9:D39; step 2 names the
  "Alternative Setup" button; step 4 rewritten for the hidden RevenueData sheet and the
  17-airport rule, including how to unhide a sheet; step 5 notes the pre-filled closure days;
  step 6 describes the new Summary and the "View Summary" button. The text box was extended
  (rows 9 to 76) so the longer text is not clipped.
- Four misspellings in the issued text corrected, wording otherwise untouched: "Federal Aviation
  Adminimstration" and "pavement clossures" and "associeted with limited facility uses" in the Overview
  text box, and the General Information label "Intial Construction Year". Every visible string in the
  workbook, the KML macro, these notes and the two reports was run through a spell check against a
  domain word list; the run is clean.
- Locator map: the boundary source carries the Tennessee/North Carolina line twice, two tracings within
  0.01 degrees of each other, which drew that stretch doubled. The second copy is dropped; the outline is
  now 12 segments and 179 points.
- Cell notes on D26 and D39 no longer carry a person's name.
- Two broken external links to files on a C: drive (v1.1.004 and a Savannah copy) removed.
- Workbook set to fully recalculate on open.

## 5. Not changed (needs a decision)

- Overview text: Neel-Schaffer's rewrite (2022 APTech / 2026 NS+ARA history, the four airport
  criteria) is in their review doc; paste it once Mat approves the wording.
- Salvage basis, not the fraction. The fractions are fixed (section 1, items 12 and 13) and are now
  computed from remaining life. What is left is the choice of *basis*: concrete salvages a share of its
  whole initial construction, asphalt a share of one mill and overlay. That asymmetry is a real
  modelling decision, it is what makes the concrete curves in chart 2 nearly flat while the asphalt
  curve falls, and Aeronautics should either confirm it or move both to a common basis. The decision
  workbook's salvage multiplier and the Summary sensitivity block make the effect visible.
- Lost revenue counts gross fuel sales and tenant rent as lost during a runway closure. A per-category
  "% lost during closure" factor on RevenueData would be more defensible.
- FAA AIP discount rate: PGL 22-01 (June 2022) replaced the fixed 7% with OMB A-94 real rates (2.0% for
  2026) and dropped the fixed 20-year period. The Summary checks the winner at both 2% and 7%; Aeronautics
  should decide which rule to cite for AIP-funded projects.
- RevenueData hygiene: MBT and MQY rows are identical; XNX and M54 hold text where numbers belong.

## 6. First open in Excel

1. Open `TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm`, enable content. Nothing to install.
2. General Information: pick an airport, set D38 = Yes, D26 = 52777.7, D28 = 5433, D25 = 2027.
   Alternative Setup > add one New HMA and one New PCC > Close. Click "View Summary": the table,
   verdict line and five charts should populate with no #N/A.
3. On an Alt sheet check that columns L:N sum to the NPW table (sum of N = Net Present Worth) and
   that the chart axis shows 2027 to 2057.
4. Set D33 = 20 and confirm Maintenance 5 and 6 drop to $0 discounted and the chart ends at 2047. On
   Maintenance Policies, the salvage row of Table 1 should now read "16 of 16 years left on the mill
   and overlay placed in year 20 (100.0% of its cost) at the 20-year analysis period", and Table 2's
   should read 20 of 40 (50.0%). Set D33 back to 30 and they return to 37.5% and 25%.
5. Pick an airport outside the 17 with D38 = Yes: G2 on each Alt sheet should show the warning and
   lost revenue should be $0, not #N/A.
6. Click the navigation buttons on General Information, on the Summary and on an alternative sheet.
7. On the Summary, check the pavement section block: the derived thickness should match the section you
   designed, and the excavation check should agree with it.
8. On an alternative sheet, open a pay item picker in column C. The list should show the pay item
   number and the whole description, not a truncated fragment. Tell us if it still reads short: the
   list width is set explicitly but the split between the two columns is left to Excel, and there is a
   further property we can set if Excel splits it against the control rather than the list.
9. On the same sheet, set **Price from** (C11) to Middle and put a cost in Pay_Items column G for one
   item you used. That item should reprice; everything else should hold at the statewide Unit Cost.

`TDOA_LCCA_v1.2.0_dropdown-test.xlsm` is the same build with TMP(NewHMA) left visible, so steps 8 and
9 can be done in about ten seconds without enabling macros or running Alternative Setup: open it, go
to the TMP(NewHMA) tab, click a picker in column C. Nothing on that copy feeds a real analysis; it
exists only for that check and can be deleted afterwards.

Scripted click-through done here (LibreOffice, UNO API): every button target exists and lands on a
visible sheet; changing D34 to 7 flips the verdict to Alternative 1 (the 7% flag already warned);
D33 = 20 drops HMA closure days from 57 to 43 and PCC rehabilitation to $0; an airport outside the
17 gives lost revenue $0 with the warning in G2 and no error cells; one alternative and none both
read cleanly. A second round emulated the VBA's Database row deletion (Summary shifts up, no
#REF!), a tie between two alternatives, four alternatives at once, the airport dropdown (79
entries, no blanks or duplicates), unique pay-item descriptions, and the Summary page setup
(landscape, one page wide). Values were not saved. A package audit (verification/audit_xml.py)
checks cell order, style and string counts, relationships, content types and that only one sheet
is selected on open.

Verification done here (LibreOffice Calc, full recalculation): the template workbook recalculates
with 2,240 formulas and 0 errors (v1.1.2 shows 99 error cells under the same recalc). The same
patch applied to the populated MBT workbook gives NPW $8,572,110 and $8,028,734 (asphalt moves only
because of the salvage fix, items 12 and 13; every other line is unchanged),
the year-indexed columns reconcile to the NPW to the dollar, and the Summary table, verdict line
and six charts populate from the MBT data (render in `Summary_sheet_MBT_render.png`). Not
verified: the charts on the alternative worksheets in Excel itself (LibreOffice does not render
charts on the ActiveX-bearing sheets); step 3 above covers it.

For the v1.2.0 pay item and pricing changes specifically: 139 static checks on the built package all
pass; all 105 combo box streams were re-parsed with an independent implementation of [MS-OFORMS] and
each consumes exactly to the end of its stream; a scripted run copies an alternative template, prices
it at Regular, fills a Middle cost for one item and confirms that item reprices while every other item
holds at Unit Cost, that East with nothing filled prices everything at Unit Cost, and that an empty
picker reads as Regular rather than producing an error (10 checks, all pass). A second, end-to-end run
on the populated four-alternative McKellar-Sipes example doubles the Middle price of every item one
alternative uses and confirms the chain all the way to the Summary: nothing moves while that sheet
says Regular, switching it to Middle exactly doubles its initial construction and raises its net
present worth, the other three alternatives are untouched, and switching back restores every number to
the cent (8 checks, all pass).

Both worked examples were re-run against the built file and deep-compared with the pre-change results,
leaf by leaf and numbers to the cent: 1,124 values for Gatlinburg and 1,368 for McKellar-Sipes, with
exactly one difference in each, General Information D12 going from blank to Sevier and to Madison. A
cell-level diff of the whole workbook against the previous build shows 643 changed cells, all of them
on the five alternative templates, General Information, Pay_Items, Typical Values and Method; the
Summary, Database, RevenueData, Maintenance Policies, Overview and Instructions sheets are untouched.
The printed worked example is 30 pages, unchanged, so the wider picker column costs no paper.

Not verified here: how Excel itself divides the drop-down list between the two columns, which is what
step 8 above asks for.
