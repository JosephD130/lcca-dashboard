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
- Pavement section read-back G80:Q86, with charts 6 and 7: the thickness of each layer, derived from the
  pay-item quantities already entered against the mainline area, the total section, a cross-check against
  the excavation quantity, and the section written out as a string to paste into the alternative
  description. Asphalt is the only layer needing an assumption, and its unit weight sits in J81 (145 pcf,
  the value that reproduces the Murfreesboro section exactly). Chart 7 fills in only when a shoulder area
  is entered and shows the same quantities spread over mainline plus shoulder; column Q then also writes that
  shoulder reading out as a second string.
- Chart data lives in columns W onward, greyed and labelled "calculated automatically; do not edit".
- Summary layout: columns A to F are hidden. The Alternative Setup form still writes its small table and
  "Chart 1" there, and the results table from column G repeats all of it with more detail, so the table and the
  charts now start at the left edge of the screen instead of sitting beside a duplicate. The two navigation
  buttons moved to G1 and H1; unhide A:F if you want to type an alternative description. The original
  "Alternatives Comparison" chart is set to plot hidden cells so it keeps working, and its axis now reads in
  millions.
- Chart legends moved to the right of each plot: with the legend underneath it collided with the category
  labels and the axis title. The money axis reads to one decimal ($16.5M rather than five gridlines all
  reading $17M), and an alternative slot that has not been created carries no name, so the axis and the legend
  no longer invent "Alt 3" and "Alt 4".
- Every chart in the workbook now carries an x and a y axis title, including the five original alternative-sheet
  charts and the original "Alternatives Comparison" chart (calendar year against cost, alternative against
  present worth, discount rate against net present worth, alternative against thickness, and so on). The
  companion decision workbook got the same treatment.
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
source URL as plain text (no hyperlink, so nothing on the sheet reaches outside the workbook); see verification/RESEARCH_SOURCES.md for what could and could not be confirmed. General Information
gets a "Typical Values" button (D46) and short hints in column F beside D25:D38.

## 3ba. Project identity, locator map and Google Earth export

The Summary never said which project it was. It now opens with a one-line identity above the results table
(airport and FAA identifier, city, TDOT region, branch, project type, construction year, period and rate),
and a PROJECT LOCATION block to the right of the table: a locator map of Tennessee with all 79 airports in
the dropdown as grey dots and this project as a labelled orange dot, and under it the airport, city and
county, region, coordinates, elevation, branch and mainline area.

The map colours the airports by TDOT Grand Division, light blue for West, blue for Middle and navy for East, with a
legend beside it. That is the same grouping Pay_Items is built around, and it makes a gap visible: Pay_Items carries a
header "Average Pay Item Unit Cost" spanning Middle, West and East columns, and all three are empty. Every alternative
prices off the single Unit Cost column, so a West-division project is estimated on the same statewide numbers as a
Middle one. A line under the project block says so and names the project's division. Nothing was changed about how
costs are calculated; the Method sheet now lists it among the assumptions worth a decision.

The map is an ordinary scatter chart. The state outline and the airport coordinates are embedded in the
workbook (map-data block on the Summary at row 115, outside the print area, with the rest of the chart data),
so it draws with no internet connection, no Bing map service and no add-in: the file stays self-contained.
Coordinates come from airportsdata (PyPI, FAA/OurAirports values) and the outline from basemap-data's
intermediate-resolution political boundaries, both retrieved 13 September 2026 and cited in geo_data.py.
Five small private fields in the dropdown have no published coordinates; they plot no dot and the sheet says so.

A Google Earth export ships as a VBA module, `LCCA_KML_Export.bas`, to import once (Alt+F11, File, Import
File, then save). Running ExportLCCAKML writes a .kml beside the workbook containing the airport with the
whole result in its bubble (every alternative with initial cost, net present worth, closure days and section,
plus the verdict), a schematic runway footprint per alternative oriented from the runway number and sized
from the area entered, an extruded bar per alternative whose height is its net present worth, and every
maintenance and rehabilitation event as a placemark stamped with the calendar year it happens, so the time
slider walks the analysis period. The macro reads cells and writes a text file; it changes nothing and goes
nowhere, and the KML carries no external icon references so it renders on a machine with no internet.

The output was linted (verification/kml_check.py: element order, style resolution, closed rings, coordinate
range, time stamps, extruded geometry) and four things were corrected before release. Elements inside each
placemark now follow the KML 2.2 sequence (name, description, time, style, geometry); Google Earth tolerated
the earlier order but a strict GIS reader need not. The document carries a legend with colour swatches
(green for the lowest present worth, blue for the others, orange for an event) and a LookAt, so opening the
file flies to the runway at a sensible angle instead of leaving the reader to find it. The footprint length
now comes from a runway width entered on the Summary (T27, default 100 ft) rather than a fixed assumption:
at McKellar-Sipes the 100 ft assumption drew a 9,008 ft runway against the real 6,005 ft, and with the true
150 ft width the footprint lands on the pavement. The present-worth bars are scaled so the tallest is always
700 m, instead of a fixed 180 m per million which put a 3.4 km spike over a 1.8 km runway.

Samples produced from both worked examples are in verification/examples, with the checker output.

## 3bb. Method sheet

A second reference sheet, "Method" (last sheet, nothing on it feeds the calculation), states every calculation
once: the step, the cell it lives in, the formula as it stands there, the rule in plain English and where the
rule comes from. Sections 2 to 8 follow one alternative from the quantities typed on its worksheet to its net
present worth (initial cost, policy years, what each event costs, closure days, lost revenue, salvage,
discounting); section 9 covers what the Summary adds; section 10 reads the current period, rate, capital
recovery factor and the present worth of $1 at year 20 live from the file; section 11 lists the assumptions a
reviewer will ask about. Reason: the discounting, salvage and lost-revenue rules previously existed only as
formulas in cells, with nothing in the workbook stating them in words. General Information gets a "Method"
button (D48).

One inconsistency found while writing it, recorded on the sheet and not changed: the PCC cost blocks estimate
joint length from 12.5 x 12.5 ft slabs, while the closure-day formulas assume 550 L.F. of joint per 50 x 100 ft
panel, about a third less. It understates PCC closure days and the lost revenue derived from them.

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
- Tab colours group the sheets: navy for General Information, blue for Summary, light blue for the
  alternative worksheets, grey for reference sheets.

## 4. Housekeeping

- Instructions text box: ActiveX "blocked content" steps added (Trust Center > ActiveX Settings,
  then restart Excel; contact IT if greyed out); step 1 now says D9:D39; step 2 names the
  "Alternative Setup" button; step 4 rewritten for the hidden RevenueData sheet and the
  17-airport rule, including how to unhide a sheet; step 5 notes the pre-filled closure days;
  step 6 describes the new Summary and the "View Summary" button. The text box was extended
  (rows 9 to 76) so the longer text is not clipped.
- Cell notes on D26 and D39 no longer carry a person's name.
- Two broken external links to files on a C: drive (v1.1.004 and a Savannah copy) removed.
- Workbook set to fully recalculate on open.
- Package hygiene (nothing in the file refers outside it or to an individual): six cached printer-driver
  parts (printerSettings1-6) dropped with their relationships, page setups keep orientation and scaling;
  the empty Power Query (DataMashup) stub in customXml removed; the SharePoint path Excel cached as the
  last save location stripped from workbook.xml; document properties now read creator "TDOT Aeronautics
  Division / Applied Research Associates", last modified by "ARA", company "TDOT Aeronautics Division".
  The VBA project binary still carries the path of the machine it was last compiled on; that can only be
  refreshed by saving the file once in Excel.

## 5. Not changed (needs a decision)

- Overview text: Neel-Schaffer's rewrite (2022 APTech / 2026 NS+ARA history, the four airport
  criteria) is in their review doc; paste it once Mat approves the wording.
- Salvage asymmetry: PCC recovers 25% of total initial cost (incl. mobilization), HMA 12.5% of one
  mill-and-overlay. In SRB that line alone decides the result. Left as policy; the new sensitivity
  block makes it visible.
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
4. Set D33 = 20 and confirm Maintenance 5 and 6 drop to $0 discounted and the chart ends at 2047.
5. Pick an airport outside the 17 with D38 = Yes: G2 on each Alt sheet should show the warning and
   lost revenue should be $0, not #N/A.
6. Click the navigation buttons on General Information, on the Summary and on an alternative sheet.
7. On the Summary, check the pavement section block: the derived thickness should match the section you
   designed, and the excavation check should agree with it.

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
patch applied to the populated MBT workbook leaves NPW unchanged at $8,809,266 and $8,028,734,
the year-indexed columns reconcile to the NPW to the dollar, and the Summary table, verdict line
and six charts populate from the MBT data (render in `Summary_sheet_MBT_render.png`). Not
verified: the charts on the alternative worksheets in Excel itself (LibreOffice does not render
charts on the ActiveX-bearing sheets); step 3 above covers it.

## 7. Worked examples

Two made-up projects were run end to end on the blank template (verification/examples): Gatlinburg-Pigeon
Forge Runway 10-28 without shoulders and three alternatives (HMA, 9 in PCC on subbase, 11 in PCC without
subbase; the 11 in PCC is lowest, $5,852,247, and the answer changes at 7 percent), and McKellar-Sipes
Runway 2-20 with 25 ft shoulders and four alternatives (two HMA, two PCC; 9 in PCC is lowest, $14,066,568,
at every rate from 2 to 8 percent). Every block of the Summary was read back and checked against an
independent calculation (43 and 61 checks, all pass), and the section read-back returned the sections the
quantities were built from, on the mainline and, for the shoulder example, in the shoulder reading.
