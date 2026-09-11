# TDOA LCCA Framework v1.2.0 — change notes

Base: `TDOA_LCCA_Framework_v1.1.2_ARA_Task2_08182026.xlsm`
Output: `TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm` plus `Output.bas`

All workbook edits were made directly in the sheet XML so the VBA project, the 192 ActiveX
pay-item comboboxes, the charts, tables and data validations are untouched. The VBA change is a
single module replacement (`Output`) that must be imported in Excel; see "Install" below.

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

## 3. Summary worksheet (VBA, `Output.bas`)

`SetupSummaryWs` keeps its name and is still called when the Alternative Setup form closes. It now
writes, all as live formulas into the alternative sheets:

- Results table: Initial, Maintenance PW, Rehabilitation PW, Lost Revenue PW, Salvage PW, NPW,
  delta vs. lowest, closure days over the analysis period, description.
- Present worth by category (stacked column per alternative).
- Expenditure stream by calendar year (clustered column, undiscounted).
- Cumulative discounted cost by year (line).
- NPW vs. discount rate 2% to 8% (line), with a "same winner at 7% (FAA AIP)?" flag.

Charts are coloured by pavement type: HMA blue, PCC orange (second alternative of a type is a tint).

## 4. Housekeeping

- Instructions text box: ActiveX "blocked content" steps added (Trust Center > ActiveX Settings,
  then restart Excel; contact IT if greyed out); step 1 now says D9:D39; step 2 names the
  "Alternative Setup" button; step 4 rewritten for the hidden RevenueData sheet and the
  17-airport rule, including how to unhide a sheet; step 5 notes the pre-filled closure days.
- Cell notes on D26 and D39 no longer carry a person's name.
- Two broken external links to files on a C: drive (v1.1.004 and a Savannah copy) removed.
- Workbook set to fully recalculate on open.

## 5. Not changed (needs a decision)

- Overview text: Neel-Schaffer's rewrite (2022 APTech / 2026 NS+ARA history, the four airport
  criteria) is in their review doc; paste it once Mat approves the wording.
- Salvage asymmetry: PCC recovers 25% of total initial cost (incl. mobilization), HMA 12.5% of one
  mill-and-overlay. In SRB that line alone decides the result. Left as policy; the new sensitivity
  block makes it visible.
- Lost revenue counts gross fuel sales and tenant rent as lost during a runway closure. A per-category
  "% lost during closure" factor on RevenueData would be more defensible.
- FAA AIP funding (7%, 20-year life) is shown in the sensitivity block, not enforced.
- RevenueData hygiene: MBT and MQY rows are identical; XNX and M54 hold text where numbers belong.

## 6. Install and verify in Excel

1. Open `TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm`, enable content.
2. Alt+F11 > in the project tree right-click the `Output` module > Remove (No to export) >
   File > Import File > `Output.bas`. Save.
3. General Information: pick an airport, set D38 = Yes, D26 = 52777.7, D28 = 5433, D25 = 2027.
   Alternative Setup > add one New HMA and one New PCC > Close. Summary should populate with
   four charts and no #N/A.
4. On an Alt sheet check that columns L:N sum to the NPW table (sum of N = Net Present Worth) and
   that the chart axis shows 2027 to 2057.
5. Set D33 = 20 and confirm Maintenance 5 and 6 drop to $0 discounted and the chart ends at 2047.
6. Pick an airport outside the 17 with D38 = Yes: G2 on each Alt sheet should show the warning and
   lost revenue should be $0, not #N/A.

## 7. Verification record (LibreOffice Calc recalculation, forced full recalc)

| Check | Result |
|---|---|
| `TDOA_LCCA_Framework_v1.2.0` full recalculation | 1,644 formulas, **0 errors** (the original v1.1.2 shows 99 error cells under the same recalc: 39 #NAME? from XLOOKUP and 60 #N/A from the unguarded PCC lookups and the missing-airport revenue lookup) |
| Patched formulas applied to the populated MBT workbook | NPW unchanged: HMA $8,809,266.42 and PCC $8,028,734.49 (analysis period 30, so the new guard changes nothing); PCC rows B15:B22 now 0 instead of #N/A |
| New year-indexed chart columns (L:N) on the MBT alternatives | sum of L + M equals total direct + lost revenue; sum of N equals NPW to the dollar; bars land on 2027, 2031, 2035 ... 2057 as the policy years dictate |
| Patched chart XML | well-formed; parsed by openpyxl's chart reader with the expected category (K) and value (L, M) ranges; indirect charts stacked with two series |
| Zip integrity and XML well-formedness of every part | pass |
| `TDOT_LCCA_Decision_Workbook.xlsx` full recalculation | 1,370 formulas, **0 errors** |
| Decision workbook vs. independent Python engine | NPW, category split (sums to NPW), EUAC, closure days, break-even values (engine returns a $0 difference at each break-even), rate sensitivity, tornado, all 12 scenario rows and the 35-cell decision map agree to the dollar |
| Charts rendered through LibreOffice to PDF | all eight per-project charts and the two Monte Carlo charts draw with the right series, axes and years |

Not verified here: the `Output.bas` VBA module (no Excel in this environment) and the appearance of the patched xlsm charts in Excel (LibreOffice does not render charts on the ActiveX-bearing sheets). Both are covered by the six-step first-open check in section 6.

## 8. Decision workbook (`TDOT_LCCA_Decision_Workbook.xlsx`)

A companion, macro-free workbook that already carries the MBT and SRB runs and answers the questions the framework Summary cannot. One sheet per project; yellow cells B5:B14 drive everything.

Blocks: RESULTS (PW by category, NPW, EUAC, $/SY, closure days, runway availability, winner, margin, winner at 7%), BREAK-EVEN VALUES (daily revenue, PCC and HMA bid level, PCC salvage fraction, discount rate that tie the alternatives), DISCOUNT-RATE SENSITIVITY 2 to 8%, EXPENDITURE BY CALENDAR YEAR with cumulative PW and crossover year, WHAT COULD FLIP THE ANSWER (tornado, 8 inputs), SCENARIO SCORECARD (12 pre-run scenarios incl. FAA AIP 7%/20 yr, no lost revenue, 40% revenue loss, no salvage, bids +20%, closures x2 and x0.5, stress case), DECISION MAP (rate x salvage grid, colour-coded winner). Charts in column T: PW by category, expenditure stream, cumulative discounted cost, rate sensitivity, tornado, closure timeline (bubble), closure days by year, NPW by scenario. MonteCarlo sheet: 5,000 joint-uncertainty draws per project with P(PCC lower).

Headline readings at TDOT policy settings:

| | MBT | SRB |
|---|---|---|
| Lower NPW | PCC by $781k (8.9%) | HMA by $586k (4.1%) |
| Holds at FAA 7%? | No, flips at 4.5% | Yes |
| Holds with zero salvage? | Barely ($33k) | Yes |
| Closure days over 30 yr, HMA / PCC | 57 / 27 | 75 / 31 |
| PCC bid level that ties | +9.9% | -4.1% |
| P(PCC lower), Monte Carlo | 53% | 22% |

Both margins are inside estimating noise. In MBT the answer is a discount-rate and salvage call; in SRB, HMA wins on cost while PCC wins on closures, so the decision is what a runway-day is worth to Aeronautics.
