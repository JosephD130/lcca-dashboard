# Two worked examples on the v1.2.0 template

Both were built on the blank `TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm` through the LibreOffice UNO API
(`run_example.py`): General Information filled, Alternative Setup emulated exactly as the form does it (copy the
hidden indirect-cost templates to Alt sheets, register them on Database, write Summary A:E), pay items and
quantities typed on each Alt sheet, full recalculation, then every result block read back and cross-checked
independently. Nothing was pre-loaded. One catalogue gap was filled the way a user would: P-152 Unclassified
Excavation has no default unit cost on Pay_Items, so $12/C.Y. was entered there for both runs.

Quantities were derived from a target section so the read-back could be checked against a known answer:
C.Y. = inches x S.Y. / 36 for volume items; tons = inches x 145 pcf x S.Y. / 2,666.67 for asphalt.

Common settings: 30 years at 3 percent (TDOT), mobilization 10 percent, engineering 5 percent, indirect cost
(lost revenue) Yes, reflective markings.

## Example 1: Gatlinburg-Pigeon Forge (GKT), Runway 10-28, no shoulders, three alternatives

5,506 x 75 ft runway: 45,883 S.Y. mainline, 0 S.Y. shoulder, 12,000 S.F. markings, construction 2028.
Daily revenue from RevenueData: $5,441.17.

| Alternative | Section entered as quantities | Initial (x1.15) | NPW at 3% | Closure days | vs. lowest |
|---|---|---|---|---|---|
| Alt 1 New HMA | 4" P-401 surface + 4" P-401 base on 6" P-209 on 6" P-154, 20" excavation | $4,285,512 | $6,156,578 | 56 | +$304,331 (5.2%) |
| Alt 2 New PCC | 9" P-501 on 6" P-209 on 6" P-154, 21" excavation | $5,918,906 | $6,065,225 | 24 | +$212,978 (3.6%) |
| Alt 3 New PCC | 11" P-501 on 6" P-209, 17" excavation | $5,681,474 | $5,852,247 | 24 | lowest |

Verdict line: "Lowest present worth: Alternative 3 | margin to next: $212,978 | 3% over 30 years | lost revenue: Yes".
G10: same lowest-cost alternative at 2 percent; changes at 7 percent (chart 2 shows the HMA curve crossing
below both concrete curves between 3.5 and 4.25 percent).

Section read-back (G83:P85): 8.00 / 6.00 / 6.00 in, total 20.00, excavation 20.00, "agrees", string
`8" P401 on 6" P209 on 6" P154`; 9.00 / 6.00 / 6.00, total 21.00, "agrees", `9" P501 on 6" P209 on 6" P154`;
11.00 / 6.00 / 0.00, total 17.00, "agrees", `11" P501 on 6" P209`. Chart 7 and the shoulder string (column Q)
stay empty because no shoulder area was entered.

Two things worth knowing from this run. Dropping the 6 in subbase and going to 11 in concrete costs less up
front than 9 in on subbase ($6/S.Y. more concrete against $55/C.Y. of subbase plus 4 in of excavation), which is
why Alt 3 wins. And the concrete curves in chart 2 are nearly flat: the 25 percent salvage credit at year 30 is
discounted as hard as the maintenance stream, so PCC NPW moves less than 1 percent between 2 and 8 percent while
the HMA curve falls 19 percent. That is the salvage policy question in the change notes made visible.

## Example 2: McKellar-Sipes Regional (MKL), Runway 2-20, with shoulders, four alternatives

6,005 x 150 ft runway with 25 ft paved shoulders each side: 100,083 S.Y. mainline, 33,361 S.Y. shoulder,
25,000 S.F. markings, construction 2027. Daily revenue from RevenueData: $8,505.94. Asphalt, base and subbase
quantities were taken off the combined 133,444 S.Y.; concrete off the mainline only.

| Alternative | Section entered as quantities (combined area) | Initial (x1.15) | NPW at 3% | Closure days | vs. lowest |
|---|---|---|---|---|---|
| Alt 1 New HMA | 3" P-401 surface + 5" P-401 base on 6" P-209 on 6" P-154, 20" excavation | $12,372,457 | $16,528,084 | 94 | +$2,461,515 (17.5%) |
| Alt 2 New PCC | 9" P-501 (mainline) on 6" P-209 on 6" P-154, excavation 21" mainline + 12" shoulder | $13,830,666 | $14,066,568 | 36 | lowest |
| Alt 3 New HMA | 4" P-401 surface + 6" P-401 base on 10" P-209 on lime-treated subgrade, no excavation item | $14,671,662 | $18,827,288 | 94 | +$4,760,719 (33.8%) |
| Alt 4 New PCC | 11" P-501 (mainline) on 6" P-209 on lime-treated subgrade, excavation 17" mainline + 6" shoulder | $14,418,903 | $14,594,219 | 36 | +$527,651 (3.8%) |

Verdict line: "Lowest present worth: Alternative 2 | margin to next: $527,651 | 3% over 30 years | lost revenue: Yes".
G10: same lowest-cost alternative at 2 percent and at 7 percent.

Section read-back, mainline reading (divides by the 100,083 S.Y. mainline, so combined-area quantities read
4/3 thicker): Alt 1 10.67 / 8.00 / 8.00 in, total 26.67, excavation 26.67, "agrees"; Alt 2 9.00 / 8.00 / 8.00,
total 25.00, excavation 25.00, "agrees"; Alt 3 13.33 / 13.33 / 0, treated subgrade "yes", "no excavation
item"; Alt 4 11.00 / 8.00 / 0, treated subgrade "yes", total 19.00, excavation 19.00, "agrees".

Shoulder reading (chart 7 and the column Q string, filled because D27 > 0): the same quantities spread over
133,444 S.Y. give back the sections that were designed: `8" P401 on 6" P209 on 6" P154`,
`9" P501 on 6" P209 on 6" P154`, `10" P401 on 10" P209`, `11" P501 on 6" P209`. Concrete keeps its
named thickness in both readings.

## Checks run on each example (all pass: 43 on GKT, 61 on MKL)

Per alternative: item cost = quantity x unit cost; initial = subtotal x 1.15; every activity's present worth
= cost / 1.03^offset; NPW = sum of the activity column. Summary: results row matches the Alt sheet and the A:E
block the form writes; the four category columns sum to NPW; closure days equal the by-year closure column;
cumulative discounted cost ends at the NPW; the verdict names the lowest NPW; "vs. lowest" is zero only for
the winner; EUAC = PW x CRF and agency + user = total; the sensitivity curve at 3 percent reproduces the NPWs
and the HMA curves fall monotonically with the rate; section thicknesses, excavation check, description
strings and the shoulder chart rows match the designed sections to 0.02 in; no formula errors on any visible
sheet.

Files per example: `results.json` (everything read back), `*_LCCA_run.xlsx` (LibreOffice-saved copy of the
populated framework, for viewing only; open the .xlsm in Excel for the real thing), `*_Summary.png` and the
two crops, `*_Alt1.png`.
