# Sources behind the "Typical Values" sheet (research done 12 Sept 2026)

Method: two web research passes. Direct fetches of faa.gov, whitehouse.gov, tn.gov, tasp2040.com, dot.ca.gov,
aeronautics.sc.gov and several others were blocked by the build environment's egress proxy, so values were
taken from search-engine excerpts of those documents and from secondary sources that quote them. Every row on
the sheet carries its URL. Confirm the primary paragraphs before quoting them to FAA or TDOT management.

## Tennessee system (TASP 2040, TDOT Aeronautics, APTech)
- 78 public-use airports; 6 commercial service, 72 general aviation (TASP inventory, March 2021).
- 69 NPIAS airports (TASP Chapter 7, 2020); NPIAS 2025-2029 Appendix A lists each airport's role.
- APTech Tennessee APMS: about 70 airports, PCI per ASTM D5340, last update 2022; TDOT also had APTech
  build a pavement LCCA tool for the APMS. PCI objectives 78 runway / 75 taxiway (TN Airport Management Guide).
- TASP system-wide cost estimates (Sept 2021): pavement rehabilitation $70.9M; new/reconstruction $173.1M.
- Grant shares: up to 90% state or federal on GA capital projects (TDOT Policy 170-02); 95/5 on federal
  discretionary; the Statewide Airfield Pavement and Marking Maintenance Program is 100% state funded.
- Recent projects: Smyrna 1/19 reconstruction (5,546 x 100 ft) about $13M, 2025; McKellar-Sipes 2/20 rebuild
  (6,005 x 150 ft) $28M low bid, 2024; Dyersburg maintenance project $820,000 (crack seal, seal coat, markings).
- Runway sample (FAA 5010 via Wikipedia/AirNav): MBT 4,753 x 100; SYI 5,503 x 100; GKT 5,506 x 75;
  CKV 5,999 x 100; MKL 6,005 x 150; MQY 8,038 x 150 and 5,546 x 100.

## Unit costs
- P-401 surface course: $185/ton low bid vs $232/ton engineer's estimate, Florence SC Taxiway A, May 2025.
- P-101 milling: $4.70/SY (3-in) to $10/SY (variable depth), same bid tab.
- P-501: $165/SY for 8-in doweled apron as a planning value (MnDOT SASP Appendix F); no southeastern
  2024-2025 PCC bid was visible. P-209, P-154, seal coats and crack sealing: not visible in excerpts; the TDOT
  APMM 2022 Summary of Bids and the SC Aeronautics bid tabulations hold them.
- Workbook defaults (Pay_Items column F, from the 2022 framework) are shown live beside these for comparison.

## Economic parameters
- FAA AIP PGL 22-01 (29 June 2022): use OMB Circular A-94 Appendix C real rates for pavement
  cost-effectiveness; removes the fixed 7% rate and the fixed 20-year analysis period.
- OMB A-94 Appendix C, March 2026 (M-26-09): real rates 20-year 2.0%, 30-year 2.0%. January 2025: 10-year
  1.8%, 30-year 2.3% (via NIST IR 85-3273-40). OMB M-25-23 (April 2025) reinstated the 1992 Circular (7% base
  case for benefit-cost analysis) but kept Appendix C for cost-effectiveness.
- Legacy: AIP Handbook 5100.38D table requiring 7% and 20 years (pre-2022); AC 150/5320-6E Appendix 1
  suggested 4% and 20 years with straight-line residual value (dropped in 6F/6G).
- AAPTP 06-06 (Auburn/NCAT): NPW, 4%, 20 years, remaining-life salvage; user cost includes lost airport revenue.
  Remaining-life salvage is a prorated share of the last treatment: remaining life divided by expected life,
  times the cost of that treatment. That method is AAPTP's and FAA's, and it is the form the workbook uses.
  The two fractions the workbook carried were not theirs: they came with the 2022 APTech framework as constants.
  Checked against the workbook's own schedules, the concrete one held at the default period (10 of 40 years left
  at year 30 is 25%) and the asphalt one did not: Table 1 places the mill and overlay at year 20, so a 16-year
  overlay life leaves 6 of 16 years at year 30, or 37.5%, not the 12.5% the sheet carried. 2 of 16 corresponds to
  an overlay at year 16. v1.2.0 computes both from remaining life instead, so they follow the analysis period and
  the overlay year rather than being typed; the expected lives they divide by (16 years for an overlay, 40 for
  concrete) are the workbook's own and are now shown in their own column for review.
  Direct fetches of eng.auburn.edu and faa.gov were blocked by the egress proxy; the method wording is from
  search-engine excerpts of the report. Confirm against the PDF before quoting it.
- Caltrans LCCA manual (2013): 4% real; 20/35/55-year periods; RealCost reports agency and user cost as
  present value and EUAC per alternative, then the lowest-cost alternative.

## Maintenance and service life
- AC 150/5380-7B: PCI inspection at least every 3 years at AIP airports; preventive-maintenance window PCI 60-80.
- FAA 2014 performance-trends study (quoted by ACPA): asphalt runways reach PCI 70 in about 12-15 years; PCC in
  about 40 years. FAA Extended Airport Pavement Life program targets 40 years.
- Joint sealant life (industry): hot-pour 3-8 yr, silicone 8-15 yr, compression seals 15-20 yr.
- AC 150/5320-6G minimums: P-401 surface 4 in (>= 12,500 lb), P-209 base 6 in, P-501 6 in (5 in under 12,500 lb).
- ACRP Report 159 and ACRP Synthesis 22 hold treatment fact sheets (life, cost, frequency).

## Items not confirmed (primary PDFs unreachable)
Exact AIP Handbook paragraph/table numbers for the legacy 7%/20-year rule; AC 150/5320-6G thickness table rows
for aircraft under 12,500 lb; the 2025 20-year real rate; verbatim RealCost results column labels; P-209, P-154,
seal-coat and crack-seal bid prices in Tennessee.
