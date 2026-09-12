# New-project test: Outlaw Field (CKV), Runway 17-35 reconstruction

Scenario entered on the v1.2.0 template (nothing pre-loaded): Clarksville, Middle region, construction 2028,
mainline 66,667 SY (6,000 x 100 ft), 15,000 SF reflective markings, 30 years at 3%, mobilization 10%,
engineering 5%, indirect cost Yes. Alternative Setup was emulated through the LibreOffice UNO API
(run_ckv.py): copy the two indirect-cost templates to Alt sheets, register them on Database, write Summary
A:E as the form does. Pay items: HMA = lime-treated subgrade, 8 in subbase, 6 in P-209, 5 in P-401 base,
4 in P-401 surface, markings; PCC = same foundation with 9 in P-501.

| Check | Expected (hand / independent) | Framework | Decision TEMPLATE |
|---|---|---|---|
| Daily revenue (RevenueData CKV) | 3,239.39 | 3,239.39 | input |
| Initial cost HMA / PCC (subtotal x 1.15) | 7,268,437 / 9,062,466 | same | same |
| Closure days HMA (F4:F10) | 7,7,9,9,19,7,9 = 67 | same | 67 |
| Closure days PCC (F4:F5) | 10, 19 = 29 | same | 29 |
| Salvage HMA (12.5% of rehab) / PCC (25% of initial) | -212,008 / -2,265,616 | same | same |
| NPW HMA / PCC | 9,829,223.75 / 9,180,602.25 (Python engine) | same | same |
| Category PWs (maint, rehab, lost, salvage) | engine | match to the dollar | match to the dollar |
| Verdict | PCC lower by 648,621; flips at 7% | G9/G10 say so | "No - decision depends on the discount rate", break-even 5% |
| Error cells after full recalc | 0 | 0 | 0 |

Files: CKV_Runway17-35_LCCA_run.xlsx (LibreOffice-saved copy of the populated framework, for viewing only),
CKV_summary.png (its Summary), CKV_decision.xlsx (decision workbook with the TEMPLATE filled), ckv_results.json.
