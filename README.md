# RealCost LCCA Dashboard

Single-file web reimplementation of the deterministic engine in FHWA RealCost 2.5
(Caltrans 2.5.7CA) for pavement life-cycle cost analysis. Built because the
macro-enabled `RealCostV2.5.7CA 64bit.xlsm` workbook is blocked by endpoint
security on managed machines — this needs no macros, no install, no Excel.

## What it does

- Agency-cost NPV per alternative: construction costs discounted from year of
  action, annual maintenance year-by-year, optional remaining-service-life
  credit at the end of the analysis period.
- Work-zone user costs (FHWA "calculated" method): hourly demand vs work-zone
  capacity, deterministic queuing, work-zone traversal delay, speed-change
  delay/VOC from the FHWA Aug-1996 tables (CPI-escalated to 2012$ as in
  RealCost 2.5.7CA), idling VOC, monetized with per-class values of time.
- Results: NPV / EUAC comparison table, stacked NPV chart, expenditure-stream
  diagram, hour-by-hour queue inspector.
- Scenario export/import (JSON), results export (CSV), auto-save to the
  browser's localStorage.

## Onboarding features

Because LCCA is unfamiliar to most people who need its output:

- **Guided tour** (11 steps) explaining what the analysis is for and what each
  section does, with spotlight highlighting.
- **Worked example** loading realistic demo costs for a 5-mile rehabilitation,
  so the charts and comparison are populated immediately rather than showing
  zeros. Clearly banner-flagged as illustrative.
- **Plain-language help** on every jargon-heavy field (vphpl, queue dissipation
  capacity, remaining service life credit, EUAC, discount rate…).
- **"What this means"** panel that writes the comparison out in prose, flags
  when the margin is inside estimate noise, and calls out the classic
  cheapest-to-build-isn't-cheapest-to-own result when it occurs.
- **Discount-rate sensitivity table** re-running every alternative from 2–6% and
  warning when the winner is not robust to that assumption.

Seeded with the analysis options, values of time, and the four alternative
M&R sequences from `RealCostV2.5.7CA 64bit.xlsm`. Construction costs and
work-zone details were blank in that workbook and default to zero — enter them
before comparing alternatives.

Not implemented (vs full RealCost): Monte Carlo simulation, COZEEP/TMP cost
modules, forced-flow queue speed curve.

## Deploy to Vercel

The app is a single static `index.html` — no build step.

1. In Vercel: **Add New → Project**, import this repo.
2. Framework preset **Other**. Leave Root Directory at the repo root —
   `index.html` is served as-is and there is no build step.
3. Deploy. Every push to `main` redeploys automatically.

Or from a terminal in a clone: `npx vercel`.

`vercel.json` sets `cleanUrls` and a few conservative security headers, and
marks `index.html` as must-revalidate so a deploy is picked up immediately.

## Credits and status

The methodology is FHWA's (RealCost; *Life-Cycle Cost Analysis in Pavement
Design*, FHWA-SA-98-079), adapted for California by Caltrans as RealCost
2.5.7CA. The added-time/added-cost and idling tables and the speed–volume
relationship come from the FHWA Technical Bulletin referenced in the RealCost
documentation; default hourly distributions derive from MicroBENCOST.

This is an **independent, unofficial reimplementation** written from the
published methodology and the input structure of the RealCost workbook. It is
not produced, reviewed, endorsed, or supported by FHWA or Caltrans, and is not
a certified substitute for the official tool. Verify results against official
RealCost before using them in project decisions or submittals.

## Parameter studio

Several inputs cannot be understood from a dropdown label. The studio (sidebar →
Input → Parameter studio) gives each of them a screen:

- **Hourly distribution** — every option drawn as a 24-hour profile, labelled
  with its peak hour and night share, and priced per day against the current
  work zone so the cheapest is obvious. Includes a builder for your own curve,
  saved onto the scenario.
- **Closure window** — click hours to build a window and watch cost and queue
  respond; rank every contiguous window of the same length by cost and apply
  the pick.
- **Capacity & speed** — typical HCM/Caltrans values with sources, and what a
  lane closure does to directional capacity at your settings.
- **Value of time** — the blended rate by class and share, and the effect of
  shifting the truck mix.

## Development

There is no build step and no dependencies. Open `index.html` in a browser, or
serve the folder with anything (`python3 -m http.server`). Everything — engine,
styles, reference data — lives in that one file so it can be dropped on any
host or opened straight from disk on a locked-down machine.

## Where the pricing lives

Unit costs and markups are **editable in the app** — section 6, "Pavement
Structures & Cost Basis". Type over any figure; estimates recalculate, any
activity cost that was derived from the estimator is re-derived, and the values
persist in the browser and travel with an exported scenario JSON. "Restore
workbook defaults" reverts.

The seed values come from the source workbook's hidden sheets, and their
defaults live in the JS constants near the top of the script:

| Constant | What it holds | Source |
|---|---|---|
| `UNIT_COSTS_DEFAULT` | material $/m³ (PCC, RSC, FSHCC, LCB, HMA, RAC) | `ESTIMATE` sheet G10:G16 |
| `MARKUPS_DEFAULT` | earthwork → supporting cost percentages | `ESTIMATE` sheet rows 19–33 |
| `CAPM_UNITS` | CAPM $/ton, $/SY, $/day rates | `capm` sheet |
| `STRUCTURES` | layer types and thicknesses | `Rehab_Thick` sheet |
| `ADDED_TIME` / `ADDED_COST` / `IDLING` | speed-change and idling tables | `Supplemental Data` sheet |
| `DISTRIBUTIONS` | hourly traffic distributions | `Supplemental Data`, `traffic_c` |

These are the workbook's illustrative example prices, not current bid prices.
Replace them with Caltrans Contract Cost Data values before real use.

## PDF report

The **PDF report** button opens the browser print dialog; choose "Save as PDF".
A dedicated print stylesheet produces a seven-page report: title block with
project identification, all inputs as static values, the full activity schedule
for every alternative, results and sensitivity tables, pavement cross-sections,
the unit-price basis, charts, and the methodology and disclaimer. No library is
involved, so nothing has to load past the artifact CSP.
