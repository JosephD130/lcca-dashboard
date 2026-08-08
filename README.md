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

1. Push this repo to GitHub (already done if you're reading this there).
2. In Vercel: **Add New → Project**, import the repo.
3. Set **Root Directory** to `lcca-dashboard`, framework preset **Other**.
4. Deploy. Every push to the branch redeploys automatically.

Or from a terminal: `cd lcca-dashboard && npx vercel`.

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
