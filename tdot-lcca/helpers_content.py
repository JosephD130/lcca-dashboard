"""Content of the 'Typical Values' sheet. Every row carries its source; 'this workbook' rows are formulas that read
the framework itself. Ranges from outside sources are as found in Sept 2026 and are for sanity checks only."""
import build_helpers as H

TASP = 'https://www.tasp2040.com/wp-content/uploads/2021/03/Updated-TASP-Inventory-Existing-System-Conditions-March-2021-v3.pdf'
TASP7 = 'https://www.tasp2040.com/wp-content/uploads/2020/12/Chapter-7-Airport-Classifications-and-NPIAS-Evaluation.pdf'
TASPC = 'https://www.tasp2040.com/wp-content/uploads/2021/09/TASP-Cost-Estimate-Chapter.pdf'
MG = 'https://www.tn.gov/content/dam/tn/tdot/aeronautics/TN%20Airport%20Management%20Guide%20FINAL.pdf'
APT = 'https://www.appliedpavement.com/tennessee-apms'
IDEA = 'https://idea.appliedpavement.com/hosting/tennessee/statewide-summary/inventory-area/area-by-airport-table.html'
NPIAS = 'https://www.faa.gov/sites/faa.gov/files/airports/planning_capacity/npias/current/ARP-NPIAS-2025-2029-Appendix-A.pdf'
P170 = 'https://www.tn.gov/content/dam/tn/tdot/documents/170-02.pdf'
SCFLO = 'https://aeronautics.sc.gov/sites/default/files/2025-05/250521_FLO_Rehab_Twy_A.pdf'
SCIDX = 'https://aeronautics.sc.gov/airport-development/sc-airport-capital-project-bid-tabulations'
APMM = 'https://www.tn.gov/tdot/aeronautics/apmm-contract.html'
MNSASP = 'https://www.dot.state.mn.us/aero/planning/documents/sasp/saspappendixf.pdf'
SMYRNA = 'https://www.smyrnaairport.com/news-events/smyrna-airport-runway-119-grand-opening'
MKLNEWS = 'https://www.wbbjtv.com/2024/08/28/mckellar-sipes-regional-airport-set-to-complete-runway-2-20-paving-preparing-for-new-jet-service/'
DYER = 'https://www.stategazette.com/news/regional-airport-receives-state-grant/article_5fd852fd-1d67-525c-b6fc-b32fbc676bdc.html'
PI = "'Pay_Items'"

def pay(desc_row, label):  # live unit cost from Pay_Items (default column F)
    return f'=IFERROR(VLOOKUP("{label}",Pay_Items!$D:$F,3,0),"")'

SECTIONS_SYSTEM = [
 ('1. What TDOT Aeronautics manages (context for project size)',
  ['Item', 'Value', 'Unit', 'Year', '', '', 'Source'],
  [['Public-use airports in the Tennessee system (TASP 2040)', 78, 'airports', 2021, '', '', TASP],
   ['Commercial service / general aviation', '6 / 72', 'airports', 2021, '', '', TASP],
   ['Airports in the FAA NPIAS (eligible for AIP)', 69, 'airports', 2021, '', '', TASP7],
   ['Airports in the TDOT pavement management network (APTech APMS, PCI per ASTM D5340)', '~70', 'airports', 2022, '', '', APT],
   ['PCI performance objective, runways / taxiways', '78 / 75', 'PCI', 2021, '', '', MG],
   ['TASP system-wide estimate: pavement rehabilitation (10-yr life goal)', 70865874, '$', 2021, '', '', TASPC],
   ['TASP system-wide estimate: pavement new / reconstruction (20-yr life goal)', 173117648, '$', 2021, '', '', TASPC],
   ['State share on GA capital projects (federal or state funded)', 'up to 90% (10% local); 95/5 on federal discretionary', '', 2023, '', '', P170],
   ['Statewide Airfield Pavement and Marking Maintenance Program (crack seal, seal coat, markings)', '100% state funded, no local share', '', 2024, '', '', APMM],
   ['Pavement areas and PCI by airport (statewide tables)', 'see APTech IDEA portal', '', 2022, '', '', IDEA],
   ['NPIAS 2025-2029 classification of each TN airport (National / Regional / Local / Basic)', 'see Appendix A', '', 2024, '', '', NPIAS]],
  'The workbook dropdown holds the TASP airports (count in the live section below). The 17 airports with revenue data are those meeting Aeronautics\' activity criteria.'),
 ('2. Typical runway geometry at Tennessee GA airports (for D26 mainline area and D28 markings)',
  ['Airport', 'Runway', 'Length x width (ft)', 'Mainline area (SY)', 'Role', 'Note', 'Source'],
  [['Murfreesboro Municipal (MBT)', '18/36', '4,753 x 100', 52811, 'GA, Middle', 'workbook MBT run used 52,778 SY', 'https://en.wikipedia.org/wiki/Murfreesboro_Municipal_Airport'],
   ['Bomar Field, Shelbyville (SYI)', '18/36', '5,503 x 100', 61144, 'GA, Middle', '', 'https://en.wikipedia.org/wiki/Shelbyville_Municipal_Airport_(Tennessee)'],
   ['Gatlinburg-Pigeon Forge (GKT)', '10/28', '5,506 x 75', 45883, 'GA, East', '75-ft width', 'https://en.wikipedia.org/wiki/Gatlinburg%E2%80%93Pigeon_Forge_Airport'],
   ['Outlaw Field, Clarksville (CKV)', '17/35', '5,999 x 100', 66656, 'GA, Middle', 'used in the new-project test', 'https://en.wikipedia.org/wiki/Clarksville%E2%80%93Montgomery_County_Regional_Airport'],
   ['McKellar-Sipes Regional (MKL)', '2/20', '6,005 x 150', 100083, 'GA, West', 'rebuilt 2024, ~$28M all-in', MKLNEWS],
   ['Smyrna (MQY)', '1/19', '5,546 x 100', 61622, 'GA, Middle', 'reconstructed 2025, ~$13M all-in', SMYRNA],
   ['Smyrna (MQY)', '14/32', '8,038 x 150', 133967, 'GA, Middle', 'former military', 'https://en.wikipedia.org/wiki/Smyrna_Airport_(Tennessee)'],
   ['Typical TN GA runway', '', '4,000 to 6,000 x 75 or 100', '33,000 to 67,000', '', 'TASP Regional Service minimum width 100 ft', TASP7],
   ['Markings area rule of thumb', '', '', '', '', '2,000 to 3,000 SF per 1,000 ft of runway for a full runway marking set (centerline, thresholds, designators, aiming points, TDZ where applicable); use the marking plan quantity when available', 'engineering judgement; verify with the marking plan']],
  'Mainline area = length x width / 9. Shoulders (D27) are usually unpaved at TN GA airports; enter paved shoulders only. Area drives the closure-day defaults on the alternative sheets.'),
]

SECTIONS_COSTS = [
 ('3. Unit costs: workbook defaults (Pay_Items, column F) against recent bids',
  ['Pay item', 'Workbook default (live)', 'Unit', 'Recent bid range', 'Year', 'Reading', 'Source'],
  [['P-401 Asphalt Surface Course', pay(37, 'Asphalt Surface Course'), '$/ton', '185 (low bid) to 232 (engineer\'s estimate)', 2025, 'Workbook default looks low against 2025 southeastern bids; confirm before relying on the default', SCFLO],
   ['P-401 Asphalt Base Course', pay(38, 'Asphalt Base Course'), '$/ton', 'typically 5 to 10% below surface course', '', '', SCIDX],
   ['P-501 Concrete Pavement, 9-inch', pay(46, 'Concrete Pavement, 9-inch'), '$/SY', '165 (8-in doweled apron, planning value)', 2023, 'Planning value from another state DOT; obtain a TN or southeastern PCC bid before use', MNSASP],
   ['P-209 Crushed Aggregate Base Course', pay(29, 'Crushed Aggregate Base Course'), '$/CY', 'not found in open sources; see SC bid tabs', '', '', SCIDX],
   ['P-154 Subbase Course', pay(19, 'Subbase Course'), '$/CY', 'not found in open sources; see SC bid tabs', '', '', SCIDX],
   ['P-101 Cold Milling', pay(16, 'Cold Milling'), '$/SY', '4.70 (3-in, low bid) to 10.00 (variable depth, EE)', 2025, '', SCFLO],
   ['P-608 Emulsified Asphalt Surface Treatment', pay(48, 'Emulsified Asphalt Surface Treatment'), '$/SY', 'TDOT APMM 2022 Summary of Bids', 2022, 'State maintenance contract prices apply to seal coats and crack sealing', APMM],
   ['P-620 Markings (reflective)', pay(56, 'Markings (prep, markings, reflective media)'), '$/SF', 'TDOT APMM 2022 Summary of Bids', 2022, '', APMM],
   ['All-in runway reconstruction, TN 2024-2025 (pavement + lighting + grading)', '', '$/SY', '210 (Smyrna 1/19) to 280 (MKL 2/20)', 2025, 'Project-level figures; the LCCA compares pavement pay items only', SMYRNA],
   ['State maintenance project example (Dyersburg): >200,000 LF crack seal, >121,000 SY seal coat, markings', 820000, '$', '', 2023, 'about $5/SY of runway treated', DYER]],
  'Pay_Items column F holds the default unit cost used by the alternative sheets (regional columns G:I are blank in v1.2.0). The defaults date from the 2022 framework; the 2025 comparison rows suggest the asphalt default is below current bids. Updating Pay_Items is the only place a cost needs to change.'),
]

SECTIONS_CLOSURE = [
 ('4. Closure-day defaults built into the alternative sheets (cells F4:F10; override with project-specific values)',
  ['Operation', 'Production rate in the formula', 'Days for a 6,000 x 100 ft runway (66,667 SY, 15,000 SF markings)', 'Where used', 'Note', '', 'Source'],
  [['Surface treatment (seal coat)', '15,000 SY/day', '=ROUNDUP(66667/15000,0)', 'HMA Maintenance 1-6', 'each operation is rounded up to whole days', '', 'this workbook (MBT and SRB runs)'],
   ['Markings: paint / layout and removal', '20,000 SF/day and 50,000 SF/day', '=ROUNDUP(15000/20000,0)+ROUNDUP(15000/50000,0)', 'every HMA activity', '', '', 'this workbook'],
   ['Patching', '10,000 SY of pavement per day at the policy rate (0.75% to 1.25% of area)', '=ROUNDUP(66667*0.0075/10000,0)', 'HMA Maintenance 3, 4, 6', '', '', 'this workbook'],
   ['Crack sealing', '5,000 SY of pavement per day at the policy rate (3.25% to 4% of area)', '=ROUNDUP(66667*0.0325/5000,0)', 'HMA Maintenance 3, 4, 6', '', '', 'this workbook'],
   ['Mill and overlay (4 in)', '3,800 SY/day', '=ROUNDUP(66667/3800,0)+ROUNDUP(15000/50000,0)', 'HMA Rehabilitation 1', 'the largest HMA closure', '', 'this workbook'],
   ['PCC joint resealing', 'about 1 LF of joint per SY of slab, 10,000 LF/day', '=ROUNDUP(66667*550/(50*100/9)/10000,0)', 'PCC Maintenance 1 and Rehabilitation 1', '', '', 'this workbook'],
   ['PCC crack sealing, partial- and full-depth patching', '5,000 to 10,000 SY of pavement per day at the policy rate', '=ROUNDUP(66667*0.003/10000,0)+ROUNDUP(66667*0.0013/5000,0)+ROUNDUP(66667*0.0025/5000,0)', 'PCC Maintenance 1', '', '', 'this workbook'],
   ['PCC slab replacement (2% of area) + cure', '1,000 SY/day plus 7 days cure', '=ROUNDUP(66667*0.02/1000,0)+7', 'PCC Rehabilitation 1', 'the largest PCC closure', '', 'this workbook'],
   ['Typical totals over 30 years', '', 'HMA 57 to 67 days in 7 closures; PCC 27 to 29 days in 2 closures', 'Summary column Q', 'MBT (52,778 SY) and CKV test (66,667 SY)', '', 'this workbook']],
  'Initial construction is not a closure in the model (the runway is assumed closed for construction regardless of alternative). Lost revenue = closure days x airport daily revenue x 100%; a per-category "share lost" factor is a policy decision listed in the change notes.'),
]
PGL = 'https://www.faa.gov/airports/aip/guidance_letters/aip_pgl_22_01'
A94C = 'https://www.whitehouse.gov/wp-content/uploads/2026/03/CircularA-94Appendix-C.pdf'
M2523 = 'https://www.whitehouse.gov/wp-content/uploads/2025/02/M-25-23-Rescission-and-Reinstatement-of-Circular-No.-A-94.pdf'
AC6G = 'https://www.faa.gov/documentLibrary/media/Advisory_Circular/150-5320-6G-Pavement-Design.pdf'
AC6E = 'https://www.faa.gov/documentLibrary/media/Advisory_Circular/150_5320_6e_part4.pdf'
CALFAQ = 'https://dot.ca.gov/programs/maintenance/pavement/faq/lcca'
CALMAN = 'https://dot.ca.gov/-/media/dot-media/programs/maintenance/documents/office-of-concrete-pavement/life-cycle-cost-analysis/lcca-25ca-manual-final-aug-1-2013-v2-a11y.pdf'
AAPTP = 'https://www.eng.auburn.edu/research/centers/ncat/files/aaptp/Report.Final.06-06.pdf'
ACPA = 'https://www.acpa.org/wp-content/uploads/2025/03/ACPA-Achieving-the-Goals-of-the-FAAs-Extended-Airport-Pavement-Life-Program.pdf'
AC7B = 'https://www.faa.gov/documentlibrary/media/advisory_circular/150-5380-7b.pdf'
AC6C = 'https://www.faa.gov/documentlibrary/media/advisory_circular/150-5380-6c.pdf'
ACRP159 = 'https://nap.nationalacademies.org/catalog/23649/'
MP = "'Maintenance Policies'"

SECTIONS_ECON = [
 ('5. Economic parameters (D33 analysis period, D34 discount rate, salvage)',
  ['Parameter', 'Value', 'Applies to', 'Year', '', 'Reading', 'Source'],
  [['TDOT policy discount rate (workbook default)', "='General Information'!$D$34", '%', '', '', 'live from General Information D34', 'this workbook'],
   ['TDOT policy analysis period (workbook default)', "='General Information'!$D$33", 'years', '', '', 'live from General Information D33', 'this workbook'],
   ['FAA AIP: discount rate for pavement cost-effectiveness (PGL 22-01, June 2022)', 'OMB Circular A-94 Appendix C real rate for the analysis period', '%', 2022, '', 'Replaced the fixed 7% rule; the analysis period is now the engineer\'s choice (the fixed 20-year requirement was removed)', PGL],
   ['OMB A-94 Appendix C real rate, 20-year and 30-year (calendar 2026)', 2.0, '%', 2026, '', 'Revised March 2026 (OMB memo M-26-09); 2025 values were 1.8% (10-yr) and 2.3% (30-yr)', A94C],
   ['Legacy FAA AIP rule before PGL 22-01', '7% and 20 years', '', 2014, '', 'AIP Handbook (Order 5100.38D) pavement LCCA table; also the 1992 A-94 benefit-cost base case reinstated by OMB M-25-23 (April 2025) for BCA, not for cost-effectiveness', M2523],
   ['FAA AC 150/5320-6E Appendix 1 (superseded)', '4% and 20 years; straight-line residual value', '', 2009, '', 'Dropped in 6F/6G; still widely cited', AC6E],
   ['AAPTP 06-06 airport LCCA guide (Auburn/NCAT)', '4%, 20 years, NPW, remaining-life salvage; user cost includes lost airport revenue', '', 2011, '', 'Basis for the "airport daily revenue" approach used here', AAPTP],
   ['Caltrans LCCA manual / RealCost', '4% real; 20 / 35 / 55 years (CAPM / rehabilitation / new construction)', '', 2013, '', 'Agency and user costs reported as present value and EUAC; the Summary comparison block follows that layout', CALMAN],
   ['Structural design life (FAARFIELD, AC 150/5320-6G)', 20, 'years', 2021, '', 'design life, not service life', AC6G],
   ['Salvage: HMA', "='Maintenance Policies'!$D$32", 'fraction of the last mill-and-overlay', '', '', 'live from Maintenance Policies D32 (policy: 2 of 16 years remaining)', 'this workbook'],
   ['Salvage: PCC', "='Maintenance Policies'!$D$46", 'fraction of initial construction', '', '', 'live from Maintenance Policies D46 (policy: 10 of 40 years remaining)', 'this workbook'],
   ['Salvage method in FAA / AAPTP guidance', 'remaining life / expected life x cost of the last treatment (straight line)', '', 2011, '', 'the workbook policy follows this form; the fractions above are the policy decision flagged in the change notes', AAPTP]],
  'The Summary sensitivity chart runs 2% to 8% so both the current FAA rate (about 2%) and the legacy 7% are on the curve. G10 on the Summary states whether the lowest-cost alternative changes at either rate.'),
 ('6. Maintenance timing and service life (Maintenance Policies sheet against published experience)',
  ['Item', 'Workbook policy (live)', 'Published range', 'Year', '', 'Reading', 'Source'],
  [['HMA surface treatment / markings', "=\"years \"&'Maintenance Policies'!$E$10&\", \"&'Maintenance Policies'!$E$12&\", \"&'Maintenance Policies'!$E$26", 'preventive-maintenance window PCI 60-80; inspect PCI every 3 years (AIP airports)', 2014, '', 'FAA ACs give the PCI window, not fixed year intervals; 4-year cycles are the TDOT policy', AC7B],
   ['HMA patching + crack sealing + treatment', "=\"years \"&'Maintenance Policies'!$E$14&\", \"&'Maintenance Policies'!$E$18&\", \"&'Maintenance Policies'!$E$28", 'crack sealing, patching, seal coats per AC 150/5380-6C', 2014, '', '', AC6C],
   ['HMA mill and overlay', "=\"year \"&'Maintenance Policies'!$E$22", 'asphalt runways reach PCI 70 in about 12 to 15 years (FAA 2014 performance trends)', 2014, '', 'policy year 20 is at the optimistic end of the published range', ACPA],
   ['PCC joint resealing + patching', "=\"year \"&'Maintenance Policies'!$E$37", 'hot-pour sealant 3-8 yr; silicone 8-15 yr; compression seals 15-20 yr', '', '', 'policy year 19 assumes silicone-class sealant life', 'https://www.tarmacview.com/glossary/joint-sealant/'],
   ['PCC slab replacement (2%) + joints', "=\"year \"&'Maintenance Policies'!$E$41", 'PCC runways reach PCI 70 at about 40 years (FAA 2014)', 2014, '', '', ACPA],
   ['Expected service life', 'HMA 20 yr design; PCC 30-40 yr', 'FAA Extended Airport Pavement Life program targets 40 years', 2024, '', '', 'https://www.airporttech.tc.faa.gov/Portals/0/FactSheets/Extended%20Airport%20Pavement%20Life%20Factsheet_May%202024.pdf'],
   ['Treatment fact sheets with life and cost by treatment', '', 'ACRP Report 159 (GA airports) and ACRP Synthesis 22', 2016, '', 'use for project-specific overrides of the policy years', ACRP159],
   ['Minimum layer thickness, AC 150/5320-6G', '', 'P-401 surface 4 in (>= 12,500 lb); P-209 base 6 in; P-501 6 in (5 in under 12,500 lb)', 2021, '', 'check the pay-item quantities against these minimums', AC6G]],
  'Maintenance Policies drives every activity year and rate; this section only shows them next to the published ranges. Sources were read from search excerpts in September 2026 because the FAA and OMB sites could not be opened from the build environment; confirm the cited paragraphs before quoting them to FAA.'),
]
SECTIONS_LIVE_TAIL = []   # live sections are appended by build_helpers.live_sections()

def sections(extra=()):
    return SECTIONS_SYSTEM + SECTIONS_COSTS + SECTIONS_CLOSURE + SECTIONS_ECON + list(extra) + H.live_sections()

HINTS = {
 25: 'Construction year; the analysis period starts here.',
 26: 'Length x width / 9. TN GA: 4,000-6,000 ft x 75-100 ft wide.',
 27: 'Paved shoulders only; usually 0 at TN GA airports.',
 28: 'About 2,000-3,000 SF per 1,000 ft of runway.',
 33: 'TDOT policy 30 years; FAA leaves the period to the engineer.',
 34: 'TDOT 3%. FAA uses the OMB A-94 real rate, 2.0% in 2026.',
 36: 'Percent of pay items; workbook default 10.',
 37: 'Percent of pay items; now applied to initial construction too.',
 38: 'Yes counts lost revenue during closures (17 airports have data).',
}
