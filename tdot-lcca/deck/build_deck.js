const pptxgen = require('pptxgenjs');
const fs = require('fs');
const IMG = __dirname + '/img/';
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625 in
pres.author = 'Applied Research Associates';
pres.title = 'TDOT Aeronautics LCCA Framework: review and v1.2.0 update';

// palette: charcoal ink, TN red accent used sparingly, HMA blue / PCC orange for data
const INK = '1F1F1F', MUTED = '5F5F5F', LIGHT = 'F3F3F1', WHITE = 'FFFFFF', DARK = '1D2733', RED = 'C8102E', BLUE = '2A78D6', ORANGE = 'EB6834', GREEN = '2E7D32';
const HF = 'Cambria', BF = 'Calibri';

function title(s, text, opts = {}) {
  s.addText(text, { x: 0.5, y: 0.35, w: 9.0, h: 0.7, fontFace: HF, fontSize: 28, bold: true, color: opts.color || INK, isTextBox: true, margin: 0, valign: 'top' });
}
function sub(s, text, y = 1.0, color = MUTED) {
  s.addText(text, { x: 0.5, y, w: 9.0, h: 0.4, fontFace: BF, fontSize: 13, color, isTextBox: true, margin: 0, italic: true });
}
function bullets(s, items, x, y, w, h, size = 13) {
  s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 5 } })),
    { x, y, w, h, fontFace: BF, fontSize: size, color: INK, isTextBox: true, margin: 0, valign: 'top' });
}
function img(s, file, x, y, w, ratio, opts = {}) {
  const h = w / ratio;
  s.addImage({ path: IMG + file, x, y, w, h, ...opts });
  return y + h;
}
function caption(s, text, x, y, w) {
  s.addText(text, { x, y, w, h: 0.45, fontFace: BF, fontSize: 10.5, color: MUTED, isTextBox: true, margin: 0, valign: 'top' });
}
function card(s, x, y, w, h, head, body, headColor = INK) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: LIGHT }, line: { color: LIGHT }, rectRadius: 0.08 });
  s.addText(head, { x: x + 0.18, y: y + 0.12, w: w - 0.36, h: 0.4, fontFace: HF, fontSize: 15, bold: true, color: headColor, isTextBox: true, margin: 0 });
  s.addText(body, { x: x + 0.18, y: y + 0.55, w: w - 0.36, h: h - 0.65, fontFace: BF, fontSize: 11.5, color: INK, isTextBox: true, margin: 0, valign: 'top' });
}
function stat(s, x, y, w, big, label, color = INK) {
  s.addText(big, { x, y, w, h: 0.75, fontFace: HF, fontSize: 30, bold: true, color, isTextBox: true, margin: 0 });
  s.addText(label, { x, y: y + 0.72, w, h: 0.75, fontFace: BF, fontSize: 11, color: MUTED, isTextBox: true, margin: 0, valign: 'top' });
}
function table(s, rows, x, y, w, colW, size = 10.5, headFill = 'E3E3E0') {
  const data = rows.map((r, i) => r.map(c => ({ text: c, options: { bold: i === 0, fill: { color: i === 0 ? headFill : WHITE }, color: INK, fontFace: BF, fontSize: size, valign: 'middle' } })));
  s.addTable(data, { x, y, w, colW, border: { type: 'solid', pt: 0.5, color: 'C9C9C4' }, rowH: 0.3, margin: 0.05 });
}
function footer(s, text) {
  s.addText(text, { x: 0.5, y: 5.2, w: 9.0, h: 0.3, fontFace: BF, fontSize: 9, color: MUTED, isTextBox: true, margin: 0 });
}

// ---------------------------------------------------------------- 1 title
let s = pres.addSlide(); s.background = { color: DARK };
s.addText('TDOT Aeronautics LCCA Framework', { x: 0.6, y: 1.5, w: 8.8, h: 0.9, fontFace: HF, fontSize: 36, bold: true, color: WHITE, isTextBox: true, margin: 0 });
s.addText('Review of v1.1.2 and the v1.2.0 update', { x: 0.6, y: 2.4, w: 8.8, h: 0.6, fontFace: BF, fontSize: 20, color: 'CFD8E3', isTextBox: true, margin: 0 });
s.addText('Airport pavement type selection with runway-closure lost revenue', { x: 0.6, y: 3.0, w: 8.8, h: 0.4, fontFace: BF, fontSize: 14, color: 'AEB8C4', isTextBox: true, margin: 0, italic: true });
s.addText('Applied Research Associates, Inc.  |  Task 005935.00000.00000.TASK2  |  September 2026', { x: 0.6, y: 4.7, w: 8.8, h: 0.35, fontFace: BF, fontSize: 11, color: 'AEB8C4', isTextBox: true, margin: 0 });
s.addNotes('Cover. Two deliverables: the patched framework workbook (v1.2.0) and a companion decision workbook with the MBT and SRB runs.');

// ---------------------------------------------------------------- 2 what we reviewed
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'What we reviewed');
sub(s, 'Five files from the Neel-Schaffer review package, read line by line down to the VBA');
card(s, 0.5, 1.55, 2.9, 3.3, 'Files',
  'Framework v1.1.2 base (ARA, 18 Aug 2026)\n\nMBT run: Runway 18-36 reconstruction, HMA vs PCC with lost revenue\n\nSRB run: runway reconstruction, PCC vs HMA with lost revenue\n\nNS user-manual review and annotated screenshots (24 Aug 2026)');
card(s, 3.55, 1.55, 2.9, 3.3, 'Scope',
  'Every formula in the five alternative templates and the General Information sheet\n\nThe VBA that builds alternatives and the Summary\n\nWhat the reviewer experienced on a locked-down TDOT machine\n\nWhether the results sheets can support a pavement-type decision');
card(s, 6.6, 1.55, 2.9, 3.3, 'Method',
  'Fixes applied inside the workbook XML so the VBA project, 192 ActiveX comboboxes, charts and validations are untouched\n\nEvery result reproduced by an independent engine\n\nFull recalculation of every workbook with LibreOffice Calc before release');
s.addNotes('The MBT and SRB files are older builds than the base file; they still carry a visible Indirect Cost Items sheet.');

// ---------------------------------------------------------------- 3 reviewer experience
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'What the reviewer hit first');
sub(s, 'Wes Mittlesteadt (Neel-Schaffer) never reached the calculations');
img(s, 'before_blocked.png', 0.5, 1.55, 4.9, 1.86);
caption(s, 'TDOT IT blocks ActiveX outright. There is no "Enable Content" prompt, so the pay-item dropdowns and the Alternative Setup button are dead on arrival.', 0.5, 4.25, 4.9);
bullets(s, [
  'ActiveX blocked with no prompt; instructions assumed one would appear',
  'A new pay item (prime coat) typed at the bottom of Pay_Items never showed in the dropdown',
  'RevenueData is hidden and the manual never says how to unhide a sheet',
  'Choosing an airport outside the 17 with revenue data produced #N/A in the alternative and the Summary',
  'Cells D33:D37 use gray text, not gray fill, so users cannot tell what is editable',
  'FAA rate rule changed in 2022 (OMB A-94 real rate, 2.0% for 2026, not the old 7%); the tool defaults to 3 percent and 30 years',
], 5.7, 1.55, 3.8, 3.4, 12);
s.addNotes('Every item here is from the NS review docs. The last one is a real policy conflict for AIP-funded projects.');

// ---------------------------------------------------------------- 4 calculation defects
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'Calculation defects in v1.1.2');
sub(s, 'Found in the audit, not in the review. Each one changes the answer.');
table(s, [
  ['Defect', 'Where', 'Effect'],
  ['Pay item 1 excluded from initial construction subtotal', 'Non-indirect HMA, PCC and Rehab templates', 'Construction cost understated by the first line item, often the largest'],
  ['Engineering percent not applied to initial construction', 'All templates', 'Only maintenance and rehab carry engineering; biases toward the construction-heavy alternative'],
  ['Analysis period only sets the salvage year', 'All templates', 'A 20-year run still counts maintenance at years 24 and 28'],
  ['Default period 10 while Overview says 30', 'General Information D33', 'Salvage discounted at year 10 unless the user notices'],
  ['Bare XLOOKUP for airport revenue', 'Indirect templates, D39', '#N/A for any airport outside the 17; #NAME? on Excel 2019'],
  ['No blank guard on PCC pay-item rows 14 to 22', 'PCC templates', '#N/A in Item Cost and Total once a row is left empty'],
  ['Four pay items described "Separation Geotextile"', 'Pay_Items', 'Lookup by description always returns the first, wrong unit cost'],
  ['Airport dropdown stops at row 84 of 88', 'General Information D9', 'Four airports cannot be selected'],
  ['Closure days ship as zero', 'Indirect templates F4:F10', 'Each analyst re-invented production-rate formulas by hand'],
], 0.5, 1.5, 9.0, [2.9, 2.1, 4.0], 9.5);
s.addNotes('Two more latent bugs sit in the disabled HMA Rehab template: a year-applied cell that adds the construction year, and a salvage reference to an empty cell.');

// ---------------------------------------------------------------- 5 results sheets
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'Outputs could not support a decision');
sub(s, 'What Aeronautics saw at the end of an analysis in v1.1.2');
img(s, 'before_altchart.png', 0.5, 1.55, 3.2, 1.17);
caption(s, 'Alternative chart: positions 1 to 31 instead of years, discounted and undiscounted bars on one axis, and a title that says "Alternative 4" on the Alt 1 sheet because VBA sets it once.', 0.5, 4.35, 3.2);
img(s, 'before_summary.png', 4.1, 1.55, 5.4, 3.08);
caption(s, 'Summary: two bars per alternative on a raw-number axis, #N/A for any airport outside the revenue list. Nothing about where the difference comes from, whether it holds at 7 percent, or how many days the runway closes.', 4.1, 3.4, 5.4);
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.1, y: 4.15, w: 5.4, h: 0.75, fill: { color: LIGHT }, line: { color: LIGHT }, rectRadius: 0.06 });
s.addText('The whole reason for the 2026 update, runway closures and lost revenue, appears nowhere in the outputs.', { x: 4.25, y: 4.2, w: 5.1, h: 0.65, fontFace: HF, fontSize: 12.5, bold: true, color: RED, isTextBox: true, margin: 0, valign: 'middle' });
s.addNotes('Screenshots from the NS review package.');

// ---------------------------------------------------------------- 6 policy issues
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'Policy questions the tool was hiding');
sub(s, 'Not bugs. Assumptions that decide the answer and were invisible in the outputs.');
const pol = [
  ['$3.9M vs $0.4M', 'SRB salvage credit, PCC vs HMA. PCC recovers 25% of total initial cost; HMA 12.5% of one mill-and-overlay. That single line is larger than the NPW gap.', ORANGE],
  ['Which FAA rate?', 'PGL 22-01 (2022) replaced the fixed 7% with the OMB A-94 real rate, 2.0% for 2026. TDOT uses 3% over 30 years. MBT changes winner at 4.5%, so the rule decides the answer.', INK],
  ['100% of gross', 'Lost revenue counts gross fuel sales, hangar rent and tenant revenue as lost on every closure day. Margin-based revenue would be far lower.', INK],
  ['17 of 79 airports', 'RevenueData covers 17 airports. MBT and MQY rows are identical; XNX and M54 hold text where numbers belong.', INK],
];
pol.forEach((p, i) => {
  const x = 0.5 + (i % 2) * 4.6, yy = 1.45 + Math.floor(i / 2) * 1.3;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: yy, w: 4.4, h: 1.15, fill: { color: LIGHT }, line: { color: LIGHT }, rectRadius: 0.06 });
  s.addText(p[0], { x: x + 0.15, y: yy + 0.08, w: 4.1, h: 0.45, fontFace: HF, fontSize: 22, bold: true, color: p[2], isTextBox: true, margin: 0 });
  s.addText(p[1], { x: x + 0.15, y: yy + 0.52, w: 4.1, h: 0.6, fontFace: BF, fontSize: 10.5, color: INK, isTextBox: true, margin: 0, valign: 'top' });
});
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.15, w: 9.0, h: 0.8, fill: { color: LIGHT }, line: { color: LIGHT }, rectRadius: 0.06 });
s.addText('v1.2.0 does not change any of these. It makes each one visible and testable so Aeronautics can decide them deliberately.', { x: 0.7, y: 4.2, w: 8.6, h: 0.7, fontFace: BF, fontSize: 13, color: INK, isTextBox: true, margin: 0, valign: 'middle' });
s.addNotes('Recommend a per-category "share lost during closure" factor on RevenueData and a funding-source switch for the AIP rule.');

// ---------------------------------------------------------------- 6b root cause, systemic
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'Root cause: five copies of one template');
sub(s, 'Why the defects exist, and why they sit in some templates and not others');
card(s, 0.5, 1.5, 4.4, 3.5, 'How the workbook is built',
  'Each pavement type has a hidden template sheet of about 1,600 near-identical formulas. VBA copies the template to create an alternative.\n\nThe 2026 lost-revenue feature was added by cloning two templates and editing the clones. There is no shared formula source, so a fix made in one copy never reaches the others.\n\nNew inputs (analysis period, revenue lookup) were wired to one cell each rather than through the whole template.');
card(s, 5.1, 1.5, 4.4, 3.5, 'What that produced',
  'Pay item 1 dropped in three templates, correct in the two 2026 clones\n\nBlank guard on row 13 in every template, missing on rows 14 to 22 of the PCC templates only\n\nAnalysis period wired to the salvage row and nothing else\n\nRevenue guard present on General Information, lost when the lookup was copied into the templates\n\nChart helper rows placed by hand at row 37 plus the policy year, with the category reference removed');
s.addNotes('Recommendation for the next revision: one template per pavement type with the lost-revenue option as a switch, plus a consistency check macro across templates.');

// ---------------------------------------------------------------- 6c root cause, top four
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'The four defects that move money');
table(s, [
  ['Defect', 'Root cause (evidence in the file)', 'Exact change'],
  ['Pay item 1 excluded', 'Subtotal G24 = SUM(G14:G23) while items occupy rows 13 to 22 (A13 = 1). The indirect clones read SUM(G13:G22), so the range was corrected there and never propagated.', 'G24 = SUM(G13:G22) in TMP(NewHMA), TMP(NewPCC), TMP(HMARehab)'],
  ['No engineering on initial cost', 'Row 26 is formatted and included in the Total (G27 = SUM(G24:G26)) but empty. The M&R blocks all carry Mobilization then Engineering; the initial block stopped after Mobilization.', 'B26 = "Engineering"; G26 = (D37/100) x G24 in all five templates'],
  ['Analysis period ignored', 'D33 is referenced once per template, by the salvage row. Every activity year comes straight from Maintenance Policies with no comparison to D33.', 'Each discounted cost: IF(year > D33, 0, cost / (1+r)^year); default D33 = 30'],
  ['#N/A on missing airport', 'General Information D39 guards the lookup with COUNTIF; the templates copied only the inner SUM(XLOOKUP(...)). XLOOKUP also fails on Excel 2019 (#NAME?).', 'F2 = IF(D38="Yes", SUMIF chain over RevenueData columns B:H, 0); G2 warning text when F2 = 0'],
], 0.5, 1.3, 9.0, [1.8, 4.2, 3.0], 9.5);
s.addNotes('Full cell-by-cell record, 463 cells, is Appendix A of the Technical Change Record.');

// ---------------------------------------------------------------- 7 fixes: calculations
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'v1.2.0: calculation fixes');
sub(s, 'Same workbook, same workflow, same VBA. Only the arithmetic changed.');
table(s, [
  ['Item', 'v1.1.2', 'v1.2.0'],
  ['Initial construction subtotal', 'SUM(G14:G23)', 'SUM(G13:G22), pay item 1 included'],
  ['Engineering on initial construction', 'not applied', 'Engineering row added, same percent as M&R'],
  ['Analysis period', 'salvage year only', 'Activities beyond the period drop to $0; default 30'],
  ['Airport revenue lookup', 'XLOOKUP, #N/A if missing', 'SUMIF chain, $0 plus a visible warning when the airport has no data'],
  ['PCC pay-item rows 14 to 22', 'no blank guard', 'guarded like row 13'],
  ['Geotextile pay items', 'four identical descriptions', 'suffixed with the spec number'],
  ['Closure durations', 'zero', 'production-rate defaults from the MBT and SRB runs, still editable'],
  ['Airport dropdown', 'rows 9 to 84', 'rows 10 to 88'],
  ['HMA Rehab template', 'two latent bugs', 'fixed in case the option is re-enabled'],
], 0.5, 1.5, 9.0, [2.6, 2.4, 4.0], 10);
s.addNotes('All edits made in the sheet XML. openpyxl or a re-save would have stripped the ActiveX controls and charts.');

// ---------------------------------------------------------------- 8 fixes: usability
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'v1.2.0: usability and charts');
table(s, [
  ['Area', 'What changed'],
  ['Instructions', 'ActiveX "blocked content" steps (Trust Center, then restart Excel, or ask IT); correct button name; D9:D39; how to unhide RevenueData and the 17-airport rule'],
  ['Alternative chart', 'Year-indexed data keyed on Year Applied; calendar years on the axis; undiscounted direct and lost revenue stacked; mixed discounted series removed'],
  ['Project identity and map', 'The Summary names the project and shows where it is: a locator map of Tennessee drawn from embedded coordinates, no internet, no add-in'],
  ['Google Earth export', 'A VBA module writes a KML: the result in the airport bubble, a runway footprint and a present-worth bar per alternative, and every event stamped with its year for the time slider'],
  ['Chart axes', 'Every chart in both workbooks now names its x and y axis; legends moved beside the plots so they no longer collide with the category labels'],
  ['Summary layout', 'The duplicate table in columns A to F is hidden (the form still writes it), so the results table and the charts start at the left edge'],
  ['Summary sheet', 'Formula-driven, no macro: results table with PW by category, delta to lowest, closure days and runway availability; a RealCost-style comparison block (agency and user cost as PW and EUAC); five charts; navigation buttons that work with macros blocked'],
  ['Pavement section', 'New on the Summary: each layer\'s thickness read back from the quantities, the section as a string for the description, a cross-check against the excavation quantity, and two stacked charts (mainline, and mainline plus shoulder when entered)'],
  ['Summary as a dashboard', 'Six tiles above the results: lowest present worth and the section it buys, margin to the next alternative (amber under five percent), equivalent annual cost, initial construction, unit cost per S.Y., and whether the winner holds at every rate tested. Data bars in the present-worth column; a new chart 8 puts all four unit costs against the published Tennessee range'],
  ['Setup flow', 'The card on General Information lists five steps and names the sheet for each; every sheet in the sequence says which step it is; Pay_Items and Maintenance Policies gained the navigation row and their own buttons'],
  ['Pay_Items and Maintenance Policies', 'The two sheets that had never been designed: frozen and repeating headers, the Unit Cost column marked as the input, part headings banded, and notes saying the division columns are empty and 29 of 56 items carry no cost. Maintenance Policies says which table drives which alternative and what Rate means'],
  ['Look and first-run UX', 'How-to card and a live "still needed" status line on General Information; navigation row on every alternative sheet; lowest-cost row highlighted on the Summary; tab colors; General Information prints on one page'],
  ['Method sheet', 'New reference sheet: every calculation once, with the formula, the rule in plain English and its source; a live block reading the current rate, period and capital recovery factor; and the assumptions a reviewer will ask about'],
  ['Typical Values sheet', 'New reference sheet: what TDOT manages, typical runway geometry and areas, workbook unit costs beside 2025 bids, closure production rates, FAA and OMB discount-rate rules, service lives, live daily revenue for the 17 airports; hints beside the General Information inputs'],
  ['Housekeeping', 'Personal names removed from cell notes and document properties; dead links, cached printer settings, Power Query stub and SharePoint path removed; sources on Typical Values as plain text; full recalculation on open'],
  ['Not changed', 'Overview wording (NS rewrite pending Mat\'s approval); salvage, lost-revenue and AIP policy'],
], 0.5, 1.3, 9.0, [2.0, 7.0], 10.5);
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.3, w: 9.0, h: 0.65, fill: { color: LIGHT }, line: { color: LIGHT }, rectRadius: 0.06 });
s.addText('Nothing to install: the workbook is self-contained. The VBA project is untouched; Alternative Setup still writes columns A:E of the Summary as before.', { x: 0.7, y: 4.35, w: 8.6, h: 0.55, fontFace: BF, fontSize: 12.5, color: INK, isTextBox: true, margin: 0, valign: 'middle' });
s.addNotes('The Summary formulas read the Database sheet the form writes, so they follow whatever alternatives exist.');

// ---------------------------------------------------------------- 8a the front page
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'The page users start on');
sub(s, 'Same form, same grey cells. What is new is the guidance around them.');
let y = img(s, 'general_information.png', 2.6, 1.35, 4.8, 1.399);
caption(s, 'A how-to card beside the logo, a status line that names whatever input is still empty and turns green when the form is complete, banded section headings, and buttons to the Summary and the Typical Values sheet.', 0.5, y + 0.05, 9.0);

// ---------------------------------------------------------------- 8b the new Summary sheet
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'The new Summary sheet');
sub(s, 'MBT data loaded. Table, verdict line, RealCost-style comparison block, five charts, all live formulas.');
y = img(s, 'summary_sheet.png', 1.1, 1.42, 7.8, 1.987);
caption(s, 'The comparison block under the verdict line follows the RealCost layout: agency and user cost, each as present worth and EUAC.', 0.5, y + 0.02, 9.0);

// ---------------------------------------------------------------- 8c typical values
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'Typical values for the inputs');
sub(s, 'A reference sheet answering "is this number reasonable" while filling in General Information.');
y = img(s, 'typical_values.png', 0.5, 1.5, 9.0, 4.737);
caption(s, 'Section 3: each Pay_Items default read live from the workbook, beside recent bid prices and what they mean. Six more sections cover system context, runway geometry, closure rates, discount rates and service lives.', 0.5, y + 0.05, 9.0);
bullets(s, [
  'TDOT Aeronautics manages 78 public-use airports; 69 are in the NPIAS and about 70 are in the pavement management network',
  'Typical TN general aviation runway: 4,000 to 6,000 ft by 75 or 100 ft, which is 33,000 to 67,000 SY of mainline',
  'The $130 per ton asphalt default sits below 2025 southeastern bids of $185 to $232',
  'Discount rate: TDOT 3 percent; FAA now points to the OMB A-94 real rate, 2.0 percent for 2026, not the old 7 percent',
], 0.5, y + 0.55, 9.0, 0.9, 11.5);
s.addNotes('Every row has a source; the research note in the repo says which figures still need the primary PDF.');

// ---------------------------------------------------------------- 9 decision workbook intro
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'The decision workbook');
sub(s, 'MBT and SRB already loaded, plus a blank TEMPLATE sheet. Yellow cells drive every table and chart.');
y = img(s, 'MBT_kpis.png', 0.5, 1.5, 9.0, 3.97);
caption(s, 'The verdict line: which alternative is lower, by how much, and two robustness flags a reviewer can read in five seconds. Shown here for MBT at TDOT settings.', 0.5, y + 0.1, 9.0);
bullets(s, [
  'RESULTS: present worth by category, EUAC, cost per SY, closure days, runway availability',
  'BREAK-EVEN VALUES: the daily revenue, bid level, salvage fraction and discount rate at which the two tie',
  'Twelve pre-run scenarios, a rate-by-salvage decision map, a tornado, and a live 1,000-draw simulation',
  'A TEMPLATE sheet: copy it, paste the next project\'s costs and closure days, done',
], 0.5, 4.3, 9.0, 1.0, 11.5);
s.addNotes('Everything is a formula. Change the discount rate cell and the sensitivity, tornado, map and charts all move.');

// ---------------------------------------------------------------- 10 bridge + closures
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'Where the difference comes from');
y = img(s, 'MBT_bridge.png', 0.5, 1.2, 9.0, 4.14);
caption(s, 'MBT bridge from HMA to PCC present worth. PCC costs $3.1M more to build and earns it back through $1.9M less maintenance and $1.1M less rehabilitation; salvage and lost revenue are small here.', 0.5, y + 0.05, 9.0);
y = img(s, 'MBT_closures.png', 0.5, y + 0.55, 9.0, 6.44);
caption(s, 'Closure timeline, bubble size is days closed. HMA closes the runway seven times for 57 days; PCC twice for 27 days.', 0.5, y + 0.05, 9.0);
s.addNotes('This is the chart Aeronautics asked for when they commissioned the lost-revenue update.');

// ---------------------------------------------------------------- 11 robustness
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'Will the answer hold?');
sub(s, 'Net present worth against discount rate. Blue is HMA, orange is PCC. TDOT 3 percent and FAA AIP 7 percent are marked.');
y = img(s, 'MBT_sens.png', 0.5, 1.5, 4.3, 2.07);
caption(s, 'MBT: PCC wins below 4.5 percent, HMA above. At the FAA AIP rate the answer reverses.', 0.5, y + 0.05, 4.3);
y = img(s, 'SRB_sens.png', 5.2, 1.5, 4.3, 2.07);
caption(s, 'SRB: HMA wins at every rate above 2.5 percent. The gap widens as the rate rises.', 5.2, y + 0.05, 4.3);
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.15, w: 9.0, h: 0.85, fill: { color: LIGHT }, line: { color: LIGHT }, rectRadius: 0.06 });
s.addText([
  { text: 'Live simulation: ', options: { bold: true } },
  { text: '1,000 joint draws on rate, bids, closures, salvage and revenue share. Probability PCC is the cheaper alternative: 52 percent at MBT, 18 percent at SRB. Press F9 in the workbook to redraw.' },
], { x: 0.7, y: 4.2, w: 8.6, h: 0.75, fontFace: BF, fontSize: 12, color: INK, isTextBox: true, margin: 0, valign: 'middle' });
s.addNotes('Break-even table on the sheet gives the exact tie points: PCC bids +9.9% or salvage below 2.5% of initial cost flips MBT.');

// ---------------------------------------------------------------- 11b tornado
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'What could flip the answer');
sub(s, 'MBT. PCC minus HMA present worth when one input moves across its plausible range. Bars crossing zero change the winner.');
y = img(s, 'MBT_tornado.png', 0.5, 1.5, 9.0, 3.91);
caption(s, 'Bid prices and the discount rate can flip MBT; closure durations, lost revenue, the analysis period and HMA salvage cannot. The same view for SRB shows only bids and the discount rate matter.', 0.5, y + 0.1, 9.0);
table(s, [
  ['Break-even, MBT', 'Value', 'Current'],
  ['PCC bid level that ties the two', '+9.9%', 'base estimate'],
  ['PCC salvage credit that ties the two', '2.5% of initial', '25%'],
  ['Discount rate at which the winner changes', '4.5%', '3%'],
], 0.5, 4.35, 9.0, [4.6, 2.2, 2.2], 10.5);
s.addNotes('Break-even values come from the BREAK-EVEN VALUES block on the sheet, computed in closed form, no goal seek.');

// ---------------------------------------------------------------- 12 results
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'What the two runs say');
sub(s, 'TDOT settings: 3 percent, 30 years, lost revenue on, policy salvage');
table(s, [
  ['', 'MBT (Murfreesboro)', 'SRB (Upper Cumberland)'],
  ['Lower present worth', 'PCC by $781k (8.9%)', 'HMA by $586k (4.1%)'],
  ['Holds at FAA 7%?', 'No, flips at 4.5%', 'Yes'],
  ['Holds with zero salvage?', 'By $33k only', 'Yes'],
  ['Closure days over 30 years, HMA / PCC', '57 / 27', '75 / 31'],
  ['PCC bid change that ties the two', '+9.9%', '-4.1%'],
  ['Probability PCC is lower (simulation)', '52%', '18%'],
], 0.5, 1.5, 9.0, [3.4, 2.8, 2.8], 12);
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.95, w: 4.4, h: 1.15, fill: { color: LIGHT }, line: { color: LIGHT }, rectRadius: 0.06 });
s.addText('MBT: both margins are inside estimating noise. The answer is a discount-rate and salvage-policy call, not an engineering one.', { x: 0.65, y: 4.0, w: 4.1, h: 1.05, fontFace: BF, fontSize: 11.5, color: INK, isTextBox: true, margin: 0, valign: 'middle' });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.1, y: 3.95, w: 4.4, h: 1.15, fill: { color: LIGHT }, line: { color: LIGHT }, rectRadius: 0.06 });
s.addText('SRB: HMA wins on cost, PCC halves the closure days, and no plausible daily revenue closes the gap. The decision is what a runway-day is worth.', { x: 5.25, y: 4.0, w: 4.1, h: 1.05, fontFace: BF, fontSize: 11.5, color: INK, isTextBox: true, margin: 0, valign: 'middle' });
s.addNotes('SRB break-even daily revenue is about $23,000 per day against $4,344 actual.');

// ---------------------------------------------------------------- 13 verification
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'Verification before release');
table(s, [
  ['Check', 'Result'],
  ['Framework v1.2.0, full recalculation in LibreOffice Calc', '2,240 formulas incl. the new Summary, 0 errors (v1.1.2 shows 99 error cells under the same recalc)'],
  ['Patch applied to the populated MBT workbook', 'NPW unchanged: $8,809,266 and $8,028,734; PCC blank rows now 0, not #N/A'],
  ['New year-indexed chart columns', 'Direct plus lost revenue reconciles to activity totals; discounted column sums to NPW to the dollar'],
  ['Decision workbook, full recalculation', '47,265 formulas, 0 errors'],
  ['Decision workbook against an independent engine', 'NPW, categories, break-evens, sensitivity, tornado, 12 scenarios and the decision map agree to the dollar'],
  ['New Summary sheet on the MBT data', 'Table, verdict, categories and closure days match; six charts render; template reads "No alternatives yet" with no errors'],
  ['New project run end to end', 'CKV Runway 17-35 built on the blank template: framework, decision TEMPLATE and an independent engine agree to the dollar'],
  ['Two worked examples', 'GKT Runway 10-28 (three alternatives, no shoulders) and MKL Runway 2-20 (four alternatives, 25 ft shoulders): 104 independent checks pass; section read-back returns the designed sections'],
  ['Scripted click-through (LibreOffice API)', 'Buttons land where they should; 7% flips the verdict as flagged; 20 years drops late closures; unknown airport gives $0 lost revenue with a warning; deleting an alternative shifts the table with no #REF!; ties and four alternatives read cleanly'],
  ['Not yet verified', 'The alternative-sheet charts in Excel itself (no Excel here); a first-open check is in the notes'],
], 0.5, 1.2, 9.0, [3.4, 5.6], 10);
s.addNotes('The first-open check: pick an airport, set Yes, add one HMA and one PCC alternative, confirm Summary and Alt sheet columns L:N.');

// ---------------------------------------------------------------- 14 how TDOT uses it
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'How TDOT uses it');
sub(s, 'Excel only. No add-ins. The decision workbook has no macros at all.');
const steps = [
  ['Run the framework as before', 'Open v1.2.0, enable content, fill General Information, build alternatives with Alternative Setup. Click View Summary: categories, closures, the 7 percent flag and five charts, no import needed.'],
  ['Copy the TEMPLATE sheet', 'In the decision workbook, right-click TEMPLATE, Move or Copy, rename. Paste the base costs and closure days from each Alt sheet into the yellow cells.'],
  ['Read the decision blocks', 'RESULTS, BREAK-EVEN VALUES, the scenario scorecard and the decision map. Change any yellow cell and every chart follows. F9 redraws the simulation.'],
];
steps.forEach((st, i) => {
  const x = 0.5 + i * 3.05;
  s.addShape(pres.shapes.OVAL, { x, y: 1.6, w: 0.55, h: 0.55, fill: { color: DARK }, line: { color: DARK } });
  s.addText(String(i + 1), { x, y: 1.6, w: 0.55, h: 0.55, fontFace: HF, fontSize: 18, bold: true, color: WHITE, align: 'center', valign: 'middle', isTextBox: true, margin: 0 });
  s.addText(st[0], { x, y: 2.3, w: 2.8, h: 0.5, fontFace: HF, fontSize: 15, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText(st[1], { x, y: 2.85, w: 2.8, h: 1.8, fontFace: BF, fontSize: 12, color: INK, isTextBox: true, margin: 0, valign: 'top' });
});
s.addNotes('Two alternatives per sheet, HMA against PCC, is how the framework is used; policy years are editable if TDOT revises the maintenance tables.');

// ---------------------------------------------------------------- 15 open items
s = pres.addSlide(); s.background = { color: WHITE };
title(s, 'Open items');
card(s, 0.5, 1.4, 4.4, 3.6, 'Decisions for Aeronautics',
  'Salvage policy: keep 25% of initial for PCC and 12.5% of one overlay for HMA, or move both to a remaining-life basis\n\nLost revenue basis: gross receipts or a per-category share lost during closure\n\nFAA AIP projects: enforce 7% and 20 years, or show both\n\nRevenueData: confirm MBT vs MQY, fix XNX and M54, add the other 62 airports or state the rule\n\nOverview text: adopt the NS rewrite');
card(s, 5.1, 1.4, 4.4, 3.6, 'Next steps',
  'First open in Excel: run the first-open check in the notes\n\nSend v1.2.0 and the decision workbook to NS for a second pass\n\nPCC rehabilitation template is still "Coming soon" if TDOT wants rehab alternatives\n\nOptional later: a no-macro web version would remove the ActiveX problem entirely');
s.addNotes('Keep it Excel-first; the web option is noted only because IT blocking was the reviewer\'s first obstacle.');

// ---------------------------------------------------------------- 16 close
s = pres.addSlide(); s.background = { color: DARK };
s.addText('In one line', { x: 0.6, y: 1.2, w: 8.8, h: 0.6, fontFace: BF, fontSize: 14, color: 'AEB8C4', isTextBox: true, margin: 0, italic: true });
s.addText('The tool now computes what it says it computes, shows why one alternative wins, and shows how many runway days that costs.', { x: 0.6, y: 1.8, w: 8.8, h: 1.6, fontFace: HF, fontSize: 26, bold: true, color: WHITE, isTextBox: true, margin: 0, valign: 'top' });
s.addText('Deliverables: TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm  |  TDOT_LCCA_Decision_Workbook.xlsx  |  Technical Change Record  |  change notes', { x: 0.6, y: 4.4, w: 8.8, h: 0.6, fontFace: BF, fontSize: 11, color: 'AEB8C4', isTextBox: true, margin: 0 });

pres.writeFile({ fileName: __dirname + '/TDOT_LCCA_Review_and_v1.2.0_Update.pptx' }).then(f => console.log('wrote', f));
