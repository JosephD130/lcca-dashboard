const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, HeadingLevel, AlignmentType, ImageRun,
  PageBreak, TableOfContents, ShadingType, BorderStyle, LevelFormat, Footer, Header, PageNumber, TabStopType } = require('docx');

const S = path.resolve(__dirname, '..');
const diff = JSON.parse(fs.readFileSync(S + '/diff.json', 'utf8'));
const SUMMARY_FORMULAS = JSON.parse(fs.readFileSync(S + '/summary_formulas.json', 'utf8'));
const IMG = __dirname + '/img/';

// ------------------------------------------------------------------ helpers
const FONT = 'Calibri', MONO = 'Courier New';
const W = 9360; // 6.5 in text width in DXA
function p(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({ text, font: FONT, size: opts.size || 22, bold: opts.bold, italics: opts.italics, color: opts.color })];
  return new Paragraph({ children: runs, spacing: { after: opts.after == null ? 120 : opts.after, before: opts.before || 0 }, alignment: opts.align, keepNext: opts.keepNext });
}
function r(text, o = {}) { return new TextRun({ text, font: o.mono ? MONO : FONT, size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color }); }
function h1(t) { return new Paragraph({ text: t, heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 } }); }
function h2(t) { return new Paragraph({ text: t, heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 120 } }); }
function h3(t) { return new Paragraph({ text: t, heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 80 } }); }
function bullet(t, level = 0) { return new Paragraph({ children: [r(t)], numbering: { reference: 'bul', level }, spacing: { after: 60 } }); }
function num(t) { return new Paragraph({ children: [r(t)], numbering: { reference: 'num', level: 0 }, spacing: { after: 60 } }); }
function mono(t, size = 16) { return new Paragraph({ children: [new TextRun({ text: t, font: MONO, size })], spacing: { after: 0 } }); }
function cell(text, w, o = {}) {
  const lines = String(text == null ? '' : text).split('\n');
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: o.head ? { type: ShadingType.CLEAR, fill: 'E7E6E6', color: 'auto' } : (o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined),
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: lines.map(l => new Paragraph({ children: [new TextRun({ text: l, font: o.mono ? MONO : FONT, size: o.size || 18, bold: o.head })], spacing: { after: 0 } })),
  });
}
function table(rows, widths, o = {}) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((row, i) => new TableRow({ tableHeader: i === 0, cantSplit: true, children: row.map((c, j) => cell(c, widths[j], { head: i === 0, mono: o.monoCols && o.monoCols.includes(j) && i > 0, size: o.size })) })),
  });
}
function image(file, widthIn, ratio) {
  const w = Math.round(widthIn * 96), h = Math.round(w / ratio);
  return new Paragraph({ children: [new ImageRun({ type: 'png', data: fs.readFileSync(IMG + file), transformation: { width: w, height: h } })], spacing: { after: 60 }, alignment: AlignmentType.CENTER });
}
function caption(t) { return p([r(t, { italics: true, size: 18, color: '595959' })], { after: 200 }); }
const GI = "'General Information'";

// ------------------------------------------------------------------ Appendix A compression
function groupSheet(name, entries) {
  const rows = [];
  const byCol = {};
  for (const e of entries) {
    const m = e.cell.match(/^([A-Z]+)(\d+)$/); const col = m[1], row = +m[2];
    if (['L', 'M', 'N'].includes(col) && row >= 37) { (byCol[col] = byCol[col] || []).push({ row, e }); }
    else rows.push([e.cell, e.before == null ? '(empty)' : e.before, e.after == null ? '(empty)' : e.after]);
  }
  for (const col of ['L', 'M', 'N']) {
    if (!byCol[col]) continue;
    const list = byCol[col].sort((a, b) => a.row - b.row);
    const first = list[0], last = list[list.length - 1];
    const befores = [...new Set(list.map(x => x.e.before == null ? '(empty)' : x.e.before))];
    rows.push([`${col}${first.row}:${col}${last.row}\n(${list.length} cells)`,
      `Hard-positioned links or blanks, e.g. ${befores.slice(0, 4).join(', ')}${befores.length > 4 ? ', ...' : ''}`,
      `Row ${first.row} shown; the row number substitutes down the column:\n${first.e.after}`]);
  }
  return rows;
}

// ------------------------------------------------------------------ content
const children = [];
const cover = [
  new Paragraph({ spacing: { before: 2400 } }),
  p([r('Tennessee Department of Transportation, Aeronautics Division', { size: 24, color: '595959' })], { after: 120 }),
  p([r('LCCA Framework Workbook', { size: 48, bold: true })], { after: 60 }),
  p([r('Technical Change Record, v1.1.2 to v1.2.0', { size: 32 })], { after: 400 }),
  p([r('Root causes, exact changes, verification, and the companion decision workbook', { size: 24, italics: true, color: '595959' })], { after: 1400 }),
  p([r('Prepared by Applied Research Associates, Inc., Transportation Division', { size: 22 })], { after: 40 }),
  p([r('Task 005935.00000.00000.TASK2', { size: 22 })], { after: 40 }),
  p([r('11 September 2026', { size: 22 })], { after: 40 }),
  p([r('Companion to: TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm and TDOT_LCCA_Decision_Workbook.xlsx', { size: 20, color: '595959' })], { after: 40 }),
  new Paragraph({ children: [new PageBreak()] }),
  h1('Contents'),
  new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-2' }),
  new Paragraph({ children: [new PageBreak()] }),
];
children.push(...cover);

// 1 purpose
children.push(h1('1. Purpose and scope'));
children.push(p('This record documents every change made to the TDOT Aeronautics Life-Cycle Cost Analysis (LCCA) framework workbook between version 1.1.2 (ARA, 18 August 2026) and version 1.2.0 (ARA, 11 September 2026). For each defect it states the symptom, the evidence in the file, the root cause, the exact cell-level change, the effect on results, and how the change was verified. It also documents the new formula-driven Summary sheet, the changes to the Instructions text and to the chart parts, and the companion decision workbook built for the Murfreesboro (MBT) and Upper Cumberland (SRB) analyzes.'));
children.push(p('The review was triggered by the Neel-Schaffer (NS) review of 24 August 2026 (user manual comments and annotated screenshots) and by Aeronautics\' request that ARA spend 8 to 16 hours reviewing and improving the tool. The NS comments are usability findings; the calculation defects below were found in ARA\'s line-by-line audit of the templates and VBA.'));
children.push(p('Scope of the v1.2.0 change: arithmetic, guards, defaults, chart data, instructions and the Summary builder. Out of scope, and unchanged: the TDOT Maintenance Policies, the pay-item database values, the salvage policy, the lost-revenue method and the alternative-creation workflow. Section 11 lists the policy questions that the update makes visible but deliberately does not decide.'));

// 2 files
children.push(h1('2. Files, versions and integrity'));
children.push(p('SHA-256 hashes of the inputs reviewed and the outputs delivered. A recipient can confirm they hold the same bytes with certutil -hashfile <file> SHA256 on Windows.'));
const hashRows = [['File', 'Role', 'SHA-256']];
const roles = {
  '8eaa2547-TDOA_LCCA_Framework_v1.1.2_ARA_Task2_08182026.xlsm': 'Input: v1.1.2 base workbook (reviewed)',
  '35f08cad-TDOA_LCCA_Framework_v1.1.2_MBT_20260521_LostRevenue_1.xlsm': 'Input: MBT run (verification data)',
  '527a5482-TDOA_LCCA_Framework_v1.1.2_SRB_20260521_LostRevenue.xlsm': 'Input: SRB run (verification data)',
  'TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm': 'Output: v1.2.0 workbook',
  'Output.bas': 'Superseded first draft (not a deliverable)',
  'TDOT_LCCA_Decision_Workbook.xlsx': 'Output: decision workbook',
};
for (const [f, h] of Object.entries(diff.hashes)) hashRows.push([f.replace(/^[0-9a-f]{8}-/, ''), roles[f] || '', h]);
children.push(table(hashRows, [3200, 2400, 3760], { monoCols: [2], size: 14 }));
children.push(p(''));
children.push(p('Version identity of the workbook: the sheet list, the VBA project (byte for byte), ActiveX controls, tables, data validations, comments and drawing parts of v1.1.2 are carried into v1.2.0 unchanged except where this record says otherwise. The two external-link parts and the calculation chain part were removed; Excel rebuilds the calculation chain on first open.'));

// 3 method
children.push(h1('3. How the review and the change were done'));
children.push(bullet('Every worksheet, defined name, data validation, chart part, drawing text box, comment and the full VBA project of v1.1.2 were extracted and read. The five alternative templates (TMP(NewHMA), TMP(NewHMA)_IndirectCost, TMP(NewPCC), TMP(NewPCC)_IndirectCost, TMP(HMARehab)) were compared formula by formula against each other and against the populated MBT and SRB alternative sheets.'));
bullet('');
children.push(bullet('Fixes were applied by editing the worksheet, chart, drawing and shared-string XML inside the .xlsm package directly. Re-saving the workbook through a library or another application would have dropped the 192 ActiveX pay-item comboboxes and the embedded charts, so no such tool touched the file. The patch is a script (patch_lcca.py) and is reproducible.'));
children.push(bullet('The patched workbook and a patched copy of the MBT workbook were fully recalculated with LibreOffice Calc (calculateAll, not cached values) and every cell scanned for error values. Results were compared with an independent Python implementation of the LCCA arithmetic.'));
children.push(bullet('Charts were validated as well-formed XML and parsed with an independent chart reader to confirm their category and value ranges. The decision workbook was recalculated the same way and rendered to PDF to inspect every chart.'));
children.push(bullet('Not done here: the alternative-sheet charts have not been displayed in Microsoft Excel, which was not available in the review environment (the Summary charts were rendered). Section 12 gives the first-open checklist that closes that gap.'));

// 4 systemic root cause
children.push(h1('4. Systemic root cause'));
children.push(p('The workbook is built from hidden template sheets, one per pavement type, that VBA copies to create each alternative. Each template holds about 1,600 formulas laid out in eight side-by-side blocks (initial construction, six maintenance activities, one rehabilitation) plus a net-present-worth table and a chart helper block. The templates are near-identical copies of one another and are maintained by hand.'));
children.push(p('The 2026 lost-revenue feature was implemented by cloning TMP(NewHMA) and TMP(NewPCC) into TMP(NewHMA)_IndirectCost and TMP(NewPCC)_IndirectCost and editing the clones. There is no single formula source, no consistency check across templates, and the new inputs added over time (analysis period, lost-revenue lookup) were wired to one cell each rather than through the whole calculation. This is the single cause behind most of the defects in Section 5:'));
children.push(bullet('The initial-construction subtotal excludes pay item 1 in the three original templates and is correct in the two 2026 clones: a fix made in the clones that never reached the originals.'));
children.push(bullet('The blank guard on the pay-item lookup exists on row 13 of every template and on rows 14 to 22 of the HMA templates, but is missing from rows 14 to 22 of both PCC templates: a fix applied to one pavement type only.'));
children.push(bullet('The analysis period cell is referenced by exactly one formula per template, the salvage row.'));
children.push(bullet('The revenue lookup on General Information carries a COUNTIF guard; the copy pasted into the templates carries only the inner lookup.'));
children.push(bullet('The chart helper columns were populated by placing links at row 37 plus the policy year by hand, with the category reference later removed (a filteredCategoryTitle extension remains in the chart XML as evidence that it once existed).'));
children.push(p('Recommendation for the next revision: keep one template per pavement type with the lost-revenue option as a switch on the sheet, and add a consistency-check macro that compares the formula text of the templates block by block. That removes the class of defect rather than the instances.'));

// 5 findings
children.push(h1('5. Findings, root causes and changes'));
children.push(p('Each finding below gives the symptom a user would see, the evidence in v1.1.2, the root cause, the exact change in v1.2.0 and its effect. Cell references are as they appear in the template sheets; the alternative sheets created from them carry the same references. Appendix A lists every changed cell.'));

function finding(id, title, sym, evid, cause, changes, effect, verif) {
  children.push(h2(`${id}. ${title}`));
  children.push(p([r('Symptom. ', { bold: true }), r(sym)]));
  children.push(p([r('Evidence in v1.1.2. ', { bold: true }), r(evid)]));
  children.push(p([r('Root cause. ', { bold: true }), r(cause)]));
  children.push(p([r('Change in v1.2.0.', { bold: true })], { keepNext: true }));
  children.push(table([['Sheet / cell', 'v1.1.2', 'v1.2.0'], ...changes], [1900, 3000, 4460], { monoCols: [1, 2], size: 15 }));
  children.push(p(''));
  children.push(p([r('Effect on results. ', { bold: true }), r(effect)]));
  children.push(p([r('Verification. ', { bold: true }), r(verif)]));
}
const F2new = `=IF(${GI}!$D$38="Yes",SUMIF(RevenueData!$A:$A,${GI}!$D$10,RevenueData!$B:$B)+SUMIF(...$C:$C)+SUMIF(...$D:$D)+SUMIF(...$E:$E)+SUMIF(...$F:$F)+SUMIF(...$G:$G)+SUMIF(...$H:$H),0)`;

finding('5.1', 'Pay item 1 excluded from the initial construction subtotal',
  'The first pay item entered on an alternative sheet does not count toward Initial Construction, Mobilization, Total, or NPW when the analysis is run without lost revenue.',
  'TMP(NewHMA)!G24, TMP(NewPCC)!G24 and TMP(HMARehab)!G24 read =SUM(G14:G23). Pay items occupy rows 13 to 22 (A13 = 1 ... A22 = 10; row 23 is empty). The two 2026 templates read =SUM(G13:G22).',
  'The subtotal range was authored for a layout in which items began on row 14 and was not updated when row 13 became the first item; the clones built in 2026 corrected it locally. Because templates are maintained as separate copies, the correction never propagated back.',
  [['TMP(NewHMA)!G24\nTMP(NewPCC)!G24\nTMP(HMARehab)!G24', '=SUM(G14:G23)', '=SUM(G13:G22)']],
  'Any v1.1.2 analysis run with "Account Indirect Cost" = No understates initial construction by the item-1 line and by 10 percent mobilization on it. In the SRB layout item 1 is the largest line (P-403 asphalt base, $1.95M). The MBT and SRB reports were run with lost revenue on and are not affected.',
  'Full recalculation shows the three templates now sum rows 13 to 22; the two clones are unchanged.');

finding('5.2', 'Engineering percent not applied to initial construction',
  'General Information D37 (Engineering, 5 percent) is added to every maintenance and rehabilitation activity but not to initial construction, so the comparison is biased toward the alternative with the higher first cost.',
  'In every template the initial block ends with row 24 Subtotal, row 25 Mobilization, row 26 empty (formatted, no label, no formula), row 27 Total =SUM(G24:G26). Every M&R block carries Subtotal, Mobilization, Engineering, Total. Row 26 was reserved for Engineering and left empty.',
  'An incomplete feature: the Total formula already spans row 26, so the intent is clear; the row was never filled in and no template check would have caught it.',
  [['All five templates, B26', '(empty)', 'Engineering'], ['All five templates, G26', '(empty)', `=(${GI}!$D$37/100)*G24`]],
  'Initial construction rises by the engineering percent of the pay-item subtotal. Re-running the reports in v1.2.0 at the same inputs: MBT HMA initial $5,306,577 to $5,547,785 (+$241,208), MBT PCC $8,410,245 to $8,792,529 (+$382,284), SRB HMA $9,009,181 to $9,418,689 (+$409,508), SRB PCC $15,445,398 to $16,147,462 (+$702,064). Because PCC salvage is 25 percent of initial cost, PCC NPW rises by less than its initial cost: MBT NPW moves from $8,809,266 / $8,028,734 to $9,050,474 / $8,371,644 (PCC still lower, margin $679k instead of $781k); SRB moves from $13,854,192 / $14,439,889 to $14,263,700 / $15,069,642 (HMA still lower, margin $806k instead of $586k). Aeronautics should decide whether engineering belongs on initial construction; if not, set D37 to zero for that block or delete G26.',
  'Recalculated values of G26 and G27 in every template; the arithmetic above was computed with the independent engine.');

finding('5.3', 'Analysis period only sets the salvage year',
  'Changing General Information D33 (Analysis Period) moves the salvage credit to that year but every maintenance and rehabilitation activity is still counted, including those scheduled after the period ends. The NS reviewer set 20 years and got salvage at year 20 with maintenance at years 24 and 28 still in the NPW.',
  'D33 is referenced by exactly one cell per template (the salvage-row year, e.g. TMP(NewHMA)!C44). Activity years come straight from Maintenance Policies (E10, E12, E14 ...). The discounted-cost formula =D37/(1+(D34/100))^C37 has no test against D33.',
  'The analysis period was added as an input after the templates were built and wired only to the salvage row. The Overview text states a fixed 30-year period, which is what the templates assume.',
  [[`Every discounted-cost cell, e.g. TMP(NewHMA)!E37:E44`, `=D37/(1+(${GI}!$D$34/100))^C37`, `=IF(C37>${GI}!$D$33,0,D37/(1+(${GI}!$D$34/100))^C37)`],
   ['General Information!D33', '10', '30']],
  'At 30 years nothing changes (all policy years are 30 or less), so the MBT and SRB results are unaffected. At 20 years, HMA Maintenance 5 and 6 (years 24, 28) and PCC Rehabilitation 1 (year 27) now drop to $0 as they should. Salvage remains at the policy fraction; see Section 11 on whether the fraction should change with the period.',
  'Patched MBT copy recalculated at 30 years: NPW unchanged to the cent ($8,809,266.42 and $8,028,734.49).');

finding('5.4', 'Default analysis period of 10 years',
  'A new analysis discounts salvage at year 10 unless the user notices and changes D33.',
  'General Information!D33 = 10 in the shipped template; the Overview says 30 years; every maintenance policy runs to year 30.',
  'A test value left in the template.',
  [['General Information!D33', '10', '30']],
  'Salvage now discounted at year 30 by default.', 'Value confirmed after recalculation.');

finding('5.5', 'Missing-airport revenue lookup returns #N/A',
  'With "Account Indirect Cost" = Yes and an airport outside the 17 with revenue data, Average Daily Revenue shows #N/A, every indirect-cost row shows #N/A, the alternative NPW shows #N/A and the Summary shows #N/A (NS screenshots, Alt 3 and Alt 4). On Excel 2019 or earlier the same cells show #NAME? because XLOOKUP does not exist.',
  `Template F2 = SUM(_xlfn.XLOOKUP(${GI}!D10,RevenueData!A:A,RevenueData!B:H)). General Information!D39 wraps the same lookup in IF(COUNTIF(...)=0,"Missing Airport Revenue",...), so the guard existed and was not carried into the templates. RevenueData also contains text in several numeric cells (" $-   ", " NA ", descriptive text on the XNX and M54 rows).`,
  'The guarded formula was written once on General Information and the unguarded inner expression copied into the templates. XLOOKUP was chosen over functions available in every Excel version.',
  [['TMP(NewHMA)_IndirectCost!F2\nTMP(NewPCC)_IndirectCost!F2', `=SUM(_xlfn.XLOOKUP(${GI}!D10,RevenueData!A:A,RevenueData!B:H))`, F2new],
   ['Same sheets, G2 (new)', '(empty)', `=IF(AND(${GI}!$D$38="Yes",F2=0),"No revenue data for this airport - lost revenue set to $0. Contact Aeronautics.","")`],
   ['General Information!D39', `...SUM(_xlfn.XLOOKUP(D10,RevenueData!A:A,RevenueData!B:H))...`, `...SUMIF(RevenueData!$A:$A,$D$10,RevenueData!$B:$B)+...+SUMIF(...$H:$H)...`]],
  'For the 17 airports the result is identical (SUMIF ignores the text cells exactly as SUM(XLOOKUP) did). For any other airport lost revenue is $0 and a warning appears in G2 and on D39, instead of #N/A propagating to the Summary. Works on every Excel version.',
  'Recalculated: F2 = 0 and G2 = "" with no airport selected; original file shows 39 #NAME? and 60 #N/A cells under the same recalculation, v1.2.0 shows none.');

finding('5.6', 'No blank guard on PCC pay-item rows 14 to 22',
  '#N/A appears in the Pay Item column, Item Cost, Subtotal and Total of a PCC alternative as soon as one of rows 14 to 22 is left empty (NS screenshot, Alt 1 New PCC row 18).',
  'TMP(NewPCC)!B14:B22 and TMP(NewPCC)_IndirectCost!B14:B22 read =VLOOKUP(C14,CHOOSE({1,2},Table2[Pay Item Description],Table2[Pay Item No.]),2,0). Row 13 in the same sheets, and rows 13 to 22 in the HMA templates, read =IF(C13="",0,VLOOKUP(...)). A placeholder row in Pay_Items (row 3, description blank, "Default") lets the unguarded lookup resolve when the linked cell holds an empty string, which is why the defect appears only sometimes.',
  'A partial fix: the guard was added to row 13 of the PCC templates and to every row of the HMA templates, but not to rows 14 to 22 of the PCC templates.',
  [['TMP(NewPCC)!B14:B22\nTMP(NewPCC)_IndirectCost!B14:B22', '=VLOOKUP(C14,CHOOSE({1,2},Table2[Pay Item Description],Table2[Pay Item No.]),2,0)', '=IF(C14="",0,VLOOKUP(C14,CHOOSE({1,2},Table2[Pay Item Description],Table2[Pay Item No.]),2,0))']],
  'Empty rows give 0, consistent with the HMA templates. No change to filled rows.',
  'Patched MBT copy: Alt 2 (New PCC) rows 17 to 22 read 0 after recalculation where the unpatched template gives #N/A.');

finding('5.7', 'Four pay items share the description "Separation Geotextile"',
  'Selecting a separation geotextile under P-208, P-209 or P-219 returns the pay item number and unit cost of the first geotextile line in the table.',
  'Pay_Items!D20 (P-154-5.2, "Separation geotextile"), D28 (P-208-5.2), D30 (P-209-5.2) and D32 (P-219-5.2) carry the same text. Every template lookup keys on Table2[Pay Item Description] because the ActiveX combobox writes the description, not the item number, to its linked cell.',
  'The lookup key is the description because that is what the control returns; the pay-item list was extended with one geotextile line per base-course specification without making the descriptions unique.',
  [['Pay_Items!D20', 'Separation geotextile', 'Separation Geotextile (P-154)'], ['Pay_Items!D28', 'Separation Geotextile', 'Separation Geotextile (P-208)'], ['Pay_Items!D30', 'Separation Geotextile', 'Separation Geotextile (P-209)'], ['Pay_Items!D32', 'Separation Geotextile', 'Separation Geotextile (P-219)']],
  'Each geotextile line now resolves to its own item number and unit cost. Existing alternative sheets that selected one of these items will show the old description in the linked cell and must be re-selected.',
  'Recalculated; descriptions confirmed unique in Table2.');

finding('5.8', 'Airport dropdown stops four airports short',
  'Warren County Memorial, William L. Whitehurst Field, Winchester Municipal and Wolf River cannot be selected in General Information D9.',
  'The data validation on D9 lists $L$9:$L$84; the airport table occupies L10:L88.',
  'A static range in the validation while the table grew; a table reference would have followed the rows.',
  [['General Information!D9 validation', '$L$9:$L$84', '$L$10:$L$88']],
  'All 79 airports selectable.', 'Validation formula confirmed after patch.');

finding('5.9', 'Closure durations ship as zero',
  'Every lost-revenue row is $0 until the user types closure days into F4:F10 of each alternative sheet. The MBT and SRB files carry hand-written production-rate formulas in those cells; the template does not.',
  'TMP(NewHMA)_IndirectCost!F4:F10 = 0 and TMP(NewPCC)_IndirectCost!F4:F5 = 0. MBT and SRB Alt sheets hold formulas such as =ROUNDUP((C5/15000),0)+ROUNDUP((C8/20000),0)+ROUNDUP((C8/50000),0) (surface treatment at 15,000 SY/day plus markings) and =ROUNDUP((C5/3800),0)+ROUNDUP((C8/50000),0) (mill and overlay at 3,800 SY/day).',
  'The production-rate defaults were developed during the MBT and SRB analyzes, after the template was issued, and were never back-ported.',
  [['TMP(NewHMA)_IndirectCost!F4:F10', '0', 'Production-rate formulas from the MBT run (Appendix A gives each cell); cells remain gray inputs and may be overridden'],
   ['TMP(NewPCC)_IndirectCost!F4:F5', '0', 'Production-rate formulas from the MBT run (Appendix A)']],
  'New analyzes start from the same closure assumptions as the two completed runs. On a blank template F4:F10 evaluate to 0 (HMA) and 0 and 7 (PCC, the rehabilitation formula carries a fixed 7-day term) until an area is entered.',
  'Recalculated on the blank template and on the MBT copy (6, 6, 8, 8, 15, 6, 8 days for HMA; 9 and 18 for PCC), matching the MBT report.');

finding('5.10', 'Chart helper rows hard-positioned; no years on the axis',
  'The expenditure chart on each alternative sheet reads 1 to 31 on its axis, mixes undiscounted and discounted bars, and would silently misplace a bar if a policy year changed. Its title is set once by VBA at creation ("Alternative 4" on an Alt 1 sheet in the NS screenshots).',
  'Column L is populated only at rows 37, 41, 45, 49, 53, 57, 61, 65 and 67 with links such as L41 = D37, that is, row 37 plus the policy year of Maintenance 1. The chart series have no <c:cat> element; a c15:filteredCategoryTitle extension referencing K37:K67 shows the category was set and later filtered out.',
  'The helper block was built by hand to match the current policy years and never linked to them; the category axis was dropped at some point.',
  [['TMP(NewHMA)!L37:L67 (and M)', 'L37 = D36, L41 = D37, ... L67 = D44; other rows empty', `=IF(K37-$C$6>${GI}!$D$33,0,IF(K37=$C$6,$D$36,0)+SUMIF($C$37:$C$44,K37-$C$6,$D$37:$D$44))  (M uses column E)`],
   ['TMP(NewHMA)_IndirectCost!L37:N67', 'as above, N = E36 etc.', `L: SUMIFS over D with criteria "<>*Indirect*"; M: SUMIFS over D with "*Indirect*"; N: SUMIF over E (all with the period guard)`],
   ['PCC templates L38:N68', 'as above, one row lower', 'same formulas with the PCC row references ($D$37, $C$38:$C$42 ...)'],
   ['Chart parts chart1 to chart5', 'no category; 2 or 3 series (Actual/Discounted or Direct/Indirect/Total Discounted)', 'category = K37:K67 (calendar years); direct templates keep the Actual series; indirect templates keep Direct and Indirect, stacked; axis format "0"']],
  'Bars now land on the calendar year that Maintenance Policies dictates and respect the analysis period; the axis shows 2027 to 2057. Column N (total discounted by year) remains for the Summary.',
  'Patched MBT copy: sum of L plus M equals the activity totals ($10,959,044 HMA; $7,281,371 PCC) and sum of N equals NPW to the dollar. Year-by-year values in Appendix C. Chart XML parsed by an independent reader with the intended ranges.');

finding('5.11', 'HMA Rehabilitation template latent defects',
  'Not visible today: the template is very-hidden and disabled in the VBA since 19 November 2022. If re-enabled, Maintenance 5 would be discounted by (1+r) to the power 2044 and the salvage row would multiply an empty cell.',
  "TMP(HMARehab)!C42 = C36+'Maintenance Policies'!E65 adds the construction year (e.g. 2020) to the policy year. D44 = -'Maintenance Policies'!D71*AV22, but the rehabilitation total in that template is AN22; AV22 is empty.",
  'The template was abandoned when the rehab options were hidden and drifted from the others.',
  [["TMP(HMARehab)!C42", "=C36+'Maintenance Policies'!E65", "='Maintenance Policies'!E65"], ["TMP(HMARehab)!D44", "=-'Maintenance Policies'!D71*AV22", "=-'Maintenance Policies'!D71*AN22"]],
  'None today. Correct behavior if the option returns.', 'Recalculated without error.');

// 6 charts (covered) -> 6 instructions text
children.push(h1('6. Instructions text box'));
children.push(p('The Instructions sheet holds its text in a drawing text box, not in cells. Six runs were changed; the wording of every other run is unchanged. The text box was extended from row 59 to row 76 because it clips overflow and the text grew by about a third.'));
const trows = [['Run', 'v1.1.2', 'v1.2.0']];
for (const [i, a, b] of diff.textdiff) trows.push([String(i), a, b]);
children.push(table(trows, [700, 4330, 4330], { size: 14 }));
children.push(p(''));
children.push(p('The Overview text box is unchanged. NS supplied a rewrite (2022 APTech framework, 2026 NS and ARA update, the four airport criteria); it should be pasted once the wording is approved.'));
children.push(p('One more defect surfaced on the map itself. The boundary source publishes the Tennessee/North Carolina line twice, as two tracings that agree to within 0.01 degrees, so that stretch of the outline was drawn doubled. The second copy is dropped and the embedded outline is now twelve segments and 179 points.'));
children.push(p('Four misspellings carried in the issued text were corrected and nothing else in the wording was touched: the Overview text box read "Federal Aviation Adminimstration", "pavement clossures" and "associeted with limited facility uses", and the General Information label at C25 read "Intial Construction Year". Every user-visible string in the delivery was then run through a spell check: the shared and inline strings on each visible sheet, the text inside formula literals, the chart and drawing text runs, the comments and message strings in the KML macro, and the prose in these two reports and the change notes, against a word list carrying the domain vocabulary, the Tennessee place names in the airport list and the code identifiers. The run is clean; the checker is kept as verification/spellcheck.py so it can be re-run on the next revision.'));
children.push(p('Cell notes on General Information D26 and D39 carried the author\'s name; the author is now "ARA". The workbook carried two external links to files on a personal drive (v1.1.004 and a Savannah copy, referenced only by three sheet-scoped copies of the Airport_Name name); the links and those three names were removed. The workbook-level Airport_Name name (General Information D9) is unchanged. The package itself was also cleaned so that nothing in it refers outside the file or to an individual: the six cached printer-driver parts (printerSettings1 to 6, each tied to a page-setup element by a relationship) were dropped and the page setups keep their orientation and scaling; an empty Power Query (DataMashup) stub in customXml was removed with its relationship and content type; the SharePoint path Excel had cached as the last save location was stripped from workbook.xml; and the document properties now name the organization (creator "TDOT Aeronautics Division / Applied Research Associates", last modified by "ARA", company "TDOT Aeronautics Division") instead of two individuals. The source column of the Typical Values sheet is plain text rather than HYPERLINK formulas, so the only outward link left in the workbook is the tn.gov link that was already on Instructions. The one item that cannot be scrubbed without Excel is the VBA project binary, which carries the path of the machine it was last compiled on; it is inert but visible to anyone who unpacks the file.'));

// 7 Summary sheet
children.push(h1('7. The Summary sheet'));
children.push(p('The VBA project is not changed. The v1.1.2 Summary consisted of five columns written by the Alternative Setup form (alternative, name, initial construction, present worth, description) and one clustered-bar chart. v1.2.0 keeps those five columns and that chart exactly as the form writes them, moves the chart beside the new charts, and adds from column G onward a results table, a verdict line and five charts that are ordinary worksheet formulas. They read the list of alternative worksheets that the form stores on the hidden Database sheet, reach into each worksheet with INDIRECT, and therefore follow whatever alternatives exist, up to the four rows the form allows.'));
children.push(table([
  ['Block', 'What it shows', 'How it is computed'],
  ['Results table G3:R7', 'Worksheet, alternative, type, initial construction, maintenance PW, rehabilitation PW, lost revenue PW, salvage PW, net present worth, difference to the lowest NPW, closure days in the period, runway availability', 'INDEX(Database!A:A, row) etc. for the names (whole-column INDEX survives the Database row deletion the VBA performs when an alternative is removed; a fixed reference would become #REF!); INDEX/MATCH on "Total" and "Net Present Worth"; SUMIFS over the NPW-table window (rows 37 to 52) on the row labels ("Maintenance*", "Rehabilitation*", "*Indirect*", "Salvage*"); closure days = sum of the by-year closure block (indirect cost / daily revenue, years within the analysis period), falling back to SUM of F4:F10 when lost revenue is off'],
  ['Verdict line G9:G10', 'Lowest present worth, margin to the next alternative, discount rate, period, lost-revenue setting; whether the same alternative is lowest at 2 percent (the OMB A-94 real rate that FAA PGL 22-01 substituted for the fixed 7 percent in June 2022) and at 7 percent', 'INDEX/MATCH on the NPW column; SMALL for the margin; comparison against the 7 percent row of the sensitivity block'],
  ['Comparison block G12:P17', 'Agency cost (initial + M&R + salvage) and user cost (lost airport revenue) as present worth and equivalent uniform annual cost, total of both, difference to the lowest in dollars and percent, and which alternative is lowest', 'The Caltrans and FHWA RealCost deterministic-results layout applied to an airport: agency cost and user cost by alternative, each as PW and EUAC. EUAC = PW x r(1+r)^P/((1+r)^P-1) from the results table columns; no new inputs'],
  ['Chart 1 (G22)', 'Present worth by category, stacked, salvage below zero', 'Category block W3:AA9'],
  ['Chart 2 (N22)', 'NPW versus discount rate, 2 to 8 percent in 0.25 steps', 'Sensitivity block W10:AA36: initial + SUMPRODUCT((year <= period) x cost / (1+r)^year) per alternative'],
  ['Charts 3 and 4 (G40, N40)', 'Expenditure by calendar year (undiscounted) and cumulative discounted cost', 'By-year block W39:AJ71: SUMIF on Year Applied, running sum of the discounted column'],
  ['Chart 5 (G58)', 'Runway closure days by calendar year', 'Indirect-cost rows divided by daily revenue (F2), by year'],
  ['Pavement section G80:Q86, charts 6 and 7', 'Each layer\'s thickness read back from the pay-item quantities, the total section, a cross-check against the excavation quantity, and the section as a string to paste into the alternative description; a second string, filled only when a shoulder area is entered, gives the same quantities spread over mainline plus shoulder. Chart 6 stacks the mainline section, surface on top; chart 7 fills in only when a shoulder area is entered', 'SUMIFS over the alternative sheet rows 13 to 22 keyed on the unit and the pay-item number: volume items give inches = 36 x C.Y. / mainline S.Y.; asphalt gives inches = 2,666.67 x tons / (pcf x mainline S.Y.) with the unit weight in J81; concrete carries its thickness in the pay-item name. No cost depends on any of it'],
  ['Summary starts at the results table', 'Columns A to F hidden: the form still writes its table and Chart 1 there, and the results table repeats it with more detail. Navigation buttons moved to G1 and H1', 'Removes the duplicate block beside the charts so the table and the seven charts align on the left edge; the original Alternatives Comparison chart is set to plot hidden cells and now reads in millions'],
  ['Chart legends and scales', 'Legends moved to the right of each plot; money axes read to one decimal in millions; an alternative that has not been created is unnamed', 'The bottom legend collided with the category labels and the axis title; five gridlines all reading $17M told the reader nothing; empty slots no longer appear as "Alt 3" and "Alt 4" on axes and in legends'],
  ['Axis titles', 'Every chart in the workbook and in the decision workbook now names both axes: calendar year against cost, alternative against present worth, discount rate against net present worth, alternative against thickness', 'Added to the five original alternative-sheet charts and the original Alternatives Comparison chart as well as the seven new Summary charts'],
  ['Navigation', '"General Information" and "Instructions" buttons on Summary; "View Summary" (D44) and "Typical Values" (D46) buttons on General Information', 'HYPERLINK cells; no macro, so they work when ActiveX is blocked'],
], [1900, 3300, 4160], { size: 14 }));
children.push(p(''));
children.push(p('The chart-data blocks are labelled "calculated automatically; do not edit". The print area covers the table and charts one page wide in landscape. Appendix B lists the formulas as they stand in row 4 and in the first row of each block; the remaining rows repeat them with the row number substituted. A first draft of this update replaced the VBA Output module instead (Output.bas); it was superseded by this formula-driven sheet so that nothing has to be imported, and it is not part of the deliverable.'));

children.push(h1('7c. Look and first-run usability'));
children.push(p('The layout, fonts, grey input cells, logo and button positions are unchanged, so the workbook still reads as the same tool. What changed is the first-run guidance and the visual hierarchy. None of it touches a calculation: every cell added is text, a HYPERLINK or a format.'));
children.push(table([
  ['Where', 'Change', 'Why'],
  ['General Information F2:J7', 'A "How to use this workbook" card beside the logo: fill the grey cells, click Alternative Setup, click View Summary, plus the note that grey cells are inputs and white cells calculate', 'A new user had nothing on screen telling them the order of operations; the Instructions sheet is a separate tab they had to find'],
  ['General Information F9:J9', 'A live status line naming whichever required input is still empty (airport, construction year, mainline area, markings area, analysis period, discount rate). Conditional formatting turns it green when the list is empty', 'The reviewer built alternatives before the inputs were complete and got zeroes; the line makes that state visible before the click'],
  ['General Information B8, B20, B32', 'The three section labels sit on a light band with a blue rule', 'Gives the form three visible blocks instead of one long list'],
  ['General Information D35', 'The salvage sentence is styled as a note, not as an input value', 'It sits in the input column and reads like a value to fill in'],
  ['General Information print setup', 'Print area A1:J48, landscape, one page wide', 'The sheet printed across two pages with the hidden lookup columns'],
  ['Every alternative worksheet, row 1', '"General Information" and "Summary" buttons and a one-line reminder of which cells are inputs', 'The alternative sheets had no way back; the reminder answers the question the reviewer asked about which cells to fill'],
  ['Summary results table and comparison block', 'Conditional formatting highlights the lowest-cost row', 'The answer is visible before reading any number'],
  ['All sheets', 'Tab colors: navy for General Information, blue for Summary, light blue for alternative worksheets, grey for reference', 'Groups sixteen tabs into four kinds'],
], [1700, 3700, 3760], { size: 13 }));
children.push(p(''));
children.push(p('The navigation row and the tab color live on the hidden templates, so every alternative the Alternative Setup form creates carries them without any change to the VBA.'));

children.push(h1('7b. Typical Values sheet and input hints'));
children.push(p('A new last sheet, "Typical Values", is reference only: no cell on it feeds the calculation, and it was added last so no sheet index shifts. It answers the question a user asks while typing into General Information: is this number in the right range. Rows that can be read from the workbook itself are live formulas (Pay_Items unit costs, Maintenance Policies years and salvage fractions, the daily revenue of each of the 17 airports); rows from outside sources carry the source URL in the last column as plain text (not a hyperlink, so nothing on the sheet reaches outside the workbook).'));
children.push(table([
  ['Section', 'What it gives the user'],
  ['1. What TDOT Aeronautics manages', '78 public-use airports (6 commercial service, 72 general aviation), 69 in the NPIAS, about 70 in the APTech pavement management network; PCI objectives 78 runway and 75 taxiway; TASP system-wide pavement estimates; grant shares (up to 90 percent state or federal, 95/5 discretionary, maintenance program 100 percent state)'],
  ['2. Typical runway geometry', 'Seven Tennessee airports with runway dimensions and the mainline area they imply, a typical range of 4,000 to 6,000 ft by 75 or 100 ft (33,000 to 67,000 SY), and a rule of thumb for the markings area. Directly answers cells D26 and D28'],
  ['3. Unit costs', 'Each workbook default from Pay_Items column F beside recent bid prices: P-401 surface $185 to $232 per ton in 2025 against the $130 default; P-101 milling $4.70 to $10.00 per SY; P-501 $165 per SY as a planning value; and the all-in $210 to $280 per SY of recent Tennessee runway reconstructions'],
  ['4. Closure-day defaults', 'Every production rate built into cells F4:F10, with the days each gives for a 6,000 by 100 ft runway, so a user can see why the defaults are 7, 9 and 19 days and override them with knowledge'],
  ['5. Economic parameters', 'TDOT 3 percent over 30 years (live); FAA PGL 22-01 of June 2022 pointing to the OMB A-94 real rate, 2.0 percent for 2026; the pre-2022 7 percent and 20-year rule; AC 150/5320-6E at 4 percent; AAPTP 06-06 and Caltrans at 4 percent; the salvage fractions live from Maintenance Policies against the remaining-life method those guides use'],
  ['6. Maintenance timing and service life', 'The policy years from Maintenance Policies beside published experience: asphalt runways reach PCI 70 in about 12 to 15 years and concrete in about 40 (FAA 2014 performance trends), joint sealant lives of 3 to 20 years by type, the AC 150/5320-6G minimum layer thicknesses'],
  ['7. Daily revenue and dropdown counts', 'Each of the 17 airports with revenue data, its daily revenue, what a 7-day and a 19-day closure costs, and which RevenueData columns hold text and are therefore counted as zero'],
], [2400, 6960], { size: 13 }));
children.push(p(''));
children.push(p('General Information carries a "Typical Values" button (D46) beside the "View Summary" button, and short hints in column F beside the construction year, areas, analysis period, discount rate, mobilization, engineering and indirect-cost cells.'));
children.push(p('Research method and limits: the values were gathered in September 2026 from TDOT, TASP, FAA, OMB, Caltrans, AAPTP and state bid tabulations. The build environment could not open several of those sites directly, so figures come from search excerpts of the cited documents; verification/RESEARCH_SOURCES.md lists what was confirmed and what still needs a look at the primary PDF (the AIP Handbook paragraph numbers, the 6G thickness table rows, and Tennessee bid prices for P-209, P-154, seal coats and crack sealing).'));

children.push(h1('7bb. Project identity, locator map and Google Earth export'));
children.push(p('The Summary carried no project identity: printed, it was a table of alternatives with nothing saying which airport it belonged to. It now opens with a one-line header above the results table (airport and FAA identifier, city, TDOT region, branch, project type, construction year, analysis period and discount rate) and a PROJECT LOCATION block beside the table: a locator map of Tennessee showing all 79 airports in the dropdown as grey dots with this project as a labelled orange dot, and beneath it the airport, city and county, region, coordinates, elevation, branch and mainline area.'));
children.push(p('The airports are colored by TDOT Grand Division, light blue for West, blue for Middle, navy for East, with a legend in the cells beside the plot so the map keeps its full height. That grouping was chosen because it is the one Pay_Items is built around, and putting it on the map exposes a gap. Pay_Items carries a header, Average Pay Item Unit Cost, spanning three columns headed Middle, West and East; all three are empty. Every alternative prices off the single Unit Cost column beside them, so a West-division project is estimated on exactly the same numbers as a Middle one. A line under the project block states that and names the division the project sits in. Nothing about the calculation changed: the Method sheet now carries it as an assumption worth a decision, alongside the salvage asymmetry and the lost-revenue treatment.'));
children.push(p('The map is an ordinary XY scatter chart, not a mapping service. The Tennessee outline and the airport coordinates are embedded in the map-data block on the Summary, below the other chart data and outside the print area, so the map draws with no internet connection, no Bing map service and no add-in. This matters on a locked-down machine: Excel\'s own Filled Map and 3D Map both call out to Bing and would leave an empty frame. Coordinates come from airportsdata (PyPI, carrying FAA and OurAirports values) and the outline from the intermediate-resolution political boundaries in basemap-data, both retrieved 13 September 2026, cited with their retrieval date in geo_data.py. Five small private fields in the dropdown have no published coordinates: they plot no dot and the sheet says so rather than guessing.'));
children.push(p('A Google Earth export ships alongside as a VBA module, LCCA_KML_Export.bas, imported once through Alt+F11, File, Import File. It was deliberately not injected into the workbook\'s existing VBA project: that binary holds the Alternative Setup form, it could not be tested in Excel in the review environment, and a malformed project would break every macro in the file. Running ExportLCCAKML writes a .kml beside the workbook with four things in it.'));
children.push(table([
  ['In the KML', 'What it shows'],
  ['Airport placemark', 'The whole result in the description bubble: every alternative with its initial cost, net present worth, closure days and pavement section, then the verdict line and the discount-rate check'],
  ['Runway footprint per alternative', 'A schematic rectangle oriented from the runway number in the branch name and sized from the mainline area entered, green for the lowest net present worth. Labelled as schematic, not survey geometry'],
  ['Net present worth bars', 'One extruded polygon per alternative, height proportional to its net present worth, so the comparison stands up in three dimensions beside the runway'],
  ['Timed events', 'Every maintenance and rehabilitation event as a placemark stamped with the calendar year it happens and carrying its cost and closure days, so the Google Earth time slider walks the analysis period'],
], [2400, 6960], { size: 13 }));
children.push(p(''));
children.push(p('The macro reads cells and writes a text file: it changes nothing, references no library and reaches no network (the KML carries no external icon references either, so it renders on a machine with no internet). Because VBA cannot be executed in the review environment, the same logic was implemented in Python and run against both worked examples; the output parses as KML and the McKellar-Sipes file carries four folders, thirty-five placemarks, eight polygons and twenty-six timed events spanning 2027 to 2057. Both samples ship in verification/examples.'));
children.push(p('The output was then linted against the KML 2.2 rules (verification/kml_check.py checks element order, style resolution, closed rings, coordinate range, time-stamp placement and extruded geometry, and walks the time slider year by year). Four corrections followed. Element order inside each placemark now matches the KML sequence of name, description, time, style, geometry, which Google Earth tolerated either way but a strict GIS reader need not. The document gained a legend with color swatches and a LookAt, so the file flies to the runway on open instead of landing the reader in mid-Atlantic. The footprint length now comes from a runway width on the Summary rather than a fixed 100 ft: at McKellar-Sipes that assumption drew a 9,008 ft runway against the real 6,005 ft, and the true 150 ft width puts the footprint on the pavement. And the present-worth bars are scaled so the tallest is always 700 m, where the fixed scale had put a 3.4 km spike over a 1.8 km runway. The file now lints clean, with the slider running 2027 to 2057 and the right events appearing in each year.'));

children.push(h1('7d. The Summary as a dashboard, and the setup flow'));
children.push(p('The Summary used to open with a twelve-column table. It now opens with the answer. Six tiles sit above the results: the lowest present worth with the section it buys, the margin to the next alternative, the equivalent annual cost, the initial construction of the winner, the unit cost per square yard, and whether the winner survives every discount rate in the sensitivity block. Each tile is a formula over cells that already existed on the sheet, so the dashboard adds no calculation and nothing new to maintain. The margin tile carries a rule rather than a fixed colour: it turns amber when the two best alternatives are within five percent of each other, the point at which a reviewer should treat them as tied rather than read a winner out of the table.'));
children.push(p('The rate-sensitivity tile is the one that needed care. Its first form compared the winner at 2 percent and at 7 percent and reported "holds 2-8%" if they agreed, which a flip anywhere between them would not have caught, and its value was painted green whatever it said. The sensitivity block now carries a column naming the lowest-cost alternative at each of its 25 rates, the tile counts that column, and a rule turns it green only while it says the winner holds and amber when it says the winner changes.'));
children.push(p('Three smaller changes go with it. The net-present-worth column carries data bars, so the size of the difference is visible in the table itself. Chart 1 now uses one light-to-dark ramp across its five categories instead of the mixed greys it had, so the stack reads in the order the categories are listed. And a chart 8 was added beside chart 5: each alternative\'s initial construction divided by the mainline area, against the $210 to $280 per square yard all-in range for recent Tennessee runway work already cited on Typical Values. The band is a floating bar drawn from two cells, so the range can be updated without touching the chart. The comparison is deliberately unequal in a stated direction: the published range covers pavement, lighting and grading while this workbook prices the pavement contract, so every bar should sit below the band. A bar far below it, or above it, is a quantity worth checking before the analysis is believed.'));
children.push(p('The chart the Alternative Setup form maintains was parked below the dashboard with a line saying what it is. It shows initial construction and present worth per alternative, which chart 1 shows with the categories broken out, but the macro still writes to it, so it was moved rather than removed.'));
children.push(p('The Summary row map moved with the strip, and that is the one change in this release that reaches outside the file. The results table is rows 12 to 15 where it was 4 to 7; the verdict lines are G17 and G18 where they were G9 and G10; the comparison block starts at row 22; the pavement-section block at row 90 with its asphalt unit weight at J91; the map-data block at row 130. The KML macro, the click-through script, the static verifier and both worked examples were updated together and re-run. A private copy of the macro taken before this release would read the wrong cells and would need the same edit.'));
children.push(p('Alongside the dashboard, the workbook now states the order a user is meant to work in. The card on General Information lists five steps and names the sheet for each: Overview and Instructions, General Information, Pay_Items, Alternative Setup, Summary. Each of those sheets says which step it is, in the same place on the sheet, so someone who lands on one of them mid-way knows where they are. General Information gained buttons for Pay_Items and Maintenance Policies, so every step is one click from the sheet the user starts on, and it still prints as one landscape page.'));
children.push(p('Pay_Items and Maintenance Policies were the two sheets that had never been designed. Pay_Items now carries the navigation band and the step line in row 1, keeps its header frozen and repeats it on every printed page, and marks the Unit Cost column as the input it is with a grey fill and a currency format. The part headings read as bands down the left of the table. Four notes under it state what the sheet drives, that the Middle, West and East average-cost columns are empty so a West division project is priced on statewide numbers, that 29 of the 56 pay items carry no unit cost at all and will price at zero if an alternative uses them, and where the published sources sit. Table2 over C2:I59 and every unit cost are untouched. Maintenance Policies gained the same navigation band, a title, and four lines beside the TDOT logo: that it is reference only because the schedules live in the hidden templates, which table drives which alternative type, what the Rate column means, and that closure days come from the production rates on Typical Values rather than from this sheet. Its four table headers sit on a navy band.'));

children.push(h1('7e. The TDOT mark'));
children.push(p('The logo was on four of the nine sheets a user opens \u2014 Overview, Instructions, General Information and Maintenance Policies \u2014 each anchored and sized a little differently, and missing from the Summary, Pay_Items, Typical Values, Method and every alternative worksheet. Measuring the four placements turned up a defect none of them had been checked for: the artwork is 723 by 316 pixels, an aspect of 2.288, and the placements ran 2.123 on two sheets and 2.060 on the other two. Every one of them was drawn 8 to 11 percent taller than the artwork, so the lettering was stretched and the TN square was not square.'));
children.push(p('Three placements were considered. The page header would have carried the mark on every sheet at no cost in rows, but only in Page Layout view and on paper, not in the normal editing view. A masthead on the four document sheets with the header everywhere else would have cost nothing at all, but would have left the working sheets unbranded on screen. The third was chosen: row 1 becomes a band on all thirteen sheets a user can reach. Row 1 already existed everywhere and already carried the navigation buttons, so the band costs only its extra height \u2014 21 points to 40. Forty is the floor at which the second line of the lockup stays legible; the mark is 1.27 inches wide there, at the artwork\'s own aspect, anchored to the right-hand edge of each sheet\'s content where it cannot collide with the buttons or the step line.'));
children.push(p('The honest cost is the one the option carries: on Pay_Items and on the alternative worksheets that height sits above the frozen header, so it is gone from every screen rather than only the first. Against that, row 1 is now a print title on all thirteen sheets, so the mark prints at the top of every page rather than only the first \u2014 Pay_Items repeats rows 1 and 2, the band and its table header. Three sheets had no drawing part in the package at all and now have one: Pay_Items, Typical Values and Method.'));
children.push(p('Two print defects surfaced while checking that the mark actually lands on the page. Overview and Instructions hold a text box 7.6 inches wide on a portrait page with 7.1 inches between the margins, so the right-hand side of the framework description had been clipped in print since v1.1.2; both sheets now scale to one page wide. And each alternative worksheet was printing its chart-data columns across half a dozen pages, because it had no print area at all; the print area is now the cost and present-worth tables, A1 to G56. Between them the worked example went from 41 printed pages to 30.'));
children.push(p('The step lines moved out of the band to make room for the mark. Each sheet carries a short "STEP n of 5" chip beside its buttons and the sentence sits on the row below: row 2 on Overview and Instructions, row 3 on the alternative worksheets, row 10 on the Summary. On Overview and Instructions the Aeronautics address block moved to the left margin under the band, into the space the logo used to occupy, so the letterhead still reads as one block. On the Summary the project identity line moved from L1 to a row of its own at G2, which is what freed the right-hand end of the band.'));

children.push(h1('7f. The pay item picker, and which division an alternative is priced from'));
children.push(p('The pay item picker on every alternative worksheet is a Forms 2.0 combo box listing two columns of Pay_Items \u2014 the pay item number and its description. A Forms combo box sizes its drop-down list to the control unless ListWidth says otherwise, and with no explicit column widths it divides that evenly between the columns, so each column had about 1.45 inches. Ten characters of pay item number fit in that; the descriptions run to 76 characters and did not, which is why the list read as truncated fragments and the reviewer could not tell two geotextiles apart without opening Pay_Items.'));
children.push(p('All 105 pickers now carry a ListWidth of 9 inches, so an even split gives each column 4.5 inches, wide enough for the longest description in the catalogue with the number still beside it. The picker column on each template goes from 41.7 to 53.7 characters and every combo box is re-anchored to end exactly at that column\'s edge, so what was picked reads in the cell as well as in the list. This is a change inside the persisted ActiveX streams rather than in any XML part: the MorphDataControl property mask gains fListWidth, the DataBlock gains the four bytes that property needs at its alignment, and cbMorphData grows to match. All 105 were then re-parsed with an independent implementation of [MS-OFORMS] and each consumes exactly to the end of its stream.'));
children.push(p('Pay_Items has carried Middle, West and East columns beside Unit Cost since v1.1.2 under a header reading "Average Pay Item Unit Cost", and nothing in the workbook read them: a West division project was priced on the same statewide numbers as a Middle one. They are read now. Each alternative worksheet has a "Price from" list in C11, one row above the pay item table \u2014 Regular, Middle, West or East \u2014 and all 156 unit cost lookups on the five templates read the column it names. A new alternative ships set to Regular, which is the Unit Cost column, and any item the chosen division leaves blank falls back to Unit Cost, so filling one regional cost prices that one item regionally and changes nothing else. An empty C11 reads as Regular rather than producing an error. The note beside the list names the division the selected airport sits in, read from General Information D13, so the person choosing can see which column matches the project.'));
children.push(p('The columns ship empty, so every result in this release is unchanged: both worked examples and the Murfreesboro regression recompute to the numbers they produced before. What has changed is that filling them is now a data-entry task for Aeronautics rather than a code change.'));
children.push(p('Two adjacent defects were fixed in the same pass. TMP(HMARehab) was the one template whose working columns I to BL had never been hidden, so an HMA overlay alternative opened showing its internal per-event cost blocks and chart data; they are hidden now and that sheet\'s chart is set to plot hidden cells so it still draws. And an empty pay item line printed a 0 in the "Pay Item" column right beside its line number, because the lookup returned 0 rather than blank for an empty description; the line now reads blank and the item cost tests the description instead of that 0.'));
children.push(p('One row on General Information was dead. "Airport Owner" read column 6 of Table17, which is empty for all 79 airports, so it always showed blank. It now reads column 4 and is labelled "County", which is populated for every airport. State Region below it did not move, so the Summary still finds it.'));

children.push(h1('7h. Salvage: remaining life, computed rather than typed'));
children.push(p('Each salvage row on Maintenance Policies states the rule it uses in words, and both terms of that rule were already live in the file. The credit is taken at the analysis period, General Information D33, which a user sets; the asphalt mill and overlay happens in the year Maintenance Policies E22 names, which is a policy input. Only the fraction was a constant. Three quantities that have to agree and one of them frozen means the sheet can be right at a single combination, and asphalt was not right even at the default one: row 32 read "2 years (0.125) of mill and overlay" while row 22, four rows above it, places that overlay at year 20. An overlay laid in year 20 and valued in year 30 is ten years old and has six of its sixteen years left, which is 37.5 percent. Two of sixteen is what an overlay laid in year 16 would leave. The concrete figure had the same shape and happened to land correctly: ten of forty years at a thirty-year period is the 25 percent the cell carried, and it is wrong at every other period.'));
children.push(p('Both fractions are now formulas. The expected life each divides by has moved out of the formula into a new "Asset life (yrs)" column beside the salvage row, sixteen years for the overlay and forty for concrete, so the assumption is visible and editable rather than buried. The sentence in column C is built from the numbers rather than typed, so it cannot go stale again: at thirty years it reads "Salvage value: 6 of 16 years left on the mill and overlay placed in year 20 (37.5% of its cost) at the 30-year analysis period". The Year Applied shown on all four salvage rows now reads the analysis period instead of a fixed thirty, which is where the credit was always taken. An analysis period that ends before the overlay is ever laid credits nothing for asphalt, because there is no overlay to salvage. Tables 3 and 4 keep their zero, which is a stated policy, "need for reconstruction", rather than a remaining-life calculation.'));
children.push(table([
  ['Analysis period', 'Asphalt credit', 'Concrete credit', 'Murfreesboro margin, before', 'after'],
  ['20 years', '100%', '50%', 'PCC by $755,381', 'PCC by $804,006'],
  ['25 years', '68.8%', '37.5%', 'PCC by $880,950', 'PCC by $764,458'],
  ['30 years (default)', '37.5%', '25%', 'PCC by $780,532', 'PCC by $543,376'],
  ['36 years', '0%', '10%', '', ''],
], [1800, 1500, 1500, 2400, 2160], { size: 13 }));
children.push(p(''));
children.push(p('This moves published results, which is the point of raising it. At Murfreesboro at three percent over thirty years the asphalt alternative goes from $8,809,266.42 to $8,572,110.60 while concrete is unchanged at $8,028,734.49, so the margin narrows by about thirty percent and concrete remains the lower of the two. The direction of the old error was not even stable across periods: correcting it narrows the margin at thirty years and widens it at twenty, so the previous figure could not have been defended as conservative toward either pavement. What is deliberately not changed is the asymmetry of the basis, that concrete salvages a share of its whole initial construction while asphalt salvages a share of one overlay. That is a genuine modelling choice, it is what makes the concrete curves in chart 2 nearly flat while the asphalt curve falls, and it is recorded on the Method sheet as a decision for Aeronautics.'));

children.push(h1('7g. Correction: Maintenance Policies is live, not reference'));
children.push(p('The line this update first added to the top of Maintenance Policies said the sheet was reference only, that the schedules lived in the hidden templates, and that editing a number there changed nothing. That is the opposite of the truth, and it is corrected in the delivered file. Every Rate cell in column D and every Year Applied cell in column E is read by the alternative worksheets: D10 sets a maintenance quantity, E22 sets the year the HMA mill and overlay happens, D32 sets the HMA salvage credit and D46 the concrete one. Someone who trusted the old sentence could have changed a policy number expecting no effect and silently moved every result in the workbook. The sheet now says it is live and names what a change there moves, and the build verifier checks both that the sentence says so and that the templates really do read both columns, so the claim cannot drift away from the file again.'));

children.push(h1('7c. Method sheet'));
children.push(p('The workbook now carries its own calculation record. "Method" is the last sheet, reference only like Typical Values, and it states every calculation once in five columns: the step, the cell it lives in, the formula as it stands there, the rule in plain English, and where the rule comes from. Sections 2 to 8 follow one alternative from the quantities typed on its worksheet to its net present worth; section 9 covers what the Summary adds on top; section 10 reads the current analysis period, discount rate, capital recovery factor and the present worth of a dollar spent at year 20 live from the file; section 11 lists the assumptions a reviewer will ask about, each on a highlighted row. It exists because the discounting, the salvage rule and the lost-revenue rule previously existed only as formulas in cells, with nothing in the file stating them in words. General Information carries a "Method" button (D48) under the Typical Values button.'));
children.push(table([
  ['Section of the Method sheet', 'What it records'],
  ['1. The answer', 'Net present worth as the sum of the discounted column, where it is reported, and that every other block restates the same numbers'],
  ['2. Initial construction', 'Item cost as quantity times the Pay_Items unit cost, the VLOOKUP that fetches the price, the subtotal, mobilization and engineering percentages and the total'],
  ['3. When later work happens', 'Policy years drawn from Maintenance Policies, the year as an offset rather than a calendar year, and the drop of events beyond the analysis period'],
  ['4. What each event costs', 'The hidden per-event blocks: which policy rate drives each quantity, what the HMA and PCC events consist of, the assumed 12.5 ft slab layout behind the joint length, and mobilization and engineering on every event'],
  ['5. Closure days', 'Each production rate in F4:F10 (15,000 SY/day surface treatment, 3,800 SY/day mill and overlay, 10,000 LF/day crack sealing, 5,000 SY/day patching, 1,000 SY/day slab replacement plus seven days of cure) and that the cells are overridable inputs'],
  ['6. Lost airport revenue', 'The switch, the SUMIF chain over RevenueData, the warning for airports outside the 17, and closure days times daily revenue'],
  ['7. Salvage', '12.5 percent of the overlay for asphalt, 25 percent of initial construction for concrete, both credited at the end of the period'],
  ['8. Discounting', 'Cost divided by (1 + rate) to the power of the year, constant dollars with a real rate, the rate and period and where each comes from'],
  ['9. What the Summary adds', 'Results table, Database lookup, verdict, RealCost comparison block and the EUAC formula, sensitivity, by-year blocks and the section read-back'],
  ['10. Current settings (live)', 'Analysis period, discount rate, capital recovery factor, present worth of $1 at year 20, lost-revenue switch and the number of alternatives, all read from the file as it stands'],
  ['11. Assumptions worth knowing', 'Engineering now charged on initial construction; the salvage asymmetry; lost revenue treated as gross revenue; the two different PCC joint-length estimates; closure rates as defaults; 145 pcf; 2022 unit costs; no inflation and no risk analysis'],
], [2400, 6960], { size: 13 }));
children.push(p(''));
children.push(p('One inconsistency surfaced while writing the sheet and is recorded on it rather than changed: the cost side estimates the concrete joint length from 12.5 by 12.5 ft slabs, while the closure-day formula assumes 550 linear feet of joint per 50 by 100 ft panel. The closure formula therefore carries about a third less joint length than the cost formula, which understates PCC closure days and the lost revenue that follows from them. It is a policy or modelling decision for Aeronautics, not a defect introduced here.'));

// 8 decision workbook
children.push(h1('8. Companion decision workbook'));
children.push(p('TDOT_LCCA_Decision_Workbook.xlsx is a separate, macro-free workbook that carries the MBT and SRB analyzes and answers the questions the framework Summary cannot: where the difference between the alternatives comes from, whether it survives the FAA AIP rate, which input would flip it, and what the runway closures cost. It does not replace the framework; it reads the framework\'s outputs.'));
children.push(h2('8.1 Method'));
children.push(p('For each alternative, with r the discount rate, P the analysis period, and multipliers c (cost), d (closure days), s (salvage) and f (share of daily revenue lost):'));
children.push(mono('NPW = Initial x c', 18));
children.push(mono('    + sum over activities with year <= P of [ (DirectCost x c) + (ClosureDays x d x DailyRevenue x f) ] / (1 + r)^year', 18));
children.push(mono('    + SalvageBase x c x s / (1 + r)^P', 18));
children.push(p([r('SalvageBase = -fraction x basis, where basis is initial construction (PCC, 25 percent) or the mill-and-overlay cost (HMA, 12.5 percent) per the TDOT Maintenance Policies. With all multipliers at 1, r = 3, P = 30 this reproduces the framework NPW to the cent. EUAC = NPW x r(1+r)^P / ((1+r)^P - 1).')], { before: 120 }));
children.push(h2('8.2 Inputs'));
children.push(table([
  ['Cell(s)', 'Input', 'MBT value', 'Source'],
  ['B5', 'Discount rate (%)', '3', 'General Information D34'],
  ['B6', 'Analysis period (years)', '30', 'General Information D33'],
  ['B7', 'Construction year', '2027', 'General Information D25'],
  ['B8', 'Airport daily revenue ($/day)', '8,841.38', 'General Information D39 / Alt sheet F2'],
  ['B9', 'Share of daily revenue lost during a closure', '100%', 'Framework method (gross receipts); test 30 to 50 percent for margin-based revenue'],
  ['B10', 'Closure-days multiplier', '1.0', 'Scales every closure duration'],
  ['B11, B12', 'HMA and PCC cost multipliers', '1.0', 'Bid-price uncertainty on all direct costs'],
  ['B13', 'Salvage multiplier', '1.0', '0 removes the salvage credit'],
  ['B14', 'Mainline area (SY)', '52,777.7', 'General Information D26'],
  ['B15:B18', 'Salvage fraction and basis per alternative', '0.125 Rehabilitation; 0.25 Initial', 'Maintenance Policies tables 1 and 2'],
  ['Activity tables C:D and L:M', 'Base direct cost and closure days per policy activity', 'from the run', 'Alt sheet NPW table column D (Actual Cost) and cells F4:F10'],
], [1500, 2900, 1900, 3060], { size: 15 }));
children.push(p(''));
children.push(h2('8.3 Blocks and charts'));
for (const t of [
  'RESULTS: initial cost, maintenance, rehabilitation, lost revenue and salvage present worth, NPW, EUAC, NPW per SY, closure days in period, runway availability, undiscounted lost revenue; the lower-cost alternative, its margin in dollars and percent, whether the same alternative wins at 7 percent, and the probability that PCC is lower from the live simulation.',
  'BREAK-EVEN VALUES: the daily revenue, PCC bid level, HMA bid level and PCC salvage fraction at which the two alternatives tie, in closed form (the NPW difference is linear in each), and the first discount rate in the 2 to 8 percent sweep at which the winner changes.',
  'DISCOUNT-RATE SENSITIVITY: NPW of each alternative at 25 rates from 2 to 8 percent.',
  'EXPENDITURE BY CALENDAR YEAR: direct cost and lost revenue by year for each alternative, cumulative discounted cost, closure days by year, and the crossover year of the cumulative lines.',
  'WHAT COULD FLIP THE ANSWER: PCC minus HMA present worth when one input moves across its range (rate 2 to 8 percent, each salvage credit off, each cost plus or minus 20 percent, closures x0.5 to x2, lost revenue off, 20-year period), with a flag where the sign changes.',
  'SCENARIO SCORECARD: twelve pre-run scenarios with their own parameters (TDOT policy, FAA AIP 7 percent and 20 years, 3 percent and 20 years, 5 percent, no lost revenue, 40 percent revenue loss, no salvage, PCC bids +20, HMA bids +20, closures x2, closures x0.5, and a stress case).',
  'DECISION MAP: PCC minus HMA present worth on a grid of discount rate (2 to 8 percent) by salvage multiplier (0 to 1), color-scaled so the winner is read at a glance.',
  'MONTE CARLO (column AK onward): 1,000 joint draws from editable triangular ranges on rate, both cost multipliers, closure days, salvage and revenue share; probability that PCC is lower, mean, P10, P50, P90 and a histogram. RAND() based, so F9 redraws and the probability moves by a percent or two.',
  'Charts (column T onward): present worth by category (stacked), expenditure stream by year, cumulative discounted cost, NPW versus discount rate, tornado, closure timeline (bubble, size = days), closure days by year, NPW by scenario, simulation histogram.',
]) children.push(bullet(t));
children.push(h2('8.4 Adding a project'));
children.push(num('Right-click the TEMPLATE sheet tab, Move or Copy, tick Create a copy, rename the sheet.'));
children.push(num('Fill B5:B18 from the framework General Information sheet and the Maintenance Policies.'));
children.push(num('For each alternative, type the base direct cost and closure days of each policy activity into the yellow cells of the activity table, taken from the alternative sheet\'s NPW table (column D) and cells F4:F10.'));
children.push(num('Every table and chart updates. The sheet compares one HMA and one PCC alternative, which is how the framework is used; policy years are editable if the maintenance tables change.'));
children.push(p([r('Note on the loaded runs. ', { bold: true }), r('The MBT and SRB values are the framework outputs as reported on 21 May 2026, that is, without engineering on initial construction (Section 5.2). Re-running those projects in v1.2.0 would change the initial costs as stated there.')]));
children.push(image('MBT_bridge.png', 6.5, 4.14));
children.push(caption('Figure 1. MBT present-worth bridge from HMA to PCC by category (decision workbook view).'));
children.push(image('MBT_closures.png', 6.5, 6.44));
children.push(caption('Figure 2. MBT runway closure timeline; bubble size is days closed.'));

// 9 verification
children.push(h1('9. Verification record'));
children.push(table([
  ['Check', 'Method', 'Result'],
  ['v1.2.0 workbook integrity', 'Zip test; every XML, rels and VML part parsed', 'Pass; 463 cells changed across 7 sheets, 1 validation, 5 text runs, 5 chart parts, 1 comments part'],
  ['v1.2.0 full recalculation', 'LibreOffice Calc calculateAll, then scan of every cell for error values', '2,240 formulas including the new Summary, 0 errors'],
  ['New project run end to end (re-run after the design pass)', 'CKV Runway 17-35 (Outlaw Field) built on the blank v1.2.0 template through the UNO API: General Information, both alternatives, pay items (verification/new_project_CKV/run_ckv.py)', 'Framework Summary, the decision workbook TEMPLATE sheet and an independent Python engine give the same NPW ($9,829,224 HMA and $9,180,602 PCC) and the same category present worths; closure days, salvage and daily revenue match hand calculation; no error cells'],
  ['Two worked examples, three and four alternatives', 'GKT Runway 10-28 (no shoulders, HMA / 9 in PCC / 11 in PCC) and MKL Runway 2-20 (25 ft shoulders, two HMA / two PCC) built on the blank template through the UNO API with quantities derived from a target section (verification/examples/run_example.py)', 'Item costs, initial totals, every activity present worth, NPW, category sums, closure days, cumulative cost, verdict, vs. lowest, EUAC, the 3 percent point of the sensitivity curve and the section read-back all agree with an independent calculation: 43 checks on GKT and 61 on MKL, no failures, no error cells. GKT: 11 in PCC lowest at $5,852,247, HMA lowest above about 4 percent. MKL: 9 in PCC lowest at $14,066,568 at every rate; the shoulder reading returns the designed 8/6/6, 9/6/6, 10/10 and 11/6 in sections'],
  ['Pavement section read-back', 'Derived thickness compared with the sections recorded in the two completed runs and with a hand calculation', 'Murfreesboro reads back as 5.00 in P-401 on 16.00 in P-209, exactly the section typed in its description, and its excavation quantity agrees to the inch; 145 pcf is the unit weight that reproduces it. The concrete alternative reads 9 in P-501 on 6 in P-209 against an excavation quantity that matches 8 in, which is the mismatch the check is meant to surface. On the CKV run the shoulder chart populates only after a shoulder area is entered, and then scales every derived layer by mainline over mainline plus shoulder'],
  ['Typical Values sheet', 'Full recalculation of the template; every live row read back', 'No error cells; Pay_Items defaults, Maintenance Policies years and salvage fractions and the 17 daily revenues resolve; the sheet is last so no sheet index or print area shifted'],
  ['New Summary on the MBT data', 'Summary transplanted into the MBT workbook; full recalculation; rendered to PDF', 'Table reads NPW $8,809,266 and $8,028,734, categories sum to NPW, closure days 57 (HMA) and 9 for PCC in that file (its rehabilitation indirect row is mislabelled, see Appendix C; the v1.2.0 template gives 27), flags the winner change at 7 percent; six charts parse in an independent reader and render; on the empty template the sheet reads "No alternatives yet" with no error cells'],
  ['v1.1.2 under the same recalculation', 'Same', '99 error cells: 39 #NAME? (XLOOKUP) and 60 #N/A (unguarded PCC lookups, revenue lookup)'],
  ['Template formulas on real data', 'Formula changes applied to the MBT workbook alternative sheets; full recalculation', 'NPW unchanged: HMA $8,809,266.42, PCC $8,028,734.49; PCC rows 17 to 22 = 0 (were #N/A)'],
  ['Year-indexed chart columns', 'Sum of L + M against activity totals; sum of N against NPW; per-year values against policy years', 'HMA: L+M = $10,959,044 = totals; N = $8,809,266 = NPW. PCC: L+M = $7,281,371; N = $8,028,734. Appendix D'],
  ['Chart parts', 'Parsed with an independent chart reader', 'Category K37:K67 (or K38:K68) and value ranges as intended; indirect charts stacked with two series'],
  ['Decision workbook recalculation', 'LibreOffice Calc calculateAll; error scan', '47,265 formulas, 0 errors'],
  ['Decision workbook against independent engine', 'Python implementation of Section 8.1', 'NPW, category split (sums to NPW), EUAC, closure days, break-even values (engine returns a $0 difference at each), 25-point sensitivity, tornado, 12 scenarios and 35-cell decision map agree to the dollar'],
  ['Decision workbook charts', 'Rendered to PDF and inspected', 'All eight per-project charts and the simulation histogram draw with the intended series, axes and years'],
  ['Scripted click-through', 'LibreOffice UNO API: open both workbooks, follow every button target, change discount rate, analysis period, indirect-cost flag and airport, remove alternatives (verification/uno_walkthrough.py)', 'All three buttons land on visible sheets; 7 percent flips the verdict to Alternative 1; 20 years drops HMA closures to 43 days and PCC rehabilitation to $0; an unknown airport gives $0 lost revenue with the G2 warning and no error cells; one alternative and none read cleanly'],
  ['Scripted click-through, round 2', 'Emulate the VBA Database row deletion; tie; four alternatives; airport dropdown; pay-item uniqueness; page setup (verification/uno_walkthrough_2.py, _3_delete_tie.py)', 'Summary shifts up with no #REF!; tie reads "tied with the next alternative"; four rows and all chart blocks populate; dropdown has 79 airports, no blanks or duplicates; 56 unique pay-item descriptions; Summary prints landscape one page wide'],
  ['Package audit', 'verification/audit_xml.py: cell and row order, style and shared-string counts, relationships, content types, defined names, selected tabs', 'No issues; one selected tab (General Information); v1.1.2 shows the calcChain and three external Airport_Name names'],
  ['Not verified in this environment', '', 'Display of the alternative-sheet charts in Excel (LibreOffice does not draw charts on the ActiveX-bearing sheets). Closed by Section 12.'],
], [2300, 3000, 4060], { size: 15 }));

// 10 what is different in one place
children.push(h1('10. What is different, in one place'));
children.push(p('For a reader who needs only the list:'));
for (const t of [
  'Numbers that change for every analysis: initial construction now includes pay item 1 (non-indirect templates) and the engineering percent (all templates).',
  'Numbers that change only when the analysis period is shorter than the policy schedule: activities after the period drop out.',
  'Numbers that change from #N/A to a value: any analysis with lost revenue for an airport outside the 17; any PCC alternative with an empty pay-item row.',
  'Defaults that change: analysis period 30 instead of 10; closure days pre-filled instead of 0.',
  'Lookups that change: the four geotextile lines now resolve to their own item numbers.',
  'Presentation that changes: alternative charts by calendar year, stacked direct and lost revenue; Summary with categories, closures, rate sensitivity, five charts and navigation buttons, all formula-driven.',
  'Text that changes: Instructions runs 2, 4, 5, 7 and 8; two cell notes.',
  'Removed: two external links, three sheet-scoped names, the cached calculation chain, six printer-settings parts, the Power Query stub, the cached SharePoint path and the personal names in the document properties.',
  'Not changed: Maintenance Policies, Pay_Items unit costs, RevenueData, salvage fractions, the discount formula for in-period activities, the VBA that creates alternatives, the ActiveX controls, every other sheet.',
]) children.push(bullet(t));

// 11 policy
children.push(h1('11. Policy questions surfaced, not decided'));
children.push(table([
  ['Question', 'Why it matters', 'Where it is visible now'],
  ['Salvage asymmetry', 'PCC recovers 25 percent of total initial cost including mobilization; HMA 12.5 percent of one mill-and-overlay. At Murfreesboro the rule decides the answer: PCC wins by $780,532 as the workbook stands, and on a common HMA rule it becomes HMA by $63,633. At Upper Cumberland the year-30 credit is -$3.86M against -$0.41M, several times the $586K margin, but HMA wins under every variant, so there it moves the margin and not the winner.', 'Decision workbook RESULTS (Salvage PW), tornado, decision map, salvage multiplier input'],
  ['Salvage when the period is shortened', 'The policy fractions are defined at year 30. At 20 years the workbook applies the same fraction at year 20, a simplification that under-credits PCC.', 'Scenario scorecard rows "FAA AIP" and "TDOT 3% but 20-yr life"'],
  ['Lost-revenue basis', 'The framework sums gross fuel sales, tie-downs, hangar storage, flowage fees and tenant revenue as lost on every closure day. A margin-based share would be far lower.', 'Input B9 (share lost); scenario "Revenue loss at 40% of gross"'],
  ['FAA AIP discount rate', 'PGL 22-01 (June 2022) replaced the fixed 7 percent and 20-year rule with the OMB A-94 Appendix C real rate for the chosen period: 2.0 percent for calendar 2026. MBT changes winner at 4.5 percent, so the rule that applies decides the answer.', 'Sensitivity block; the verdict line now flags both 2 percent and 7 percent; FAA scenario in the decision workbook'],
  ['Engineering on initial construction', 'Section 5.2 applies it; Aeronautics may prefer the previous treatment.', 'General Information D37; template row 26'],
  ['RevenueData hygiene', 'MBT and MQY rows are identical; XNX and M54 hold text; 62 airports have no data.', 'Warning in template G2; D39 message'],
  ['Overview wording', 'NS rewrite pending approval.', 'Overview text box'],
], [1900, 4200, 3260], { size: 15 }));

// 12 checklist
children.push(h1('12. First-open checklist in Excel'));
for (const t of [
  'Open TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm; enable content (or follow the ActiveX steps now in the Instructions). Nothing to import.',
  'Summary: expect "No alternatives yet", five empty charts, and two navigation buttons that jump to General Information and Instructions.',
  'General Information: select Murfreesboro Municipal Airport, D38 = Yes, D25 = 2027, D26 = 52777.7, D28 = 5433. Alternative Setup: add one New HMA and one New PCC; Close. Click "View Summary". Expect the table, verdict line and all charts to populate with no #N/A.',
  'On Alt 1: enter the MBT pay items and quantities. Expect Initial Construction $5,547,785 (the reported $5,306,577 plus 5 percent engineering on the $4,824,160 subtotal). Expect columns L:N to sum to the NPW table and the chart axis to read 2027 to 2057.',
  'Set D33 = 20. Expect Maintenance 5 and 6 to show $0 discounted and the chart to end at 2047. Restore 30.',
  'Select an airport outside the 17 with D38 = Yes. Expect G2 on each Alt sheet to show the warning and lost revenue $0, not #N/A.',
  'Open TDOT_LCCA_Decision_Workbook.xlsx; on MBT confirm NPW $8,809,266 and $8,028,734; press F9 twice and confirm the simulation probability moves only slightly around 52 percent.',
]) children.push(num(t));

// Appendix A
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h1('Appendix A. Cell-level change list'));
children.push(p('Every cell whose formula or value differs between v1.1.2 and v1.2.0, generated by comparing the two workbooks cell by cell (463 cells). Runs of identical formulas in the chart helper columns L, M and N are shown once with the first row; the row number substitutes down the column.'));
for (const [sheet, entries] of Object.entries(diff.diff)) {
  children.push(h2(`A.${Object.keys(diff.diff).indexOf(sheet) + 1} ${sheet} (${entries.length} cells)`));
  const rows = sheet === 'Summary'
    ? [['Cell', 'v1.1.2', 'v1.2.0'], ['Whole sheet', 'Header row A3:E3 and one chart', `Rebuilt: ${entries.length} cells of formulas, labels and buttons (Section 7, Appendix B); A3:E3 unchanged; rows 4 to 7 of A:E left to the form`]]
    : [['Cell', 'v1.1.2', 'v1.2.0'], ...groupSheet(sheet, entries)];
  children.push(table(rows, [1300, 3400, 4660], { monoCols: [1, 2], size: 13 }));
  children.push(p(''));
}
children.push(h2('A.8 Other parts'));
children.push(table([
  ['Part', 'Change'],
  ['General Information data validation D9', '$L$9:$L$84 to $L$10:$L$88'],
  ['xl/comments1.xml and vmlDrawing2.vml', 'Author "Ebenezer Duah" replaced by "ARA" in the D26 and D39 notes'],
  ['xl/sharedStrings.xml', 'Four strings appended for the geotextile descriptions'],
  ['xl/charts/chart1.xml to chart5.xml', 'Series reduced to Actual (direct templates) or Direct + Indirect stacked (indirect templates); category reference K37:K67 or K38:K68 added; series extension lists and caches removed; category axis number format "0"'],
  ['xl/drawings/drawing3.xml', 'Instructions text runs 2, 4, 5, 7, 8 (Section 6)'],
  ['xl/workbook.xml', 'externalReferences element and three sheet-scoped Airport_Name names removed; calcPr fullCalcOnLoad="1"; print area defined for Summary'],
  ['xl/worksheets/sheet14.xml (Summary), xl/drawings/drawing10.xml, xl/charts/chart7.xml to chart11.xml, xl/styles.xml', 'Summary sheet rebuilt (formula blocks, buttons, page setup); five chart parts and their anchors added to the existing drawing; the styles the new cells use appended to the style table'],
  ['xl/externalLinks/*, xl/calcChain.xml', 'Removed with their relationships and content-type overrides'],
], [3200, 6160], { size: 15 }));

// Appendix B
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h1('Appendix B. Summary sheet formula reference'));
children.push(p('Formulas as written in the v1.2.0 Summary sheet for the first alternative row (row 4) and the first row of each chart-data block. Rows 5 to 7 and the following rows of each block repeat them with the row number substituted.'));
children.push(table([['Cell', 'Formula'], ...SUMMARY_FORMULAS.map(x => [x[0], x[1]])], [1200, 8160], { monoCols: [1], size: 12 }));

// Appendix C reconciliation
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h1('Appendix C. MBT reconciliation of the year-indexed chart columns'));
children.push(p('Values from the recalculated MBT copy after the template formulas were applied (calendar year, direct cost, lost revenue, total discounted). Rows with no activity are zero and omitted.'));
children.push(table([
  ['Alt 1 New HMA', 'Direct (L)', 'Lost revenue (M)', 'Discounted (N)'],
  ['2027', '5,306,577', '0', '5,306,577'], ['2031', '482,830', '53,048', '476,121'], ['2035', '482,830', '53,048', '423,027'], ['2039', '547,129', '70,731', '433,355'],
  ['2043', '561,004', '70,731', '393,677'], ['2047', '2,302,558', '132,621', '1,348,299'], ['2051', '482,830', '53,048', '263,617'], ['2055', '577,147', '70,731', '283,172'], ['2057', '-287,820', '0', '-118,578'],
  ['Sum', '10,455,085', '503,959', '8,809,266 = NPW'],
], [2340, 2340, 2340, 2340], { size: 15 }));
children.push(p(''));
children.push(table([
  ['Alt 2 New PCC', 'Direct (L)', 'Lost revenue (M)', 'Discounted (N)'],
  ['2027', '8,410,245', '0', '8,410,245'], ['2046', '306,564', '79,572', '220,208'], ['2054', '587,551', '0*', '264,509'], ['2057', '-2,102,561', '0', '-866,227'],
  ['Sum', '7,201,798', '79,572', '8,028,734 = NPW'],
], [2340, 2340, 2340, 2340], { size: 15 }));
children.push(p([r('* In the MBT file the PCC rehabilitation indirect row is labelled "Rehabilitation 1" rather than "Rehabilitation 1 Indirect Cost", so its $159,145 is counted as direct there and its 18 closure days are not counted by the Summary in that file. The v1.2.0 template label is correct.', { size: 16, italics: true })], { before: 80 }));

// ------------------------------------------------------------------ document
const doc = new Document({
  creator: 'Applied Research Associates, Inc.', title: 'TDOT LCCA Framework Technical Change Record v1.2.0', features: { updateFields: true },
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: 'Cambria', size: 30, bold: true, color: '1D2733' }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: 'Cambria', size: 25, bold: true, color: '1D2733' }, paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: 'Cambria', size: 22, bold: true }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: new Header({ children: [p([r('TDOT Aeronautics LCCA Framework, Technical Change Record v1.1.2 to v1.2.0', { size: 16, color: '7F7F7F' })], { after: 0 })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Applied Research Associates, Inc.  |  Page ', font: FONT, size: 16, color: '7F7F7F' }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '7F7F7F' })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(__dirname + '/TDOT_LCCA_v1.2.0_Technical_Change_Record.docx', buf); console.log('wrote docx', buf.length); });
