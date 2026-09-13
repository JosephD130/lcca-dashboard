#!/usr/bin/env python3
"""Surgical XML patch for TDOA_LCCA_Framework v1.1.2 -> v1.2.0.
Edits sheet XML in place inside the .xlsm zip so VBA, ActiveX comboboxes, charts,
tables and data validations are preserved (openpyxl would drop them).
"""
import re, html, sys, os, zipfile, shutil
from build_summary import transplant, add_plain_sheet, add_style, add_dxf, FONT, FILL, BORDER_BOTTOM, BORDER_BOX
import build_helpers, helpers_content, build_method
build_helpers.set_sections(helpers_content.sections()); build_helpers.HINTS.update(helpers_content.HINTS)

GI = "'General Information'"

def colnum(col):
    n = 0
    for ch in col: n = n*26 + (ord(ch)-64)
    return n

def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)$', ref); return m.group(1), int(m.group(2))

def esc(f):  # XML-escape a formula for <f>
    return f.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

def get_row(x, r):
    m = re.search(r'<row r="%d"[^>]*>.*?</row>|<row r="%d"[^>]*/>' % (r, r), x, re.S)
    return m

def cell_re(ref):
    # the attribute run is lazy: a self-closing cell ends at its own "/>", not at the next cell's "</c>"
    return re.compile(r'<c r="%s"(?:\s[^>]*?)?(?:/>|>.*?</c>)' % ref, re.S)

def remove_cell(x, ref):
    m = get_row(x, split_ref(ref)[1]); assert m, ref
    row = m.group(0)
    cm = cell_re(ref).search(row)
    if not cm: return x
    assert 't="shared"' not in cm.group(0) or 'ref=' not in cm.group(0), "shared master removed: "+ref
    return x[:m.start()] + row[:cm.start()] + row[cm.end():] + x[m.end():]

def put_cell(x, ref, cellxml):
    """Replace cell if present else insert in column order inside its row."""
    col, r = split_ref(ref)
    m = get_row(x, r); assert m, "row %d missing" % r
    row = m.group(0)
    cm = cell_re(ref).search(row)
    if cm:
        row2 = row[:cm.start()] + cellxml + row[cm.end():]
    else:
        if row.endswith('/>'):
            row2 = row[:-2] + '>' + cellxml + '</row>'
        else:
            # find first cell with greater column
            pos = None
            for c in re.finditer(r'<c r="([A-Z]+)(\d+)"', row):
                if colnum(c.group(1)) > colnum(col): pos = c.start(); break
            if pos is None: pos = row.rfind('</row>')
            row2 = row[:pos] + cellxml + row[pos:]
    return x[:m.start()] + row2 + x[m.end():]

def fcell(ref, formula, style=None, t=None, v='0'):
    s = ' s="%s"' % style if style is not None else ''
    tt = ' t="%s"' % t if t else ''
    return '<c r="%s"%s%s><f>%s</f><v>%s</v></c>' % (ref, s, tt, esc(formula), v)

def style_of(x, ref):
    cm = cell_re(ref).search(x)
    if not cm: return None
    sm = re.search(r'\ss="(\d+)"', cm.group(0)); return sm.group(1) if sm else None

# ------------------------------------------------------------------ template patch
def patch_template(x, *, init, fa, la, kstart, kend, indirect, pcc_guard=False, fix_g24=True,
                   add_engineering=True, durations=None, fix_f2=False, hmarehab=False):
    # A. initial-construction subtotal must include pay item 1
    if fix_g24:
        x = x.replace('<f>SUM(G14:G23)</f>', '<f>SUM(G13:G22)</f>')
    # B. engineering on initial construction (row 26 was blank; G27 already sums G24:G26)
    if add_engineering:
        st = style_of(x, 'G26')
        x = put_cell(x, 'B26', '<c r="B26" t="s"><v>323</v></c>')  # shared string 323 = "Engineering"
        x = put_cell(x, 'G26', fcell('G26', "(%s!$D$37/100)*G24" % GI, st))
    # C. analysis-period guard on every discounted-cost row
    x = re.sub(r"<f>D(\d+)/\(1\+\('General Information'!\$D\$34/100\)\)\^C\1</f>",
               lambda m: '<f>' + esc("IF(C%s>%s!$D$33,0,D%s/(1+(%s!$D$34/100))^C%s)" % (m.group(1), GI, m.group(1), GI, m.group(1))) + '</f>',
               x)
    # F. blank guard on pay-item lookups (PCC templates rows 14-22 lacked it)
    if pcc_guard:
        x = re.sub(r'<f>VLOOKUP\(C(\d+),CHOOSE\(\{1,2\},Table2\[Pay Item Description\],Table2\[Pay Item No\.\]\),2,0\)</f>',
                   lambda m: '<f>IF(C%s="",0,VLOOKUP(C%s,CHOOSE({1,2},Table2[Pay Item Description],Table2[Pay Item No.]),2,0))</f>' % (m.group(1), m.group(1)), x)
    # G. HMA rehab template latent bugs
    if hmarehab:
        x = x.replace("<f>C36+'Maintenance Policies'!E65</f>", "<f>'Maintenance Policies'!E65</f>")
        x = x.replace("<f>-'Maintenance Policies'!D71*AV22</f>", "<f>-'Maintenance Policies'!D71*AN22</f>")
    # D. revenue lookup without XLOOKUP, tolerant of text cells, guarded against missing airport
    if fix_f2:
        terms = '+'.join("SUMIF(RevenueData!$A:$A,%s!$D$10,RevenueData!$%s:$%s)" % (GI, c, c) for c in 'BCDEFGH')
        x = put_cell(x, 'F2', fcell('F2', "IF(%s!$D$38=\"Yes\",%s,0)" % (GI, terms), style_of(x, 'F2')))
        x = put_cell(x, 'G2', fcell('G2', 'IF(AND(%s!$D$38="Yes",F2=0),"No revenue data for this airport - lost revenue set to $0. Contact Aeronautics.","")' % GI, None, 'str', ''))
    # H. default closure durations (production rates carried over from the MBT/SRB analyses)
    if durations:
        for ref, f in durations.items():
            x = put_cell(x, ref, fcell(ref, f, style_of(x, ref) or '112'))
    # E. year-indexed chart columns (replaces hard-positioned L/M/N rows)
    stL = style_of(x, 'L%d' % kstart) or '63'
    for r in range(kstart, kend + 1):
        for col in 'LMN':
            x = remove_cell(x, '%s%d' % (col, r))
        K = 'K%d' % r
        guard = "IF(%s-$C$6>%s!$D$33,0," % (K, GI)
        rngC = "$C$%d:$C$%d" % (fa, la); rngD = "$D$%d:$D$%d" % (fa, la); rngE = "$E$%d:$E$%d" % (fa, la); rngB = "$B$%d:$B$%d" % (fa, la)
        if indirect:
            L = guard + "IF(%s=$C$6,$D$%d,0)+SUMIFS(%s,%s,%s-$C$6,%s,\"<>*Indirect*\"))" % (K, init, rngD, rngC, K, rngB)
            M = guard + "SUMIFS(%s,%s,%s-$C$6,%s,\"*Indirect*\"))" % (rngD, rngC, K, rngB)
            N = guard + "IF(%s=$C$6,$E$%d,0)+SUMIF(%s,%s-$C$6,%s))" % (K, init, rngC, K, rngE)
            cells = [('L', L), ('M', M), ('N', N)]
        else:
            L = guard + "IF(%s=$C$6,$D$%d,0)+SUMIF(%s,%s-$C$6,%s))" % (K, init, rngC, K, rngD)
            M = guard + "IF(%s=$C$6,$E$%d,0)+SUMIF(%s,%s-$C$6,%s))" % (K, init, rngC, K, rngE)
            cells = [('L', L), ('M', M)]
        for col, f in cells:
            x = put_cell(x, '%s%d' % (col, r), fcell('%s%d' % (col, r), f, stL))
    return x

# ------------------------------------------------------------------ chart patch
AX_TITLE = ('<c:title><c:tx><c:rich><a:bodyPr%s/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0"/></a:pPr>'
            '<a:r><a:rPr lang="en-US" sz="900" b="0"/><a:t>%s</a:t></a:r></a:p></c:rich></c:tx>'
            '<c:overlay val="0"/></c:title>')


def axis_titles(c, xt, yt):
    """Give a chart an x and a y axis title. Inserted before <c:numFmt> so the element order the schema
    requires (axId, scaling, delete, axPos, gridlines, title, numFmt, ...) is kept."""
    for tag, text, body in (('c:catAx', xt, ''), ('c:dateAx', xt, ''), ('c:valAx', yt, ' rot="-5400000" vert="horz"')):
        m = re.search(r'<%s>.*?</%s>' % (tag, tag), c, re.S)
        if not m or '<c:title>' in m.group(0): continue
        ax = m.group(0)
        anchor = ax.find('<c:numFmt')
        if anchor < 0: anchor = ax.find('<c:majorTickMark')
        if anchor < 0: continue
        ax = ax[:anchor] + (AX_TITLE % (body, html.escape(text, quote=False))) + ax[anchor:]
        c = c[:m.start()] + ax + c[m.end():]
    return c

def patch_chart(c, sheet, kstart, kend, keep_series):
    sers = re.findall(r'<c:ser>.*?</c:ser>', c, re.S)
    assert sers, 'no series'
    new = []
    for i, s in enumerate(sers[:keep_series]):
        s = re.sub(r'<c:cat>.*?</c:cat>', '', s, flags=re.S)
        if '<c:extLst>' in s:  # series-level extLst is the last child; nested extLsts inside it defeat a non-greedy regex
            s = s[:s.find('<c:extLst>')] + '</c:ser>'
        s = re.sub(r'<c:numCache>.*?</c:numCache>', '', s, flags=re.S)
        s = re.sub(r'<c:idx val="\d+"/><c:order val="\d+"/>', '<c:idx val="%d"/><c:order val="%d"/>' % (i, i), s)
        cat = "<c:cat><c:numRef><c:f>'%s'!$K$%d:$K$%d</c:f></c:numRef></c:cat>" % (sheet, kstart, kend)
        s = s.replace('<c:val>', cat + '<c:val>', 1)
        new.append(s)
    first = c.find('<c:ser>'); last = c.rfind('</c:ser>') + len('</c:ser>')
    c = c[:first] + ''.join(new) + c[last:]
    if keep_series > 1:
        c = c.replace('<c:grouping val="clustered"/>', '<c:grouping val="stacked"/>')
        c = re.sub(r'<c:overlap val="-?\d+"/>', '<c:overlap val="100"/>', c)
    # category axis: show every 5th year, no decimals
    c = re.sub(r'(<c:catAx>.*?)<c:numFmt formatCode="General" sourceLinked="1"/>', r'\1<c:numFmt formatCode="0" sourceLinked="0"/>', c, count=1, flags=re.S)
    return axis_titles(c, 'Calendar year', 'Cost ($)')

# ------------------------------------------------------------------ main
def main(src, out, mbt_mode=False):
    work = out + '.work'
    if os.path.isdir(work): shutil.rmtree(work)
    with zipfile.ZipFile(src) as z: z.extractall(work)
    P = lambda p: os.path.join(work, p)
    def rd(p): return open(P(p), encoding='utf-8').read()
    def wr(p, s): open(P(p), 'w', encoding='utf-8').write(s)

    HMA_DUR = {
        'F4': "ROUNDUP((C5/15000),0)+ROUNDUP((C8/20000),0)+ROUNDUP((C8/50000),0)",
        'F5': "ROUNDUP((C5/15000),0)+ROUNDUP((C8/20000),0)+ROUNDUP((C8/50000),0)",
        'F6': "ROUNDUP((C5/15000),0)+ROUNDUP((C8/20000),0)+ROUNDUP((C8/50000),0)+ROUNDUP((C5*'Maintenance Policies'!D15/10000),0)+ROUNDUP((C5*'Maintenance Policies'!D14/5000),0)",
        'F7': "ROUNDUP((C5/15000),0)+ROUNDUP((C8/20000),0)+ROUNDUP((C8/50000),0)+ROUNDUP((C5*'Maintenance Policies'!D18/10000),0)+ROUNDUP((C5*'Maintenance Policies'!D19/5000),0)",
        'F8': "ROUNDUP((C5/3800),0)+ROUNDUP((C8/50000),0)",
        'F9': "ROUNDUP((C5/15000),0)+ROUNDUP((C8/20000),0)+ROUNDUP((C8/50000),0)",
        'F10': "ROUNDUP((C5/15000),0)+ROUNDUP((C8/20000),0)+ROUNDUP((C8/50000),0)+ROUNDUP((C5*'Maintenance Policies'!D28/10000),0)+ROUNDUP((C5*'Maintenance Policies'!D29/5000),0)",
    }
    PCC_DUR = {
        'F4': "ROUNDUP(C5*550/(50*100/9)/10000,0)+ROUNDUP(C5*'Maintenance Policies'!D38/10000,0)+ROUNDUP(C5*'Maintenance Policies'!D39/5000,0)+ROUNDUP(C5*'Maintenance Policies'!D40/5000,0)",
        'F5': "ROUNDUP(C5*550/(50*100/9)/10000,0)+ROUNDUP(C5*'Maintenance Policies'!D44/10000,0)+ROUNDUP(C5*'Maintenance Policies'!D43/5000,0)+ROUNDUP(C5*'Maintenance Policies'!D42/5000,0)+ROUNDUP(C5*'Maintenance Policies'!D41/1000,0)+7",
    }

    if mbt_mode:
        # verification only: MBT Alt 1 (New HMA) = sheet14 (indirect HMA layout), Alt 2 (New PCC) = sheet15 (indirect PCC layout)
        x = rd('xl/worksheets/sheet14.xml'); x = patch_template(x, init=36, fa=37, la=51, kstart=37, kend=67, indirect=True, fix_g24=False, add_engineering=False); wr('xl/worksheets/sheet14.xml', x)
        x = rd('xl/worksheets/sheet15.xml'); x = patch_template(x, init=37, fa=38, la=42, kstart=38, kend=68, indirect=True, pcc_guard=True, fix_g24=False, add_engineering=False); wr('xl/worksheets/sheet15.xml', x)
        for ch, sheet, ks, ke, keep in [('chart6', 'Alt 1 (New HMA)', 37, 67, 2), ('chart7', 'Alt 2 (New PCC)', 38, 68, 2)]:
            c = rd('xl/charts/%s.xml' % ch); c = patch_chart(c, sheet, ks, ke, keep); wr('xl/charts/%s.xml' % ch, c)
    else:
        # --- templates
        x = rd('xl/worksheets/sheet10.xml'); x = patch_template(x, init=36, fa=37, la=44, kstart=37, kend=67, indirect=False); wr('xl/worksheets/sheet10.xml', x)
        x = rd('xl/worksheets/sheet8.xml');  x = patch_template(x, init=36, fa=37, la=51, kstart=37, kend=67, indirect=True, fix_g24=False, durations=HMA_DUR, fix_f2=True); wr('xl/worksheets/sheet8.xml', x)
        x = rd('xl/worksheets/sheet11.xml'); x = patch_template(x, init=37, fa=38, la=40, kstart=38, kend=68, indirect=False, pcc_guard=True); wr('xl/worksheets/sheet11.xml', x)
        x = rd('xl/worksheets/sheet1.xml');  x = patch_template(x, init=37, fa=38, la=42, kstart=38, kend=68, indirect=True, pcc_guard=True, fix_g24=False, durations=PCC_DUR, fix_f2=True); wr('xl/worksheets/sheet1.xml', x)
        x = rd('xl/worksheets/sheet12.xml'); x = patch_template(x, init=36, fa=37, la=44, kstart=37, kend=67, indirect=False, hmarehab=True); wr('xl/worksheets/sheet12.xml', x)
        # --- charts (expenditure stream: undiscounted, calendar years on the axis)
        for ch, sheet, ks, ke, keep in [('chart3', 'TMP(NewHMA)', 37, 67, 1), ('chart2', 'TMP(NewHMA)_IndirectCost', 37, 67, 2),
                                        ('chart4', 'TMP(NewPCC)', 38, 68, 1), ('chart1', 'TMP(NewPCC)_IndirectCost', 38, 68, 2),
                                        ('chart5', 'TMP(HMARehab)', 37, 67, 1)]:
            c = rd('xl/charts/%s.xml' % ch); c = patch_chart(c, sheet, ks, ke, keep); wr('xl/charts/%s.xml' % ch, c)
        # --- General Information
        x = rd('xl/worksheets/sheet4.xml')
        x = x.replace('<c r="D33" s="88"><v>10</v></c>', '<c r="D33" s="88"><v>30</v></c>')
        x = x.replace('<formula1>$L$9:$L$84</formula1>', '<formula1>$L$10:$L$88</formula1>')
        terms = '+'.join("SUMIF(RevenueData!$A:$A,$D$10,RevenueData!$%s:$%s)" % (c, c) for c in 'BCDEFGH')
        x = put_cell(x, 'D39', fcell('D39', 'IF(D38<>"Yes","",IF(COUNTIF(RevenueData!$A:$A,$D$10)=0,"Missing Airport Revenue",%s))' % terms, '114', 'str', ''))
        wr('xl/worksheets/sheet4.xml', x)
        # --- comments: no personal names
        c = rd('xl/comments1.xml').replace('Ebenezer Duah:', 'ARA:').replace('Ebenezer Duah\n', 'ARA\n').replace('Ebenezer Duah', 'ARA'); wr('xl/comments1.xml', c)
        v = rd('xl/drawings/vmlDrawing2.vml'); wr('xl/drawings/vmlDrawing2.vml', v.replace('Ebenezer Duah', 'ARA'))
        # --- Pay_Items: make the four geotextile descriptions unique (VLOOKUP keys on description)
        ss = rd('xl/sharedStrings.xml')
        m = re.search(r'<sst[^>]*count="(\d+)" uniqueCount="(\d+)"', ss); cnt, uc = int(m.group(1)), int(m.group(2))
        newstr = ['Separation Geotextile (P-154)', 'Separation Geotextile (P-208)', 'Separation Geotextile (P-209)', 'Separation Geotextile (P-219)']
        ss = ss.replace('</sst>', ''.join('<si><t>%s</t></si>' % s for s in newstr) + '</sst>')
        ss = re.sub(r'uniqueCount="\d+"', 'uniqueCount="%d"' % (uc + 4), ss, count=1)
        wr('xl/sharedStrings.xml', ss)
        x = rd('xl/worksheets/sheet5.xml')
        for ref, k in zip(['D20', 'D28', 'D30', 'D32'], range(4)):
            x = re.sub(r'(<c r="%s"[^>]*><v>)\d+(</v></c>)' % ref, r'\g<1>%d\g<2>' % (uc + k), x)
        wr('xl/worksheets/sheet5.xml', x)
        # --- Instructions text box
        d = rd('xl/drawings/drawing3.xml')
        runs = list(re.finditer(r'<a:t>(.*?)</a:t>', d, re.S))
        def set_run(i, fn):
            nonlocal d
            runs2 = list(re.finditer(r'<a:t>(.*?)</a:t>', d, re.S)); r = runs2[i]
            txt = html.unescape(r.group(1)); txt = fn(txt)
            d = d[:r.start(1)] + html.escape(txt, quote=False) + d[r.end(1):]
        set_run(2, lambda t: t + ' If Excel shows a yellow bar reading "BLOCKED CONTENT: The ActiveX content in this file is blocked", go to File > Options > Trust Center > Trust Center Settings > ActiveX Settings, select "Prompt me before enabling all controls with minimal restrictions", click OK, then close and reopen Excel. If that setting is greyed out, ask your IT department to allow ActiveX for this file.')
        set_run(4, lambda t: t.replace('D9 through D38', 'D9 through D39'))
        set_run(5, lambda t: t.replace('“Create Alternatives”', '“Alternative Setup”').replace('"Create Alternatives"', '"Alternative Setup"'))
        set_run(7, lambda t: '4. To include lost airport revenue during runway closures, set cell D38 on the "General Information" worksheet to Yes before creating alternatives. The average daily revenue for the selected airport is pulled automatically from the hidden RevenueData worksheet, which covers the 17 airports meeting Aeronautics\' activity criteria (cell D39 shows the value). If the airport is not in that list, D39 reads "Missing Airport Revenue", lost revenue is carried as $0, and a warning appears on each alternative worksheet; contact Aeronautics to add revenue data. To view the hidden worksheet: Home > Format > Hide & Unhide > Unhide Sheet.')
        set_run(9, lambda t: t.replace('The Summary worksheet provides a summary table and chart for the initial construction cost and NPW for each included alternative.', 'The Summary worksheet shows a results table for each alternative (initial construction, maintenance, rehabilitation, lost revenue and salvage present worth, net present worth, closure days), a line naming the lowest-cost alternative and whether it still wins at the FAA AIP rate of 7%, and five charts. It updates by itself whenever an input changes. Use the "View Summary" button under Alternative Setup to open it and the buttons at the top of the Summary to return.'))
        set_run(8, lambda t: t + ' Closure durations in cells F4 through F10 are pre-filled from production-rate defaults (surface treatment 15,000 SY/day, mill and overlay 3,800 SY/day, PCC joint and slab work) and may be overridden with project-specific values.')
        # the text grew by about a third and the box clips overflow (vertOverflow="clip"): extend it from row 59 to row 76
        assert '<xdr:row>59</xdr:row><xdr:rowOff>114300</xdr:rowOff></xdr:to>' in d
        d = d.replace('<xdr:row>59</xdr:row><xdr:rowOff>114300</xdr:rowOff></xdr:to>', '<xdr:row>76</xdr:row><xdr:rowOff>114300</xdr:rowOff></xdr:to>')
        d = d.replace('<a:ext cx="7334250" cy="8210550"/>', '<a:ext cx="7334250" cy="10963275"/>')
        wr('xl/drawings/drawing3.xml', d)
    # --- formula-driven Summary sheet with embedded charts (no VBA import needed)
    if mbt_mode:
        summ_part = 'xl/worksheets/sheet16.xml'; transplant(work, summ_part, 'xl/drawings/drawing12.xml', 9, 15)
    else:
        summ_part = 'xl/worksheets/sheet14.xml'; transplant(work, summ_part, 'xl/drawings/drawing10.xml', 7, 13)
    # --- navigation button on General Information (hyperlink cell styled like the Summary buttons, no macro)
    gi_part = 'xl/worksheets/sheet4.xml'
    x = rd(gi_part)
    lab = style_of(x, 'B41')                      # "Alternative Setup:" label style
    btn = style_of(rd(summ_part), 'G1')           # dark button style appended by the transplant (Summary row 1)
    x = put_cell(x, 'B44', '<c r="B44"%s t="inlineStr"><is><t>LCCA Summary:</t></is></c>' % (' s="%s"' % lab if lab else ''))
    x = put_cell(x, 'D44', '<c r="D44"%s t="str"><f>HYPERLINK("#Summary!G1","View Summary  \u25ba")</f><v>View Summary  \u25ba</v></c>' % (' s="%s"' % btn if btn else ''))
    x = re.sub(r'<row r="44" ', '<row r="44" ht="21" customHeight="1" ', x, count=1)
    # --- 'Typical Values' reference sheet (appended last, so no sheet indices shift) + button and input hints on General Information
    # the --mbt verification copy predates the RevenueData sheet, so the live revenue rows would not resolve there
    helper_part = None if mbt_mode else add_plain_sheet(work, build_helpers.build, 'Typical Values')
    hint_style = style_of(rd(helper_part), 'A2') if helper_part else None
    if helper_part:
        x = put_cell(x, 'B46', '<c r="B46"%s t="inlineStr"><is><t>Typical input values:</t></is></c>' % (' s="%s"' % lab if lab else ''))
        x = put_cell(x, 'D46', '<c r="D46"%s t="str"><f>HYPERLINK("#\'Typical Values\'!A1","Typical Values  \u25ba")</f><v>Typical Values  \u25ba</v></c>' % (' s="%s"' % btn if btn else ''))
        x = re.sub(r'<row r="46" ', '<row r="46" ht="21" customHeight="1" ', x, count=1)
    # --- 'Method' reference sheet: every calculation with its formula, rule and source (appended after Typical Values)
    method_part = None if mbt_mode else add_plain_sheet(work, build_method.build, 'Method')
    if method_part:
        x = put_cell(x, 'B48', '<c r="B48"%s t="inlineStr"><is><t>How it is calculated:</t></is></c>' % (' s="%s"' % lab if lab else ''))
        x = put_cell(x, 'D48', '<c r="D48"%s t="str"><f>HYPERLINK("#Method!A1","Method  \u25ba")</f><v>Method  \u25ba</v></c>' % (' s="%s"' % btn if btn else ''))
        x = re.sub(r'<row r="48" ', '<row r="48" ht="21" customHeight="1" ', x, count=1)
    # --- the two sheets a user passes through while setting a project up get their own buttons, so every
    # step of the flow on the card above is one click from here
    for row, label, target, text in [(50, 'Unit costs:', 'Pay_Items!A1', 'Pay_Items  \u25ba'),
                                     (52, 'Maintenance policies:', "'Maintenance Policies'!C2", 'Maintenance Policies  \u25ba')]:
        x = ensure_row(x, row)
        x = put_cell(x, 'B%d' % row, '<c r="B%d"%s t="inlineStr"><is><t>%s</t></is></c>'
                     % (row, ' s="%s"' % lab if lab else '', label))
        x = put_cell(x, 'D%d' % row, '<c r="D%d"%s t="str"><f>%s</f><v>%s</v></c>'
                     % (row, ' s="%s"' % btn if btn else '', esc('HYPERLINK("#%s","%s")' % (target, text)), text))
        x = re.sub(r'<row r="%d"([^>]*?)(/?)>' % row,
                   lambda m: '<row r="%d"%s ht="21" customHeight="1"%s>' % (row, m.group(1), m.group(2)), x, count=1)
    for row, hint in (build_helpers.HINTS.items() if helper_part else []):
        x = put_cell(x, 'F%d' % row, '<c r="F%d"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (row, ' s="%s"' % hint_style if hint_style else '', html.escape(hint, quote=False)))
    wr(gi_part, x)
    # --- design pass: look and first-run usability (no calculation changes)
    if mbt_mode:
        tpl = [('xl/worksheets/sheet14.xml', GUIDE_HMA), ('xl/worksheets/sheet15.xml', GUIDE_PCC)]
        tabs = [('xl/worksheets/sheet16.xml', 'FF2A78D6')]
    else:
        tpl = [('xl/worksheets/sheet8.xml', GUIDE_HMA), ('xl/worksheets/sheet10.xml', GUIDE_HMA),
               ('xl/worksheets/sheet1.xml', GUIDE_PCC), ('xl/worksheets/sheet11.xml', GUIDE_PCC),
               ('xl/worksheets/sheet12.xml', GUIDE_HMA)]
        tabs = [('xl/worksheets/sheet14.xml', 'FF2A78D6'),   # Summary
                ('xl/worksheets/sheet15.xml', 'FFBFBFBF'),   # Typical Values
                ('xl/worksheets/sheet2.xml', 'FF595959'),    # Overview
                ('xl/worksheets/sheet3.xml', 'FF595959'),    # Instructions
                ('xl/worksheets/sheet5.xml', 'FFBFBFBF'),    # Pay_Items
                ('xl/worksheets/sheet7.xml', 'FFBFBFBF')]    # Maintenance Policies
        if method_part: tabs.append((method_part, 'FFBFBFBF'))
    # --- the original Alternatives Comparison chart kept on the Summary: give it axis titles too
    cmp_chart = 'xl/charts/chart6.xml' if not mbt_mode else None
    if cmp_chart and os.path.exists(P(cmp_chart)):
        c = axis_titles(rd(cmp_chart), 'Alternative', 'Cost ($)')
        c = re.sub(r'<c:plotVisOnly val="[01]"/>', '<c:plotVisOnly val="0"/>', c)   # its source columns A:E are hidden now
        c = re.sub(r'(<c:valAx>.*?)<c:numFmt[^/]*/>', r'\1<c:numFmt formatCode="$#,##0.0,,&quot;M&quot;" sourceLinked="0"/>', c, count=1, flags=re.S)
        wr(cmp_chart, c)
    design_pass(work, rd, wr, gi_part, summ_part, tpl, tabs)

    # --- workbook: drop broken external links, force full recalculation on open
    w = rd('xl/workbook.xml')
    w = re.sub(r'<externalReferences>.*?</externalReferences>', '', w, flags=re.S)
    w = re.sub(r'<definedName name="Airport_Name" localSheetId="\d+">[^<]*</definedName>', '', w)
    w = re.sub(r'<calcPr ([^/]*)/>', r'<calcPr \1 fullCalcOnLoad="1"/>', w)
    wr('xl/workbook.xml', w)
    r = rd('xl/_rels/workbook.xml.rels')
    r = re.sub(r'<Relationship [^>]*Target="externalLinks/[^"]*"/>', '', r)
    r = re.sub(r'<Relationship [^>]*Target="calcChain.xml"/>', '', r)
    wr('xl/_rels/workbook.xml.rels', r)
    ct = rd('[Content_Types].xml')
    ct = re.sub(r'<Override PartName="/xl/externalLinks/[^"]*"[^>]*/>', '', ct)
    ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', '', ct)
    wr('[Content_Types].xml', ct)
    for p in ['xl/calcChain.xml']:
        if os.path.exists(P(p)): os.remove(P(p))
    if os.path.isdir(P('xl/externalLinks')): shutil.rmtree(P('xl/externalLinks'))
    spelling_fixes(rd, wr)
    package_hygiene(work, P, rd, wr)
    # --- repack
    if os.path.exists(out): os.remove(out)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(P('[Content_Types].xml'), '[Content_Types].xml')
        for root, _, files in os.walk(work):
            for f in files:
                full = os.path.join(root, f); rel = os.path.relpath(full, work).replace(os.sep, '/')
                if rel == '[Content_Types].xml': continue
                z.write(full, rel)
    shutil.rmtree(work)
    print('wrote', out)



GUIDE_HMA = ('STEP 4 of 5: the quantities for this alternative. Grey cells are yours: C4 to C8 project data '
             '(filled from General Information), F4 to F10 closure days, C13 to C22 pay items and E13 to E22 '
             'quantities. Everything else calculates. Unit costs come from Pay_Items.')
GUIDE_PCC = ('STEP 4 of 5: the quantities for this alternative. Grey cells are yours: C4 to C7 project data '
             '(filled from General Information), F4 and F5 closure days, C13 to C22 pay items and E13 to E22 '
             'quantities. Everything else calculates. Unit costs come from Pay_Items.')


# ------------------------------------------------------------------ design pass (look and first-run UX)
def ensure_row(x, r):
    """Insert an empty <row r="N"/> in order if the sheet has no such row."""
    if get_row(x, r): return x
    m = None
    for rm in re.finditer(r'<row r="(\d+)"', x):
        if int(rm.group(1)) > r: m = rm; break
    row = '<row r="%d">' % r + '</row>'
    if m: return x[:m.start()] + row + x[m.start():]
    return x.replace('</sheetData>', row + '</sheetData>', 1)


def insert_before(x, xml, tags):
    """Insert an element before the first of `tags` present (worksheet child order matters to Excel)."""
    for t in tags:
        i = x.find('<' + t)
        if i >= 0: return x[:i] + xml + x[i:]
    return x.replace('</worksheet>', xml + '</worksheet>')


def tab_color(x, rgb):
    """Set the sheet tab colour (tabColor is the first child of sheetPr)."""
    x = re.sub(r'<tabColor[^>]*/>', '', x)
    m = re.search(r'<sheetPr([^>]*?)/>', x)
    if m: return x[:m.start()] + '<sheetPr%s><tabColor rgb="%s"/></sheetPr>' % (m.group(1), rgb) + x[m.end():]
    m = re.search(r'<sheetPr([^>]*)>', x)
    if m: return x[:m.end()] + '<tabColor rgb="%s"/>' % rgb + x[m.end():]
    return x.replace('<dimension', '<sheetPr><tabColor rgb="%s"/></sheetPr><dimension' % rgb, 1)


def design_pass(work, rd, wr, gi_part, summ_part, template_parts, tabs):
    """Visual and first-run usability pass: quick-start card and live input checklist on General
    Information, section bands, navigation on every alternative sheet, and tab colours. Nothing here
    changes a calculation; every cell added is text or a HYPERLINK."""
    NAVY, BLUE, PALE, WHITE, GREY = 'FF1D2733', 'FF2A78D6', 'FFEAF2FB', 'FFFFFFFF', 'FF595959'
    s_title = add_style(work, font=FONT(10, b=True, color=WHITE), fill=FILL(NAVY),
                        alignment='horizontal="left" vertical="center" indent="1"')
    s_body = add_style(work, font=FONT(9), fill=FILL(PALE),
                       alignment='horizontal="left" vertical="top" wrapText="1" indent="1"')
    s_status = add_style(work, font=FONT(10, b=True), fill=FILL(PALE), border=BORDER_BOTTOM,
                         alignment='horizontal="left" vertical="center" wrapText="1" indent="1"')
    s_band = add_style(work, font=FONT(10, b=True), fill=FILL(PALE), border=BORDER_BOTTOM,
                       alignment='horizontal="left" vertical="center"')
    s_note = add_style(work, font=FONT(9, i=True, color=GREY), alignment='horizontal="left" vertical="center"')
    d_ok = add_dxf(work, '<dxf><font><color rgb="FF186A3B"/></font><fill><patternFill><bgColor rgb="FFE8F6EC"/></patternFill></fill></dxf>')
    d_todo = add_dxf(work, '<dxf><font><color rgb="FF7F6000"/></font><fill><patternFill><bgColor rgb="FFFFF3CD"/></patternFill></fill></dxf>')

    # ---- General Information: quick-start card, live checklist, section bands
    x = rd(gi_part)
    btn_dark = style_of(rd(summ_part), 'G1'); btn_blue = style_of(rd(summ_part), 'H1')
    for r in range(2, 8): x = ensure_row(x, r)
    card = ('1.   Overview and Instructions: what the framework does and the rules behind it.\n'
            '2.   General Information (this sheet): fill in the grey cells, D9 to D39.\n'
            '3.   Pay_Items: check the unit costs your alternatives will be priced from.\n'
            '4.   Alternative Setup: add each alternative, then enter its quantities on the sheet it creates.\n'
            '5.   Summary: read the comparison. It updates by itself.\n'
            'Grey cells are yours; white cells calculate. Typical Values has the usual ranges, Method every formula.')
    x = put_cell(x, 'F2', '<c r="F2" s="%s" t="inlineStr"><is><t>HOW TO USE THIS WORKBOOK</t></is></c>' % s_title)
    x = put_cell(x, 'F3', '<c r="F3" s="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (s_body, html.escape(card, quote=False)))
    missing = ('&'.join('IF(D%d="","%s, ","")' % (row, lab) for row, lab in
                        [(9, 'Airport Name'), (25, 'Construction Year'), (26, 'Mainline Area'),
                         (28, 'Markings Area'), (33, 'Analysis Period'), (34, 'Discount Rate')]))
    f = ('IF(LEN(%s)=0,"All required inputs are filled. Next: click Alternative Setup.","Still needed: "&LEFT(%s,LEN(%s)-2))'
         % (missing, missing, missing))
    x = put_cell(x, 'F9', '<c r="F9" s="%s" t="str"><f>%s</f><v>Still needed: Airport Name, Construction Year, Mainline Area, Markings Area</v></c>'
                 % (s_status, html.escape(f, quote=False)))
    for row in (8, 20, 32):   # "Airport Information:", "Project Information:", "LCCA Parameters:"
        for col in 'BCDE':
            cell = '<c r="%s%d" s="%s"%s' % (col, row, s_band, '/>' if col != 'B' else '')
            if col == 'B':
                cur = cell_re('B%d' % row).search(x)
                inner = re.search(r'>(.*)</c>', cur.group(0), re.S).group(1) if cur and '</c>' in cur.group(0) else ''
                t = re.search(r' t="([^"]+)"', cur.group(0)).group(1) if cur and ' t="' in cur.group(0) else None
                cell = '<c r="B%d" s="%s"%s>%s</c>' % (row, s_band, ' t="%s"' % t if t else '', inner)
            x = put_cell(x, '%s%d' % (col, row), cell)
    x = put_cell(x, 'D35', '<c r="D35" s="%s" t="inlineStr"><is><t>Remaining service life at end of analysis period (as percent of cost)</t></is></c>' % s_note)
    merges = '<mergeCells count="3"><mergeCell ref="F2:J2"/><mergeCell ref="F3:J7"/><mergeCell ref="F9:J9"/></mergeCells>'
    cf = ('<conditionalFormatting sqref="F9:J9"><cfRule type="expression" dxfId="%d" priority="1"><formula>LEFT($F$9,3)="All"</formula></cfRule>'
          '<cfRule type="expression" dxfId="%d" priority="2"><formula>LEFT($F$9,3)&lt;&gt;"All"</formula></cfRule></conditionalFormatting>' % (d_ok, d_todo))
    x = re.sub(r'<mergeCells count="\d+">.*?</mergeCells>', '', x, flags=re.S)
    x = insert_before(x, merges + cf, ['dataValidations', 'hyperlinks', 'printOptions', 'pageMargins'])
    # the card now lists five steps, so its rows are taller
    x = re.sub(r'<row r="([3-7])"([^>]*?)(/?)>', lambda m: '<row r="%s"%s ht="21" customHeight="1"%s>' % (m.group(1), m.group(2), m.group(3)), x)
    x = re.sub(r'<row r="9"([^>]*?)(/?)>', lambda m: '<row r="9"%s ht="28" customHeight="1"%s>' % (m.group(1), m.group(2)), x, count=1)
    # print the sheet the way it reads on screen: landscape, and still on ONE page now that the card is
    # taller and two more buttons sit below it
    x = re.sub(r'<pageSetup\b[^>]*/>', '', x)
    x = re.sub(r'(<pageMargins[^>]*/>)', r'\1<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="1"/>', x, count=1)
    if '<pageSetUpPr' not in x:
        x = re.sub(r'(<sheetPr[^>]*>)', r'\1<pageSetUpPr fitToPage="1"/>', x, count=1) if re.search(r'<sheetPr[^>]*>(?!/)', x) else x
    x = tab_color(x, NAVY)
    wr(gi_part, x)
    w = rd('xl/workbook.xml')
    if '_xlnm.Print_Area" localSheetId="3"' not in w:
        pa = '<definedName name="_xlnm.Print_Area" localSheetId="3">\'General Information\'!$A$1:$J$52</definedName>'
        w = w.replace('</definedNames>', pa + '</definedNames>') if '</definedNames>' in w else w.replace('</sheets>', '</sheets><definedNames>' + pa + '</definedNames>', 1)
        wr('xl/workbook.xml', w)

    # ---- alternative worksheets: navigation row and a one-line guide to the grey cells
    for part, note in template_parts:
        t = rd(part)
        t = ensure_row(t, 1)
        t = put_cell(t, 'A1', '<c r="A1" s="%s" t="str"><f>HYPERLINK("#\'General Information\'!D9","◄ General Information")</f><v>◄ General Information</v></c>' % btn_dark)
        t = put_cell(t, 'B1', '<c r="B1" s="%s"/>' % btn_dark)          # merged with A1: column A alone is too narrow
        t = put_cell(t, 'C1', '<c r="C1" s="%s" t="str"><f>HYPERLINK("#Summary!G1","Summary")</f><v>Summary</v></c>' % btn_blue)
        t = put_cell(t, 'D1', '<c r="D1" s="%s" t="inlineStr"><is><t xml:space="preserve">  %s</t></is></c>' % (s_note, html.escape(note, quote=False)))
        t = insert_before(t, '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
                          ['phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins'])
        t = re.sub(r'<row r="1"([^>]*?)(/?)>', lambda m: '<row r="1"%s ht="21" customHeight="1"%s>' % (m.group(1), m.group(2)), t, count=1)
        t = re.sub(r'(<sheetView[^>]*?) topLeftCell="[^"]*"', r'\1', t)
        t = tab_color(t, 'FF8EA9DB')
        wr(part, t)

    # ---- the two reference sheets a user passes through while setting a project up
    reference_sheets(work, rd, wr, btn_dark, btn_blue)

    # ---- tab colours elsewhere
    for part, rgb in tabs:
        x = rd(part); wr(part, tab_color(x, rgb))



def restyle(x, rows, cols, style):
    """Point the given cells at an existing cellXfs entry, leaving their values and formulas alone."""
    for r in rows:
        m = get_row(x, r)
        if not m: continue
        row = m.group(0)
        for col in cols:
            cm = cell_re('%s%d' % (col, r)).search(row)
            if not cm: continue
            cell = cm.group(0)
            cell = re.sub(r'\ss="\d+"', '', cell, count=1)
            cell = cell.replace('<c r="%s%d"' % (col, r), '<c r="%s%d" s="%s"' % (col, r, style), 1)
            row = row[:cm.start()] + cell + row[cm.end():]
        x = x[:m.start()] + row + x[m.end():]
    return x


def row_height(x, r, ht):
    m = get_row(x, r)
    if not m: return x
    row = re.sub(r'\sht="[^"]*"|\scustomHeight="[^"]*"', '', m.group(0), count=2)
    row = row.replace('<row r="%d"' % r, '<row r="%d" ht="%s" customHeight="1"' % (r, ht), 1)
    return x[:m.start()] + row + x[m.end():]


def sheet_view(x, freeze=None, gridlines=None):
    """Freeze panes below `freeze` (a cell ref) and optionally hide the gridlines."""
    m = re.search(r'<sheetView\b[^>]*?(/>|>.*?</sheetView>)', x, re.S)
    if not m: return x
    sv = m.group(0)
    if gridlines is False and 'showGridLines' not in sv:
        sv = sv.replace('<sheetView', '<sheetView showGridLines="0"', 1)
    if freeze:
        col, r = split_ref(freeze)
        pane = ('<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
                % (colnum(col) - 1, r - 1, freeze))
        sv = re.sub(r'<pane\b[^>]*/>', '', sv)
        sv = sv[:-2] + '>' + pane + '</sheetView>' if sv.endswith('/>') else sv.replace('>', '>' + pane, 1)
    return x[:m.start()] + sv + x[m.end():]


def page_setup(x, titles=None):
    """Landscape, fit to one page wide, and repeat a header row on every printed page."""
    x = re.sub(r'<pageSetup\b[^>]*/>', '', x)
    x = re.sub(r'<pageMargins[^>]*/>', lambda m: m.group(0)
               + '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>', x, count=1)
    if '<pageSetUpPr' not in x:
        x = (x.replace('</sheetPr>', '<pageSetUpPr fitToPage="1"/></sheetPr>', 1) if '</sheetPr>' in x
             else x.replace('<dimension', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension', 1))
    return x


def reference_sheets(work, rd, wr, btn_dark, btn_blue):
    """Pay_Items and Maintenance Policies: navigation, a line saying where each sheet sits in the setup
    flow, a marked input column and readable table bands. Nothing here changes a number: Pay_Items keeps
    its Table2 (C2:I59) and every unit cost, and the maintenance schedules live in the hidden templates."""
    NAVY, PALE, WHITE, GREY, AMBER = 'FF1D2733', 'FFEAF2FB', 'FFFFFFFF', 'FF595959', 'FF7F6000'
    s_note = add_style(work, font=FONT(9, i=True, color=GREY),
                       alignment='horizontal="left" vertical="center" wrapText="1"')
    s_title = add_style(work, font=FONT(12, b=True))
    s_band = add_style(work, font=FONT(10, b=True), fill=FILL(PALE),
                       alignment='horizontal="left" vertical="center" wrapText="1"')
    s_hdr = add_style(work, font=FONT(10, b=True, color=WHITE), fill=FILL(NAVY),
                      alignment='horizontal="left" vertical="center" wrapText="1"')
    s_input = add_style(work, font=FONT(10), fill=FILL('FFD9D9D9'), border=BORDER_BOX,
                        alignment='horizontal="right" vertical="center"', numfmt='&quot;$&quot;#,##0.00')
    s_flag = add_style(work, font=FONT(9, color=AMBER), fill=FILL('FFFFF3CD'),
                       alignment='horizontal="left" vertical="center" wrapText="1"')

    def button(x, ref, text, target, style):
        return put_cell(x, ref, '<c r="%s" s="%s" t="str"><f>%s</f><v>%s</v></c>'
                        % (ref, style, esc('HYPERLINK("#%s","%s")' % (target, text)), html.escape(text, quote=False)))

    def text(x, ref, value, style):
        return put_cell(x, ref, '<c r="%s" s="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                        % (ref, style, html.escape(value, quote=False)))

    # ---------------------------------------------------------------- Overview and Instructions (step 1)
    # both sheets are a text box over an empty grid, every column 8.71 wide, so the buttons are merged runs
    for part, line in [('xl/worksheets/sheet2.xml',
                        'STEP 1 of 5: what the framework does and the rules behind it. '
                        'Read this and Instructions, then go to General Information.'),
                       ('xl/worksheets/sheet3.xml',
                        'STEP 1 of 5: how to run an analysis, start to finish. '
                        'When you are ready, go to General Information and fill in the grey cells.')]:
        x = rd(part)
        x = ensure_row(x, 1)
        x = button(x, 'B1', '\u25c4 General Information', "'General Information'!D9", btn_dark)
        for ref in ('C1', 'D1'): x = put_cell(x, ref, '<c r="%s" s="%s"/>' % (ref, btn_dark))
        x = button(x, 'E1', 'Summary', 'Summary!G1', btn_blue)
        x = put_cell(x, 'F1', '<c r="F1" s="%s"/>' % btn_blue)
        x = text(x, 'H1', line, s_note)
        x = row_height(x, 1, 22)
        x = re.sub(r'<mergeCells count="\d+">.*?</mergeCells>', '', x, flags=re.S)
        x = insert_before(x, '<mergeCells count="2"><mergeCell ref="B1:D1"/><mergeCell ref="E1:F1"/></mergeCells>',
                          ['phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks',
                           'printOptions', 'pageMargins', 'drawing'])
        wr(part, x)

    # ---------------------------------------------------------------- Pay_Items (step 3 of the flow)
    x = rd('xl/worksheets/sheet5.xml')
    def priced(r):
        m = get_row(x, r)
        cm = cell_re('F%d' % r).search(m.group(0)) if m else None
        return bool(cm) and ('<v>' in cm.group(0) or '<f>' in cm.group(0))
    items = [r for r in range(4, 60) if get_row(x, r) and cell_re('D%d' % r).search(get_row(x, r).group(0))]
    blank = sum(1 for r in items if not priced(r))
    x = ensure_row(x, 1)
    x = button(x, 'A1', '\u25c4 General Information', "'General Information'!D9", btn_dark)
    x = button(x, 'B1', 'Summary', 'Summary!G1', btn_blue)
    x = text(x, 'D1', '  STEP 3 of 5: the unit costs every alternative is priced from. '
                      'The grey column is the one to edit.', s_note)
    x = row_height(x, 1, 22)
    # the Unit Cost column is the only input on this sheet
    x = restyle(x, range(3, 60), 'F', s_input)
    # the part headings in column A read as bands down the left of the table
    x = restyle(x, [4, 26, 34, 37, 44, 48, 54], 'A', s_band)
    for r, t in [(61, 'Every alternative worksheet prices its quantities from the Unit Cost column above, by looking up the '
                      'pay item description. Change a cost here and every alternative reprices when the workbook recalculates.'),
                 (62, 'The Middle, West and East columns under "Average Pay Item Unit Cost" are empty in this release, so a '
                      'West division project is priced on the same statewide numbers as a Middle one. The Summary states '
                      'this beside the locator map and the Method sheet carries it as an open assumption.'),
                 (63, '%d of the %d pay items carry no unit cost. An alternative that uses one of them prices it at $0, so '
                      'check the Unit Cost column before entering quantities.' % (blank, len(items))),
                 (64, 'Typical ranges and the published sources behind these costs are on the Typical Values sheet.')]:
        x = ensure_row(x, r)
        x = text(x, 'A%d' % r, t, s_flag if r == 63 else s_note)
        for col in 'BCDEFGHIJ':
            x = put_cell(x, '%s%d' % (col, r), '<c r="%s%d" s="%s"/>' % (col, r, s_flag if r == 63 else s_note))
        x = row_height(x, r, 30)
    x = re.sub(r'<mergeCells count="\d+">(.*?)</mergeCells>',
               lambda m: '<mergeCells count="%d">%s%s</mergeCells>'
               % (m.group(1).count('<mergeCell') + 4, m.group(1),
                  ''.join('<mergeCell ref="A%d:J%d"/>' % (r, r) for r in (61, 62, 63, 64))), x, flags=re.S)
    x = sheet_view(x, freeze='A3', gridlines=False)
    x = page_setup(x)
    wr('xl/worksheets/sheet5.xml', x)
    w = rd('xl/workbook.xml')
    if '_xlnm.Print_Titles" localSheetId="4"' not in w:
        pt = '<definedName name="_xlnm.Print_Titles" localSheetId="4">Pay_Items!$2:$2</definedName>'
        w = w.replace('</definedNames>', pt + '</definedNames>') if '</definedNames>' in w else \
            w.replace('</sheets>', '</sheets><definedNames>' + pt + '</definedNames>', 1)
        wr('xl/workbook.xml', w)

    # ---------------------------------------------------------------- Maintenance Policies (reference)
    # the TDOT logo sits over B2:B7, so everything here goes in row 1 and in column C beside it
    x = rd('xl/worksheets/sheet7.xml')
    for r in range(1, 8): x = ensure_row(x, r)
    x = button(x, 'B1', '\u25c4 General Information', "'General Information'!D9", btn_dark)
    x = button(x, 'C1', 'Summary', 'Summary!G1', btn_blue)
    x = row_height(x, 1, 22)
    x = text(x, 'C2', 'MAINTENANCE AND REHABILITATION POLICIES', s_title)
    x = text(x, 'C3', 'Reference: what each alternative type does to the pavement and when. The schedules themselves live '
                      'in the hidden alternative templates, so editing a number here changes nothing.', s_note)
    x = text(x, 'C4', 'Table 1 drives a New HMA alternative, Table 2 a New PCC alternative, Table 3 an HMA overlay and '
                      'Table 4 a PCC rehabilitation. Alternative Setup copies the matching template.', s_note)
    x = text(x, 'C5', 'Rate is the share of the quantity the activity covers: 1 means the whole mainline area, the whole '
                      'markings area or the whole joint length; 0.0075 means three quarters of one percent of it. '
                      'Year Applied counts years after construction.', s_note)
    x = text(x, 'C6', 'Closure days come from the production rates on Typical Values, not from this sheet.', s_note)
    for r, h in ((2, 20), (3, 36), (4, 28), (5, 34), (6, 16)): x = row_height(x, r, h)
    x = restyle(x, [8, 35, 49, 74], 'B', s_title)
    x = restyle(x, [9, 36, 50, 75], 'BCDE', s_hdr)
    for r in (9, 36, 50, 75): x = row_height(x, r, 20)
    x = sheet_view(x, freeze='A8', gridlines=False)
    x = page_setup(x)
    wr('xl/worksheets/sheet7.xml', x)

SPELLING = [
    # the four misspellings that were in the issued v1.1.2 text, left in place through v1.1.x
    ('xl/drawings/drawing2.xml', 'Federal Aviation Adminimstration', 'Federal Aviation Administration'),
    ('xl/drawings/drawing2.xml', 'pavement clossures', 'pavement closures'),
    ('xl/drawings/drawing2.xml', ' associeted with limited facility uses', ' associated with limited facility uses'),
    ('xl/sharedStrings.xml', 'Intial Construction Year', 'Initial Construction Year'),
]


def spelling_fixes(rd, wr):
    """Correct the misspellings carried in the issued text. Wording is otherwise untouched. Project copies
    of v1.1.2 carry a shorter Overview, so each target is corrected where it is present."""
    for part in sorted({p for p, _, _ in SPELLING}):
        x = rd(part)
        for pt, old, new in SPELLING:
            if pt == part: x = x.replace(old, new)
        wr(part, x)


def package_hygiene(work, P, rd, wr):
    """Strip what does not belong in a distributed workbook: cached printer drivers, the empty Power Query
    (DataMashup) stub, the author names and the SharePoint path Excel cached in the package."""
    import glob, datetime
    # printer settings: drop the parts, their relationships and the r:id on each <pageSetup>
    for rels in glob.glob(P('xl/worksheets/_rels/*.rels')):
        r = open(rels, encoding='utf-8').read()
        ids = re.findall(r'<Relationship Id="(rId\d+)"[^>]*Target="\.\./printerSettings/[^"]*"/>', r)
        if not ids: continue
        r = re.sub(r'<Relationship [^>]*Target="\.\./printerSettings/[^"]*"/>', '', r)
        sheet = 'xl/worksheets/' + os.path.basename(rels)[:-5]
        x = rd(sheet)
        for i in ids: x = re.sub(r'(<pageSetup[^>]*?) r:id="%s"' % i, r'\1', x)
        wr(sheet, x)
        if '<Relationship ' in r: open(rels, 'w', encoding='utf-8').write(r)
        else: os.remove(rels)
    if os.path.isdir(P('xl/printerSettings')): shutil.rmtree(P('xl/printerSettings'))
    ct = rd('[Content_Types].xml')
    ct = re.sub(r'<Default Extension="bin" ContentType="[^"]*printerSettings"/>', '', ct)
    # Power Query stub (customXml DataMashup with no queries)
    if os.path.isdir(P('customXml')):
        shutil.rmtree(P('customXml'))
        ct = re.sub(r'<Override PartName="/customXml/[^"]*"[^>]*/>', '', ct)
        r = rd('xl/_rels/workbook.xml.rels'); r = re.sub(r'<Relationship [^>]*Target="\.\./customXml/[^"]*"/>', '', r); wr('xl/_rels/workbook.xml.rels', r)
    wr('[Content_Types].xml', ct)
    # cached SharePoint path of the last save
    w = rd('xl/workbook.xml')
    w = re.sub(r'<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath [^>]*/></mc:Choice></mc:AlternateContent>', '', w)
    wr('xl/workbook.xml', w)
    # document properties: organisation names instead of individuals, modified date = build date
    c = rd('docProps/core.xml')
    c = re.sub(r'<dc:creator>[^<]*</dc:creator>', '<dc:creator>TDOT Aeronautics Division / Applied Research Associates</dc:creator>', c)
    c = re.sub(r'<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '<cp:lastModifiedBy>ARA</cp:lastModifiedBy>', c)
    c = re.sub(r'<dcterms:modified xsi:type="dcterms:W3CDTF">[^<]*</dcterms:modified>',
               '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>' % datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'), c)
    wr('docProps/core.xml', c)
    a = rd('docProps/app.xml')
    a = re.sub(r'<Company>[^<]*</Company>', '<Company>TDOT Aeronautics Division</Company>', a) if '<Company>' in a else a.replace('</Properties>', '<Company>TDOT Aeronautics Division</Company></Properties>')
    wr('docProps/app.xml', a)

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], mbt_mode=('--mbt' in sys.argv))
