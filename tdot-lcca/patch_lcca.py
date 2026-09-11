#!/usr/bin/env python3
"""Surgical XML patch for TDOA_LCCA_Framework v1.1.2 -> v1.2.0.
Edits sheet XML in place inside the .xlsm zip so VBA, ActiveX comboboxes, charts,
tables and data validations are preserved (openpyxl would drop them).
"""
import re, sys, os, zipfile, shutil, html

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
    return re.compile(r'<c r="%s"(?:\s[^>]*)?(?:/>|>.*?</c>)' % ref, re.S)

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
    return c

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
        set_run(8, lambda t: t + ' Closure durations in cells F4 through F10 are pre-filled from production-rate defaults (surface treatment 15,000 SY/day, mill and overlay 3,800 SY/day, PCC joint and slab work) and may be overridden with project-specific values.')
        wr('xl/drawings/drawing3.xml', d)
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

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], mbt_mode=('--mbt' in sys.argv))
