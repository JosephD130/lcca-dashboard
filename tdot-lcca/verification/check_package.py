#!/usr/bin/env python3
"""Offline package validation: the classes of defect Excel repairs when it opens a workbook.

There is no Open XML SDK or ECMA schema on this machine, so this does not schema-validate. It
checks the things that actually produce "Excel completed file level validation and repair":
broken relationships and content types, style or shared-string indices past the end of their
table, table columns whose names no longer match their header cells, merge counts and overlaps,
rows or cells out of order, and drawing or control references that do not resolve.

usage: python3 verification/check_package.py <workbook.xlsm>
"""
import sys, re, os, zipfile, itertools, html
import xml.dom.minidom as minidom

WB = sys.argv[1] if len(sys.argv) > 1 else 'TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm'
z = zipfile.ZipFile(WB)
names = z.namelist()
rd = lambda n: z.read(n).decode('utf-8', 'replace')
fails = []


def check(label, ok, detail=''):
    print(('PASS  ' if ok else 'FAIL  ') + label + ('' if ok or detail == '' else '   |   %s' % (detail,)))
    if not ok: fails.append(label)


def colnum(c):
    n = 0
    for ch in c: n = n * 26 + (ord(ch) - 64)
    return n


SHEETS = sorted(n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n))
rels_of = lambda part: '%s/_rels/%s.rels' % (os.path.dirname(part), os.path.basename(part))


def rel_map(part):
    r = rels_of(part)
    if r not in names: return {}
    return {m.group(1): (m.group(2), m.group(3))
            for m in re.finditer(r'Id="([^"]+)"[^>]*?Type="[^"]*?/(\w+)"[^>]*?Target="([^"]+)"', rd(r))}


def owner_of(rels_part):
    """The part a .rels file describes. '_rels/.rels' describes the package root."""
    d = os.path.dirname(rels_part)
    return '' if d == '_rels' else os.path.join(os.path.dirname(d), os.path.basename(rels_part)[:-5])


def resolve(part, target):
    if target.startswith('/'): return target.lstrip('/')
    base = os.path.dirname(part)
    return os.path.normpath(os.path.join(base, target)).replace(os.sep, '/') if base else target


print('=' * 78); print('PACKAGE: ' + os.path.basename(WB)); print('=' * 78)

# ---- 1. the container
bad = z.testzip()
check('every zip entry passes its CRC', bad is None, bad)
check('no entry name is absolute or escapes the package',
      not [n for n in names if n.startswith('/') or '..' in n.split('/')],
      [n for n in names if n.startswith('/') or '..' in n.split('/')])
check('no duplicate entry names', len(names) == len(set(names)),
      [n for n in names if names.count(n) > 1][:3])

# ---- 2. every XML part is well formed
unparseable = []
for n in names:
    if not (n.endswith('.xml') or n.endswith('.rels')): continue
    try: minidom.parseString(z.read(n))
    except Exception as e: unparseable.append((n, str(e)[:70]))
check('every XML part is well formed', not unparseable, unparseable[:3])

# ---- 3. content types
ct = rd('[Content_Types].xml')
defaults = {m.lower() for m in re.findall(r'<Default Extension="([^"]+)"', ct)}
overrides = {m.lstrip('/') for m in re.findall(r'<Override PartName="([^"]+)"', ct)}
untyped = [n for n in names
           if n != '[Content_Types].xml' and n not in overrides
           and n.rsplit('.', 1)[-1].lower() not in defaults]
check('every part has a content type, by default or override', not untyped, untyped[:4])
# A Default of application/xml satisfies the rule above but is not what Excel expects of a
# worksheet, chart or table: those need their own Override or the file is repaired.
CT_OF = {r'xl/worksheets/sheet\d+\.xml$': 'spreadsheetml.worksheet+xml',
         r'xl/charts/chart\d+\.xml$': 'drawingml.chart+xml',
         r'xl/tables/table\d+\.xml$': 'spreadsheetml.table+xml',
         r'xl/drawings/drawing\d+\.xml$': 'drawing+xml',
         r'xl/workbook\.xml$': 'ms-excel.sheet.macroEnabled.main+xml',
         r'xl/styles\.xml$': 'spreadsheetml.styles+xml',
         r'xl/sharedStrings\.xml$': 'spreadsheetml.sharedStrings+xml'}
ct_of_part = {m.group(1).lstrip('/'): m.group(2)
              for m in re.finditer(r'<Override PartName="([^"]+)"[^>]*ContentType="([^"]+)"', ct)}
mistyped = [(n, ct_of_part.get(n)) for n in names
            for pat, want in CT_OF.items()
            if re.match(pat, n) and want not in (ct_of_part.get(n) or '')]
check('worksheets, charts, tables and drawings each carry their own content type',
      not mistyped, mistyped[:4])
check('every content-type override points at a part that exists',
      not [o for o in overrides if o not in names], [o for o in overrides if o not in names][:4])

# ---- 4. relationships
dangling, dup_ids = [], []
for n in names:
    if not n.endswith('.rels'): continue
    owner = owner_of(n)
    ids = re.findall(r'Id="([^"]+)"', rd(n))
    if len(ids) != len(set(ids)): dup_ids.append(n)
    for m in re.finditer(r'<Relationship [^>]*Target="([^"]+)"([^>]*)/>', rd(n)):
        if 'TargetMode="External"' in m.group(2): continue
        t = resolve(owner, m.group(1))
        if t not in names: dangling.append((n, m.group(1)))
check('every internal relationship target exists', not dangling, dangling[:4])
check('relationship ids are unique within each part', not dup_ids, dup_ids[:3])

unresolved = []
for n in names:
    if not n.endswith('.xml') or '/_rels/' in n: continue
    have = set(rel_map(n))
    used = set(re.findall(r'r:(?:id|embed|link|pict)="([^"]+)"', rd(n)))
    miss = used - have
    if miss: unresolved.append((n, sorted(miss)[:3]))
check('every r:id used in a part is declared in that part\'s rels', not unresolved, unresolved[:4])

# ---- 5. the workbook and its sheets
wbx = rd('xl/workbook.xml')
wrels = rel_map('xl/workbook.xml')
sheet_of, missing = {}, []
for m in re.finditer(r'<sheet name="([^"]+)"[^>]*?r:id="([^"]+)"', wbx):
    nm, rid = html.unescape(m.group(1)), m.group(2)
    if rid not in wrels: missing.append(nm); continue
    sheet_of[nm] = resolve('xl/workbook.xml', wrels[rid][1])
check('every <sheet> resolves to a worksheet part', not missing, missing)
snames = list(sheet_of)
check('sheet names are unique, at most 31 chars, and legal',
      len(snames) == len(set(snames)) and all(len(s) <= 31 and not set(s) & set(r'[]:*?/\\') for s in snames),
      [s for s in snames if len(s) > 31 or set(s) & set(r'[]:*?/\\')])
check('every worksheet part is referenced by exactly one <sheet>',
      sorted(sheet_of.values()) == SHEETS, sorted(set(SHEETS) - set(sheet_of.values())))

# ---- 6. styles: every index a cell uses has to exist
sx = rd('xl/styles.xml')
count_of = lambda tag: len(re.findall(r'<%s[ />]' % tag, re.search(
    r'<%ss[^>]*>(.*?)</%ss>' % (tag, tag), sx, re.S).group(1))) if re.search(r'<%ss[^>]*>' % tag, sx) else 0
n_xf = len(re.findall(r'<xf ', re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', sx, re.S).group(1)))
n_font, n_fill, n_border = count_of('font'), count_of('fill'), count_of('border')
n_dxf = len(re.findall(r'<dxf>', sx))
numfmts = set(re.findall(r'<numFmt numFmtId="(\d+)"', sx))
BUILTIN = set(str(i) for i in range(0, 50)) | set(str(i) for i in range(37, 60))
overflow = []
for n in SHEETS:
    x = rd(n)
    for s in set(re.findall(r'<c [^>]*s="(\d+)"', x)) | set(re.findall(r'<row [^>]*s="(\d+)"', x)) \
            | set(re.findall(r'<col [^>]*style="(\d+)"', x)):
        if int(s) >= n_xf: overflow.append((n, s))
check('every style index a cell, row or column uses is inside cellXfs',
      not overflow, ('cellXfs=%d' % n_xf, overflow[:4]))
bad_xf = []
for xf in re.findall(r'<xf [^>]*/>|<xf [^>]*>.*?</xf>', re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', sx, re.S).group(1), re.S):
    for attr, cap in (('fontId', n_font), ('fillId', n_fill), ('borderId', n_border)):
        m = re.search(r'%s="(\d+)"' % attr, xf)
        if m and int(m.group(1)) >= cap: bad_xf.append((attr, m.group(1), cap))
    m = re.search(r'numFmtId="(\d+)"', xf)
    if m and m.group(1) not in BUILTIN and m.group(1) not in numfmts: bad_xf.append(('numFmtId', m.group(1), 'undeclared'))
check('every cellXf points at a font, fill, border and number format that exists',
      not bad_xf, bad_xf[:4])
bad_dxf = [(n, d) for n in SHEETS for d in re.findall(r'dxfId="(\d+)"', rd(n)) if int(d) >= n_dxf]
check('every conditional format points at a dxf that exists', not bad_dxf, ('dxfs=%d' % n_dxf, bad_dxf[:4]))

# ---- 7. shared strings
ss = rd('xl/sharedStrings.xml') if 'xl/sharedStrings.xml' in names else '<sst count="0"/>'
n_si = len(re.findall(r'<si>', ss))
bad_si = []
for n in SHEETS:
    for m in re.finditer(r'<c [^>]*t="s"[^>]*>\s*<v>(\d+)</v>', rd(n)):
        if int(m.group(1)) >= n_si: bad_si.append((n, m.group(1)))
check('every shared-string index is inside the string table', not bad_si, ('si=%d' % n_si, bad_si[:4]))

# ---- 8. sheetData order, merges, columns
order_bad, merge_bad, col_bad, dim_bad = [], [], [], []
for n in SHEETS:
    x = rd(n)
    rows = [int(r) for r in re.findall(r'<row r="(\d+)"', x)]
    if rows != sorted(rows) or len(rows) != len(set(rows)): order_bad.append((n, 'rows'))
    for rm in re.finditer(r'<row r="(\d+)"(?=[ />])[^>]*?(?:/>|>(.*?)</row>)', x, re.S):
        r, body = int(rm.group(1)), rm.group(2) or ''
        refs = re.findall(r'<c r="([A-Z]+)(\d+)"', body)
        if any(int(rr) != r for _, rr in refs): order_bad.append((n, 'row %d holds a foreign cell' % r))
        cols = [colnum(c) for c, _ in refs]
        if cols != sorted(cols) or len(cols) != len(set(cols)): order_bad.append((n, 'row %d' % r))
    m = re.search(r'<mergeCells count="(\d+)">(.*?)</mergeCells>', x, re.S)
    if m:
        refs = re.findall(r'<mergeCell ref="([^"]+)"/>', m.group(2))
        if int(m.group(1)) != len(refs): merge_bad.append((n, 'count %s vs %d' % (m.group(1), len(refs))))
        def box(ref):
            a, b = ref.split(':')
            (c1, r1), (c2, r2) = re.match(r'([A-Z]+)(\d+)', a).groups(), re.match(r'([A-Z]+)(\d+)', b).groups()
            return (colnum(c1), int(r1), colnum(c2), int(r2))
        boxes = [box(r) for r in refs]
        for (i, a), (j, b) in itertools.combinations(list(enumerate(boxes)), 2):
            if not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1]):
                merge_bad.append((n, '%s overlaps %s' % (refs[i], refs[j])))
    seen = []
    for cm in re.finditer(r'<col [^>]*min="(\d+)"[^>]*max="(\d+)"|<col [^>]*max="(\d+)"[^>]*min="(\d+)"', x):
        lo, hi = (int(cm.group(1)), int(cm.group(2))) if cm.group(1) else (int(cm.group(4)), int(cm.group(3)))
        if lo > hi or lo < 1 or hi > 16384: col_bad.append((n, lo, hi))
        if any(not (hi < a or b < lo) for a, b in seen): col_bad.append((n, 'overlaps', lo, hi))
        seen.append((lo, hi))
check('rows and cells are in order inside sheetData, with none repeated or misfiled',
      not order_bad, order_bad[:4])
check('no sheet declares a merge count it does not have, and no two merges overlap',
      not merge_bad, merge_bad[:4])
check('every <col> span is legal and none overlaps another', not col_bad, col_bad[:4])

# ---- 9. tables
tbl_sheet = {}
for n in names:
    m = re.match(r'xl/worksheets/_rels/(sheet\d+)\.xml\.rels$', n)
    if not m: continue
    for t in re.findall(r'Target="\.\./(tables/table\d+\.xml)"', rd(n)):
        tbl_sheet['xl/' + t] = 'xl/worksheets/%s.xml' % m.group(1)
name_by_part = {v: k for k, v in sheet_of.items()}
hdr_bad, shape_bad, id_bad = [], [], []
tids = []
for t, sh in sorted(tbl_sheet.items()):
    tx = rd(t)
    tids.append(re.search(r'<table [^>]*id="(\d+)"', tx).group(1))
    ref = re.search(r' ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', tx)
    c0, r0, c1 = colnum(ref.group(1)), int(ref.group(2)), colnum(ref.group(3))
    cols = [html.unescape(c) for c in re.findall(r'<tableColumn [^>]*name="([^"]*)"', tx)]
    if len(cols) != c1 - c0 + 1: shape_bad.append((t, len(cols), c1 - c0 + 1))
    af = re.search(r'<autoFilter ref="([^"]+)"', tx)
    if af and af.group(1) != '%s%d:%s%d' % (ref.group(1), r0, ref.group(3), int(ref.group(4))):
        shape_bad.append((t, 'autoFilter %s' % af.group(1)))
    sx_ = rd(sh)
    rowm = re.search(r'<row r="%d"(?=[ />])[^>]*?(?:/>|>(.*?)</row>)' % r0, sx_, re.S)
    body = (rowm.group(1) or '') if rowm else ''
    hdr = []
    for i in range(len(cols)):
        ref_i = ''
        n_ = c0 + i
        letters = ''
        k = n_
        while k:
            k, rem = divmod(k - 1, 26)
            letters = chr(65 + rem) + letters
        cm = re.search(r'<c r="%s%d"(?=[ />])[^>]*?(?:/>|>(.*?)</c>)' % (letters, r0), body, re.S)
        if not cm: hdr.append(None); continue
        cb, ch_ = cm.group(1) or '', cm.group(0)
        if 't="s"' in ch_:
            v = re.search(r'<v>(\d+)</v>', cb)
            si = re.findall(r'<si>(.*?)</si>', ss, re.S)[int(v.group(1))] if v else ''
            hdr.append(html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S))))
        elif 'inlineStr' in ch_ or '<is>' in cb:
            hdr.append(html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', cb, re.S))))
        else:
            v = re.search(r'<v>(.*?)</v>', cb, re.S)
            hdr.append(html.unescape(v.group(1)) if v else None)
    if hdr != cols:
        hdr_bad.append((t, name_by_part.get(sh, sh), [p for p in zip(cols, hdr) if p[0] != p[1]]))
check('every table column name matches the header cell it sits over', not hdr_bad, hdr_bad[:3])
check('every table declares as many columns as its ref is wide, and its autoFilter matches',
      not shape_bad, shape_bad[:3])
check('table ids are unique across the workbook', len(tids) == len(set(tids)), tids)
part_ids = []
for n in SHEETS:
    for rid in re.findall(r'<tablePart r:id="([^"]+)"/>', rd(n)):
        part_ids.append((n, rid, rid in rel_map(n)))
check('every tablePart resolves', all(ok for _, _, ok in part_ids),
      [(n, r) for n, r, ok in part_ids if not ok])

# ---- 10. drawings and controls
anchor_bad, ctl_bad = [], []
for n in names:
    if not re.match(r'xl/drawings/drawing\d+\.xml$', n): continue
    for m in re.finditer(r'<xdr:(col|row)>(-?\d+)</xdr:\1>', rd(n)):
        if int(m.group(2)) < 0: anchor_bad.append((n, m.group(0)))
check('no drawing anchor has a negative column or row', not anchor_bad, anchor_bad[:4])
for n in SHEETS:
    x, have = rd(n), rel_map(n)
    for rid in re.findall(r'<control [^>]*r:id="([^"]+)"', x) + re.findall(r'<legacyDrawing r:id="([^"]+)"/>', x):
        if rid not in have: ctl_bad.append((n, rid))
check('every ActiveX control and legacy drawing reference resolves', not ctl_bad, ctl_bad[:4])

# ---- 11. defined names and print areas
dn_bad = []
for m in re.finditer(r'<definedName name="([^"]+)"([^>]*)>([^<]*)</definedName>', wbx):
    nm, attrs, val = m.group(1), m.group(2), html.unescape(m.group(3))
    lsi = re.search(r'localSheetId="(\d+)"', attrs)
    if lsi and int(lsi.group(1)) >= len(sheet_of): dn_bad.append((nm, 'localSheetId out of range'))
    for sm in re.finditer(r"'([^']+)'!|(?<![A-Za-z0-9_'])([A-Za-z_][A-Za-z0-9_.]*)!", val):
        s_ = sm.group(1) or sm.group(2)
        if s_ and s_ not in sheet_of: dn_bad.append((nm, 'unknown sheet %r' % s_))
check('every defined name points at a sheet that exists', not dn_bad, dn_bad[:4])

# ---- 12. formulas that reference a sheet by name
ref_bad = set()
for n in SHEETS:
    for f in re.findall(r'<f[^>]*>(.*?)</f>', rd(n), re.S):
        f = html.unescape(f)
        for sm in re.finditer(r"'([^']+)'!\$?[A-Z]", f):
            nm = sm.group(1)
            # INDIRECT builds its sheet name at run time, so skip anything concatenated
            if '&' in nm or '"' in nm: continue
            if nm not in sheet_of: ref_bad.add((n, nm))
check('every quoted sheet name a formula uses exists', not ref_bad, sorted(ref_bad)[:4])

print(); print('=' * 78)
print('%d failed' % len(fails))
print('ALL PASS' if not fails else 'FAILED: ' + '; '.join(fails))
sys.exit(1 if fails else 0)
