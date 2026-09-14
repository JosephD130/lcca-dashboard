"""The pay item picker and the division the alternative is priced from.

Two things happen here.

A Forms 2.0 combo box sizes its drop-down list to the control unless ListWidth says otherwise, and
with two columns and no explicit ColumnWidths each column gets half of that.  The pay item pickers
list Pay_Items!C3:D59, so the number took half a cell and the description had the other half, which
is why most descriptions were cut off mid-word.  combo_list_width sets ListWidth on every combo box
in the package so the list opens wide enough to read both columns.  It is a MorphDataControl stream
([MS-OFORMS] 2.2.5): the property mask gains fListWidth, the DataBlock gains the four bytes that
property needs, and cbMorphData grows to match.

The second change is the pricing source.  Pay_Items has had Middle, West and East columns beside
Unit Cost since v1.1.2 and nothing read them.  Each alternative worksheet now carries a Price from
control in C11, and every unit cost lookup on the sheet reads the chosen column, falling back to
Unit Cost wherever the regional cell is blank.
"""
import glob, os, re, struct

COMBO_CLSID = bytes.fromhex('301dd28b42ecce119e0d00aa006002f3')
F_LISTWIDTH = 1 << 10                     # MorphDataPropMask bit 10, [MS-OFORMS] 2.2.5.2
KNOWN_MASKS = {0x80053941, 0x80453941}    # the two shapes this workbook's combo boxes persist as
LIST_WIDTH_IN = 9.0                       # the longest pay item description is 76 characters
HIMETRIC = 2540                           # per inch, the unit ListWidth and Size use

PRICE_SOURCES = ['Regular', 'Middle', 'West', 'East']
SRC_CELL = 'C11'
IDX_CELL = '$I$11'                        # hidden helper: which column of UnitCostGrid to read

# every unit cost on an alternative worksheet looks like this, 156 of them across five templates
OLD_UNIT_COST = re.compile(
    r'IF\((\$?[A-Z]{1,2}\$?\d+)="","N/A",'
    r'VLOOKUP\(\1,CHOOSE\(\{1,2\},Table2\[Pay Item Description\],Table2\[Unit Cost\]\),2,0\)\)')


def new_unit_cost(m):
    """Read the chosen division's column, falling back to Unit Cost where that cell is blank."""
    key = m.group(1)
    at = 'INDEX(UnitCostGrid,MATCH(%s,PayItemKeys,0),%%s)' % key
    return 'IF(%s="","N/A",IF(%s="",%s,%s))' % (key, at % IDX_CELL, at % '1', at % IDX_CELL)


# ------------------------------------------------------------------ the combo boxes
def combo_list_width(work, inches=LIST_WIDTH_IN):
    value = int(round(inches * HIMETRIC))
    touched, skipped = [], []
    for p in sorted(glob.glob(os.path.join(work, 'xl', 'activeX', '*.bin'))):
        d = open(p, 'rb').read()
        if d[:16] != COMBO_CLSID:
            continue
        cb = struct.unpack('<H', d[18:20])[0]
        mask = struct.unpack('<Q', d[20:28])[0]
        if mask & F_LISTWIDTH or mask not in KNOWN_MASKS:
            skipped.append(os.path.basename(p)); continue
        # DataBlock at 28: VariousPropertyBits(4) DisplayStyle(1) pad(1) BoundColumn(2) TextColumn(2)
        # ColumnCount(2) MatchEntry(1) ShowDropButtonWhen(1) pad(2).  ListWidth is four bytes and
        # four-aligned, so it follows DisplayStyle across three bytes of padding and the trailing
        # pad the two 1-byte fields needed goes away.
        out = (d[:33] + b'\x00\x00\x00' + struct.pack('<i', value) + d[34:42] + d[44:])
        assert len(out) == len(d) + 4, p
        out = out[:18] + struct.pack('<H', cb + 4) + struct.pack('<Q', mask | F_LISTWIDTH) + out[28:]
        open(p, 'wb').write(out)
        touched.append(os.path.basename(p))
    return touched, skipped


# ------------------------------------------------------------------ worksheet plumbing
def set_col(x, first, last, **attrs):
    """Set or add a <col> definition covering exactly first..last."""
    cols = re.search(r'<cols>(.*?)</cols>', x, re.S)
    body = cols.group(1) if cols else ''
    keep = []
    for m in re.finditer(r'<col [^>]*/>', body):
        lo = int(re.search(r'min="(\d+)"', m.group(0)).group(1))
        hi = int(re.search(r'max="(\d+)"', m.group(0)).group(1))
        if lo >= first and hi <= last:
            continue                                     # wholly replaced
        if lo <= first <= hi or lo <= last <= hi:        # straddles the new range: trim it
            if lo < first:
                keep.append(re.sub(r'max="\d+"', 'max="%d"' % (first - 1), m.group(0)))
            if hi > last:
                keep.append(re.sub(r'min="\d+"', 'min="%d"' % (last + 1), m.group(0)))
            continue
        keep.append(m.group(0))
    attr = ' '.join('%s="%s"' % (k, v) for k, v in attrs.items())
    keep.append('<col min="%d" max="%d" %s/>' % (first, last, attr))
    keep.sort(key=lambda c: int(re.search(r'min="(\d+)"', c).group(1)))
    new = '<cols>' + ''.join(keep) + '</cols>'
    if cols:
        return x[:cols.start()] + new + x[cols.end():]
    return re.sub(r'(</sheetFormatPr>|<sheetFormatPr[^>]*/>)', r'\1' + new, x, count=1)


def widen_combo(x):
    """The pay item pickers sit over column C; let each one end exactly at C's right-hand edge so a
    wider column makes a wider box rather than leaving a gap or overhanging column D."""
    def fix(m):
        if '<xdr:col>2</xdr:col>' not in m.group(1):
            return m.group(0)
        to = re.sub(r'<xdr:col>\d+</xdr:col><xdr:colOff>\d+</xdr:colOff>',
                    '<xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff>', m.group(2), count=1)
        return '<from>%s</from><to>%s</to>' % (m.group(1), to)
    return re.sub(r'<from>(.*?)</from><to>(.*?)</to>', fix, x, flags=re.S)
