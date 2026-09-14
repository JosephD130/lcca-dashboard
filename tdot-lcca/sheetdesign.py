#!/usr/bin/env python3
"""Shared helpers for the sheet design passes: styles, cells, cards and the drawing anchors.

Every pass works on an unpacked copy of the .xlsm and edits the sheet XML directly, so nothing
here touches a formula, a defined name or a table range. The palette and the card language are
the workbook's own, as the Summary and General Information already use them.
"""
import os, re, html
import xml.dom.minidom as minidom

# ---------------------------------------------------------------- palette (the workbook's own)
NAVY   = 'FF1D2733'
BLUE   = 'FF2A78D6'
PALE   = 'FFEAF2FB'
WHITE  = 'FFFFFFFF'
PAPER  = 'FFF5F7F9'
LINE   = 'FFDFE4EA'
GREY   = 'FF595959'
MUTED  = 'FF5B6675'
INPUT  = 'FFD9D9D9'
GREEN  = 'FF0F7B4F'
GREEN_SOFT = 'FFE4F2EA'
AMBER  = 'FF7F6000'
AMBER_SOFT = 'FFFFF3CD'
RED    = 'FFC8102E'
RED_SOFT = 'FFF6DFE3'

FONT = (lambda sz=10, b=False, i=False, color=NAVY:
        '<font>%s%s<sz val="%s"/><color rgb="%s"/><name val="Arial"/><family val="2"/></font>'
        % ('<b/>' if b else '', '<i/>' if i else '', sz, color))
FILL = (lambda rgb: '<fill><patternFill patternType="solid"><fgColor rgb="%s"/>'
        '<bgColor indexed="64"/></patternFill></fill>' % rgb)
# A differential format states only what it overrides, and its solid fill is carried by bgColor.
DXF_FILL = lambda rgb: '<fill><patternFill><bgColor rgb="%s"/></patternFill></fill>' % rgb
DXF_FONT = (lambda b=False, i=False, color=None:
            '<font>%s%s%s</font>' % ('<b/>' if b else '', '<i/>' if i else '',
                                     '<color rgb="%s"/>' % color if color else ''))


def BORDER(left=None, right=None, top=None, bottom=None, color=LINE, style='thin'):
    def side(tag, on):
        if not on: return '<%s/>' % tag
        return '<%s style="%s"><color rgb="%s"/></%s>' % (tag, style, color, tag)
    return ('<border>%s%s%s%s<diagonal/></border>'
            % (side('left', left), side('right', right), side('top', top), side('bottom', bottom)))


# ---------------------------------------------------------------- styles.xml
def _append(xml, tag, items):
    m = re.search(r'<%s count="(\d+)"' % tag, xml)
    if not m:
        return xml.replace('</styleSheet>', '<%s count="%d">%s</%s></styleSheet>'
                           % (tag, len(items), ''.join(items), tag)), 0
    cnt = int(m.group(1))
    xml = re.sub(r'<%s count="\d+"' % tag, '<%s count="%d"' % (tag, cnt + len(items)), xml, count=1)
    if '</%s>' % tag in xml:
        end = xml.index('</%s>' % tag)
        return xml[:end] + ''.join(items) + xml[end:], cnt
    return xml.replace('<%s count="%d"/>' % (tag, cnt + len(items)),
                       '<%s count="%d">%s</%s>' % (tag, cnt + len(items), ''.join(items), tag), 1), cnt


class Styles:
    """Appends fonts, fills, borders and cellXfs to the package styles, reusing what it has added."""

    def __init__(self, work):
        self.path = os.path.join(work, 'xl/styles.xml')
        self.s = open(self.path, encoding='utf-8').read()
        self.cache = {}

    def add(self, font=None, fill=None, border=None, alignment=None, numfmt=None):
        key = (font, fill, border, alignment, numfmt)
        if key in self.cache: return self.cache[key]
        s = self.s
        attrs = ['xfId="0"']
        if numfmt is None:
            attrs.insert(0, 'numFmtId="0"')
        else:
            used = [int(i) for i in re.findall(r'<numFmt numFmtId="(\d+)"', s)]
            nid = max(used + [200]) + 1
            nf = '<numFmt numFmtId="%d" formatCode="%s"/>' % (nid, numfmt)
            if '<numFmts' in s: s, _ = _append(s, 'numFmts', [nf])
            else: s = s.replace('<fonts', '<numFmts count="1">%s</numFmts><fonts' % nf, 1)
            attrs.insert(0, 'numFmtId="%d"' % nid); attrs.append('applyNumberFormat="1"')
        for tag, item, name in (('fonts', font, 'fontId'), ('fills', fill, 'fillId'),
                                ('borders', border, 'borderId')):
            if item is None: continue
            s, off = _append(s, tag, [item])
            attrs.append('%s="%d"' % (name, off))
            attrs.append('apply%s="1"' % name[:-2].capitalize())
        body = ''
        if alignment is not None:
            attrs.append('applyAlignment="1"'); body = '<alignment %s/>' % alignment
        xf = '<xf %s>%s</xf>' % (' '.join(attrs), body) if body else '<xf %s/>' % ' '.join(attrs)
        s, idx = _append(s, 'cellXfs', [xf])
        self.s = s
        self.cache[key] = idx
        return idx

    def dxf(self, font=None, fill=None, border=None):
        """Append one differential format (a conditional-formatting look); returns its dxfId."""
        body = ''.join(p for p in (font, fill, border) if p)
        self.s, idx = _append(self.s, 'dxfs', ['<dxf>%s</dxf>' % body])
        return idx

    def save(self):
        minidom.parseString(self.s)
        open(self.path, 'w', encoding='utf-8').write(self.s)


ALIGN = (lambda h='left', v='center', wrap=False, indent=None:
         'horizontal="%s" vertical="%s"%s%s'
         % (h, v, ' wrapText="1"' if wrap else '', ' indent="%d"' % indent if indent else ''))


# ---------------------------------------------------------------- cells
def colnum(col):
    n = 0
    for ch in col: n = n * 26 + (ord(ch) - 64)
    return n


def colname(n):
    out = ''
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)$', ref)
    return m.group(1), int(m.group(2))


def esc(f):
    return html.escape(f, quote=False)


# The lazy [^>]*? matters: a greedy one swallows the "/" of a self-closing tag, the alternation
# then takes the ">...</c>" branch and the match runs on into the next cell or row.
CELL_RE = lambda ref: re.compile(r'<c r="%s"(?=[ />])[^>]*?(?:/>|>.*?</c>)' % ref, re.S)
ROW_RE = lambda r: re.compile(r'<row r="%d"(?=[ />])[^>]*?(?:/>|>.*?</row>)' % r, re.S)


def get_row(x, r):
    return ROW_RE(r).search(x)


def ensure_row(x, r):
    if get_row(x, r): return x
    later = [int(m.group(1)) for m in re.finditer(r'<row r="(\d+)"', x) if int(m.group(1)) > r]
    new = '<row r="%d"/>' % r
    if later:
        m = get_row(x, min(later))
        return x[:m.start()] + new + x[m.start():]
    if '</sheetData>' in x:
        return x.replace('</sheetData>', new + '</sheetData>', 1)
    return x.replace('<sheetData/>', '<sheetData>%s</sheetData>' % new, 1)


def remove_cell(x, ref):
    m = CELL_RE(ref).search(x)
    return x[:m.start()] + x[m.end():] if m else x


def put_cell(x, ref, cellxml):
    """Insert or replace one cell, keeping the row's cells in column order."""
    col, r = split_ref(ref)
    x = ensure_row(x, r)
    m = CELL_RE(ref).search(x)
    if m: return x[:m.start()] + cellxml + x[m.end():]
    rm = get_row(x, r)
    row = rm.group(0)
    if row.endswith('/>'):
        row2 = row[:-2] + '>' + cellxml + '</row>'
    else:
        n = colnum(col)
        after = None
        for cm in re.finditer(r'<c r="([A-Z]+)%d"' % r, row):
            if colnum(cm.group(1)) > n:
                after = cm.start(); break
        row2 = (row[:after] + cellxml + row[after:]) if after is not None else \
               row[:row.rindex('</row>')] + cellxml + '</row>'
    return x[:rm.start()] + row2 + x[rm.end():]


def blank(x, ref, style):
    return put_cell(x, ref, '<c r="%s" s="%d"/>' % (ref, style))


def text(x, ref, value, style):
    if value == '': return blank(x, ref, style)
    return put_cell(x, ref, '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                    % (ref, style, esc(value)))


def formula(x, ref, f, style, t=None, v='0'):
    """A formula cell. `t` is the cached result type; Excel recalculates on open (fullCalcOnLoad)."""
    ta = ' t="%s"' % t if t else ''
    return put_cell(x, ref, '<c r="%s" s="%d"%s><f>%s</f><v>%s</v></c>' % (ref, style, ta, esc(f), esc(v)))


def link(x, ref, label, target, style):
    return formula(x, ref, 'HYPERLINK("#%s","%s")' % (target, label), style, t='str', v=label)


def row_height(x, r, ht):
    x = ensure_row(x, r)
    m = get_row(x, r)
    row = m.group(0)
    head = row[:row.index('>') + 1] if not row.endswith('/>') else row
    new = re.sub(r' ht="[^"]*"', '', head)
    new = re.sub(r' customHeight="[^"]*"', '', new)
    ins = ' ht="%s" customHeight="1"' % ht
    new = new[:-2] + ins + '/>' if new.endswith('/>') else new[:-1] + ins + '>'
    return x[:m.start()] + new + row[len(head):] + x[m.end():]


def restyle(x, rows, cols, style):
    for r in rows:
        for c in cols:
            ref = '%s%d' % (c, r)
            m = CELL_RE(ref).search(x)
            if m:
                cur = m.group(0)
                if ' s="' in cur: new = re.sub(r' s="\d+"', ' s="%d"' % style, cur, count=1)
                else: new = cur.replace('<c r="%s"' % ref, '<c r="%s" s="%d"' % (ref, style), 1)
                x = x[:m.start()] + new + x[m.end():]
            else:
                x = blank(x, ref, style)
    return x


def merges(x, refs):
    """Replace the sheet's mergeCells with the existing set plus `refs`."""
    m = re.search(r'<mergeCells count="\d+">(.*?)</mergeCells>', x, re.S)
    have = re.findall(r'<mergeCell ref="([^"]+)"/>', m.group(1)) if m else []
    for r in refs:
        if r not in have: have.append(r)
    block = ('<mergeCells count="%d">%s</mergeCells>'
             % (len(have), ''.join('<mergeCell ref="%s"/>' % r for r in have)))
    if m: return x[:m.start()] + block + x[m.end():]
    x = re.sub(r'<mergeCells count="\d+"/>', '', x)
    return insert_before(x, block, ['phoneticPr', 'conditionalFormatting', 'dataValidations',
                                    'hyperlinks', 'printOptions', 'pageMargins', 'drawing',
                                    'legacyDrawing'])


def insert_before(x, xml, tags):
    for t in tags:
        m = re.search(r'<%s[ />]' % t, x)
        if m: return x[:m.start()] + xml + x[m.start():]
    return x.replace('</worksheet>', xml + '</worksheet>')


def cols(x, widths, default_style=None):
    """Replace <cols> with explicit widths. `widths` is {column index: width}; anything past the
    highest index keeps the sheet default."""
    st = ' style="%s"' % default_style if default_style else ''
    top = max(widths)
    out = []
    for i in range(1, top + 1):
        w = widths.get(i)
        if w is None:
            out.append('<col min="%d" max="%d" width="8.7109375"%s/>' % (i, i, st))
        elif w == 0:
            out.append('<col min="%d" max="%d" width="0"%s hidden="1" customWidth="1"/>' % (i, i, st))
        else:
            out.append('<col min="%d" max="%d" width="%s"%s customWidth="1"/>' % (i, i, w, st))
    out.append('<col min="%d" max="16384" width="8.7109375"%s/>' % (top + 1, st))
    block = '<cols>%s</cols>' % ''.join(out)
    m = re.search(r'<cols>.*?</cols>', x, re.S)
    return (x[:m.start()] + block + x[m.end():]) if m else \
        insert_before(x, block, ['sheetData'])


def set_widths(x, overrides):
    """Change the width of individual columns, keeping every other attribute (hidden, style, the
    columns outside the overrides) exactly as it was. Rebuilding <cols> wholesale would drop the
    hidden flags and the chart-data columns, so each entry is edited in place instead."""
    m = re.search(r'<cols>(.*?)</cols>', x, re.S)
    if not m: return x
    per = {}
    order = []
    for cm in re.finditer(r'<col\b[^>]*/>', m.group(1)):
        tag = cm.group(0)
        lo = int(re.search(r'min="(\d+)"', tag).group(1))
        hi = int(re.search(r'max="(\d+)"', tag).group(1))
        for i in range(lo, hi + 1):
            if i not in per: order.append(i)
            per[i] = tag
    out = []
    for i in order:
        tag = per[i]
        tag = re.sub(r'min="\d+"', 'min="%d"' % i, tag)
        tag = re.sub(r'max="\d+"', 'max="%d"' % i, tag)
        if i in overrides:
            w = overrides[i]
            tag = re.sub(r'width="[^"]*"', 'width="%s"' % w, tag) if 'width="' in tag \
                else tag[:-2] + ' width="%s"/>' % w
            if 'customWidth' not in tag: tag = tag[:-2] + ' customWidth="1"/>'
        out.append(tag)
    return x[:m.start()] + '<cols>%s</cols>' % ''.join(out) + x[m.end():]


def col_px_of(x, default=8.7109375):
    """{column index: pixel width} read back out of the sheet's own <cols>."""
    def px(w): return int(((256 * w + int(128 / 7)) / 256) * 7)
    m = re.search(r'<cols>(.*?)</cols>', x, re.S)
    out = {}
    if not m: return out
    for cm in re.finditer(r'<col\b[^>]*/>', m.group(1)):
        tag = cm.group(0)
        lo = int(re.search(r'min="(\d+)"', tag).group(1))
        hi = int(re.search(r'max="(\d+)"', tag).group(1))
        w = re.search(r'width="([\d.]+)"', tag)
        hidden = 'hidden="1"' in tag
        for i in range(lo, hi + 1):
            out[i] = 0 if hidden else px(float(w.group(1)) if w else default)
    return out


def sheet_view(x, freeze=None, gridlines=None, headings=None):
    """Set the frozen pane and the gridline/heading switches on the sheet's first view."""
    m = re.search(r'<sheetView(?=[ />])[^>]*>.*?</sheetView>|<sheetView(?=[ />])[^>]*/>', x, re.S)
    if not m: return x
    tag = m.group(0)
    if tag.endswith('/>'):
        head, body = tag[:-2] + '>', ''
    else:
        head, body = tag[:tag.index('>') + 1], tag[tag.index('>') + 1:tag.rindex('</sheetView>')]
    for attr, val in (('showGridLines', gridlines), ('showRowColHeaders', headings)):
        if val is None: continue
        head = re.sub(r' %s="[^"]*"' % attr, '', head)
        head = head[:-1] + ' %s="%d"' % (attr, 1 if val else 0) + '>'
    if freeze is not None:
        body = re.sub(r'<pane[^>]*/>', '', body)
        col, r = split_ref(freeze)
        n = colnum(col)
        parts = []
        if n > 1: parts.append('xSplit="%d"' % (n - 1))
        if r > 1: parts.append('ySplit="%d"' % (r - 1))
        if parts:
            # Excel honours bottomLeft when only rows are split; bottomRight leaves the pane inert
            active = 'bottomRight' if n > 1 else 'bottomLeft'
            body = ('<pane %s topLeftCell="%s" activePane="%s" state="frozen"/>'
                    % (' '.join(parts), freeze, active)) + body
    return x[:m.start()] + head + body + '</sheetView>' + x[m.end():]


def page_setup(x, titles=None, portrait=False, fit_wide=1):
    ps = ('<pageSetup orientation="%s" fitToWidth="%d" fitToHeight="0"/>'
          % ('portrait' if portrait else 'landscape', fit_wide))
    x = re.sub(r'<pageSetup[^>]*/>', ps, x) if '<pageSetup' in x else \
        insert_before(x, ps, ['drawing', 'legacyDrawing'])
    return x


# ---------------------------------------------------------------- cards
def box(x, st, r0, r1, c0, c1, fill=None, font=None, align=None, color=LINE):
    """Outline a rectangle with a thin border, filling every cell inside it."""
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            b = BORDER(left=(c == c0), right=(c == c1), top=(r == r0), bottom=(r == r1), color=color)
            s = st.add(font=font, fill=FILL(fill) if fill else None, border=b, alignment=align)
            ref = '%s%d' % (colname(c), r)
            m = CELL_RE(ref).search(x)
            if m and ('<f>' in m.group(0) or '<v>' in m.group(0) or '<is>' in m.group(0)):
                cur = m.group(0)
                new = re.sub(r' s="\d+"', ' s="%d"' % s, cur, count=1) if ' s="' in cur else \
                    cur.replace('<c r="%s"' % ref, '<c r="%s" s="%d"' % (ref, s), 1)
                x = x[:m.start()] + new + x[m.end():]
            else:
                x = blank(x, ref, s)
    return x


# ---------------------------------------------------------------- drawing anchors
EMU_PX = 9525


def col_px(widths, default=8.7109375):
    """{column index: pixel width} for the widths dict `cols()` was given."""
    def px(w): return int(((256 * w + int(128 / 7)) / 256) * 7)
    return {i: (px(w) if w else 0) for i, w in widths.items()}, px(default)


def anchor_right(px_of, default_px, right_col, w_emu):
    """(col, colOff) that puts a `w_emu`-wide object's right edge at the right edge of `right_col`
    (1-based). Mirrors the placement the logo band already uses."""
    edge = sum(px_of.get(i, default_px) for i in range(1, right_col + 1))
    left = edge - w_emu / EMU_PX
    c, acc = 1, 0
    while acc + px_of.get(c, default_px) <= left:
        acc += px_of.get(c, default_px); c += 1
    return c - 1, int(round((left - acc) * EMU_PX))
