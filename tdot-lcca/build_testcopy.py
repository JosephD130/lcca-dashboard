"""The ten-second dropdown test copy.

The pay item pickers live on the hidden alternative templates, so checking one in the delivered
workbook means enabling macros and running Alternative Setup first. This copy is the same build with
TMP(NewHMA) made visible and a line at the top of it saying what to look at, so the test is: open,
click a picker in column C, read the list.

usage: python3 build_testcopy.py <built workbook.xlsm> <out.xlsm>
"""
import os, re, shutil, sys, zipfile

SRC, OUT = sys.argv[1], sys.argv[2]
work = OUT + '.work'
if os.path.isdir(work): shutil.rmtree(work)
with zipfile.ZipFile(SRC) as z: z.extractall(work)
P = lambda p: os.path.join(work, p)
rd = lambda p: open(P(p), encoding='utf-8').read()
wr = lambda p, s: open(P(p), 'w', encoding='utf-8').write(s)

w = rd('xl/workbook.xml')
w, n = re.subn(r'(<sheet name="TMP\(NewHMA\)" sheetId="\d+")\s*state="hidden"', r'\1', w)
assert n == 1
wr('xl/workbook.xml', w)

# the sheet's own guide line says what this copy is for
x = rd('xl/worksheets/sheet10.xml')
note = ('TEST COPY. This alternative worksheet is normally hidden; it is visible here so the pay item '
        'picker can be checked without running Alternative Setup. Click a grey box in column C: the '
        'list should show the pay item number and the whole description. Then try Price from in C11. '
        'Nothing on this copy feeds a real analysis.')
m = re.search(r'<c r="A3"[^>]*>.*?</c>', x, re.S)
assert m
style = re.search(r's="(\d+)"', m.group(0))
x = x[:m.start()] + ('<c r="A3" s="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                     % (style.group(1) if style else '0', note)) + x[m.end():]
x = re.sub(r'<row r="3"([^>]*?)(/?)>',
           lambda mm: '<row r="3"%s ht="40" customHeight="1"%s>' % (mm.group(1), mm.group(2)), x, count=1)
x = re.sub(r'<row r="3"([^>]*?) ht="26" customHeight="1"', r'<row r="3"\1', x, count=1)
wr('xl/worksheets/sheet10.xml', x)

if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.write(P('[Content_Types].xml'), '[Content_Types].xml')
    for root, _, files in os.walk(work):
        for f in files:
            full = os.path.join(root, f)
            rel = os.path.relpath(full, work).replace(os.sep, '/')
            if rel != '[Content_Types].xml': z.write(full, rel)
shutil.rmtree(work)
print('wrote', OUT)
