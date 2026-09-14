"""Every cell whose value or formula differs between two builds of the workbook."""
import openpyxl, re, sys, warnings
warnings.filterwarnings('ignore')
A, B = sys.argv[1], sys.argv[2]
wa = openpyxl.load_workbook(A); wb = openpyxl.load_workbook(B)
def cells(ws):
    return {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None}
key = lambda r: (int(re.sub(r'[A-Z]+', '', r)), re.sub(r'\d+', '', r))
total = 0
for name in wb.sheetnames:
    a = cells(wa[name]) if name in wa.sheetnames else {}
    b = cells(wb[name])
    ch = [(k, a.get(k), b.get(k)) for k in sorted(set(a) | set(b), key=key) if a.get(k) != b.get(k)]
    if not ch: continue
    total += len(ch)
    print('== %s  (%d cells)' % (name, len(ch)))
    shapes = {}
    for k, x, y in ch:
        s = (re.sub(r'\d+', 'N', str(x))[:110], re.sub(r'\d+', 'N', str(y))[:110])
        shapes.setdefault(s, []).append(k)
    for (x, y), ks in shapes.items():
        print('   %-4s x%-4d %s' % (ks[0], len(ks), ('%s  ->  %s' % (x, y))[:190]))
print('---- %d cells differ in total' % total)
