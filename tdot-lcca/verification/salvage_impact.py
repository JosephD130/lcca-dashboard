"""What the HMA salvage fraction is worth, at the analysis periods a user can actually set.

Reads the populated Murfreesboro workbook, finds each alternative's salvage row, and recomputes the
net present worth with the fraction the remaining-life rule implies at that period instead of the
0.125 the sheet carries.
"""
import uno, os, subprocess, sys, time
from com.sun.star.beans import PropertyValue
WB, PORT = sys.argv[1], 2081
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
subprocess.run(['rm', '-rf', '/tmp/lo_sv'])
subprocess.Popen(['soffice', '-env:UserInstallation=file:///tmp/lo_sv', '--headless', '--invisible',
                  '--norestore', '--nologo', '--accept=socket,host=localhost,port=%d;urp;' % PORT],
                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = None
for _ in range(90):
    try:
        lo = uno.getComponentContext()
        r = lo.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', lo)
        ctx = r.resolve('uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext' % PORT); break
    except Exception: time.sleep(1)
desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(WB)), '_blank', 0,
                                (pv('Hidden', True), pv('MacroExecutionMode', 0)))
doc.calculateAll()
num = lambda sh, r: sh.getCellRangeByName(r).getValue()
txt = lambda sh, r: sh.getCellRangeByName(r).getString()
gi = doc.Sheets.getByName('General Information')
sm = doc.Sheets.getByName('Summary')
OVERLAY_LIFE = 16.0
T0 = 12

names = [txt(sm, 'G%d' % r) for r in range(T0, T0 + 4)]
names = [n for n in names if n]
rate = num(gi, 'D34') / 100.0

for period in (30, 25, 20):
    gi.getCellRangeByName('D33').setValue(period); doc.calculateAll()
    print('\n=== analysis period %d years, discount %.0f%% ===' % (period, rate * 100))
    rows = []
    for i, n in enumerate(names):
        sh = doc.Sheets.getByName(n)
        # find the salvage row and the rehabilitation row in the NPW table
        sv = rh = None
        for r in range(30, 60):
            lab = txt(sh, 'B%d' % r)
            if lab == 'Salvage': sv = r
            if lab.startswith('Rehabilitation 1'): rh = r
        npw = None
        for r in range(30, 60):
            if txt(sh, 'B%d' % r) == 'Net Present Worth': npw = num(sh, 'E%d' % r)
        is_pcc = 'PCC' in n
        base = OVERLAY_LIFE if not is_pcc else 40.0
        placed = num(sh, 'C%d' % rh) if (rh and not is_pcc) else 0.0
        remaining = max(0.0, min(base, placed + base - period))
        implied = remaining / base
        cur_actual = num(sh, 'D%d' % sv)
        cur_disc = num(sh, 'E%d' % sv)
        frac = 0.25 if is_pcc else 0.125
        new_actual = cur_actual / frac * implied if frac else 0.0
        new_disc = new_actual / (1 + rate) ** period
        rows.append((n, placed, remaining, base, frac, implied, npw, npw - cur_disc + new_disc))
        print('  %-18s asset placed yr %-4.0f  %.0f of %.0f yrs left -> %5.1f%%   sheet uses %5.1f%%'
              % (n, placed, remaining, base, implied * 100, frac * 100))
        print('  %-18s salvage PW %12s -> %12s     NPW %14s -> %14s'
              % ('', '{:,.0f}'.format(cur_disc), '{:,.0f}'.format(new_disc),
                 '{:,.2f}'.format(npw), '{:,.2f}'.format(npw - cur_disc + new_disc)))
    a, b = rows[0], rows[1]
    def verdict(x, y, ia, ib):
        w = x[0] if x[ia] < y[ib] else y[0]
        return '%s lower by $%s' % (w, '{:,.0f}'.format(abs(x[ia] - y[ib])))
    print('  as the workbook stands : %s' % verdict(a, b, 6, 6))
    print('  on the implied fraction: %s' % verdict(a, b, 7, 7))
gi.getCellRangeByName('D33').setValue(30); doc.calculateAll()
doc.close(True)
