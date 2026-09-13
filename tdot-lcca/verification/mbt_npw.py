"""MBT regression: the two net present worths the completed Murfreesboro run produced, recomputed from the
patched workbook, plus a count of the formulas that recalculate without error."""
import uno, subprocess, time, os
from com.sun.star.beans import PropertyValue
S = os.path.dirname(os.path.abspath(__file__))
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
subprocess.Popen(['soffice', '-env:UserInstallation=file:///tmp/lo_mbt', '--headless', '--invisible', '--norestore',
                  '--nologo', '--accept=socket,host=localhost,port=2031;urp;'],
                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
for _ in range(120):
    try:
        local = uno.getComponentContext()
        r = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = r.resolve('uno:socket,host=localhost,port=2031;urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
doc = desk.loadComponentFromURL('file://' + S + '/MBT_check.xlsm', '_blank', 0, (pv('Hidden', True),))
doc.calculateAll()
sh = doc.Sheets
n = err = 0
for i in range(sh.Count):
    s = sh.getByIndex(i)
    cur = s.createCursor(); cur.gotoEndOfUsedArea(False)
    for rr in range(cur.RangeAddress.EndRow + 1):
        for cc in range(cur.RangeAddress.EndColumn + 1):
            c = s.getCellByPosition(cc, rr)
            if c.getFormula().startswith('='):
                n += 1
                if c.getError(): err += 1
for name in ['Alt 1 (New HMA)', 'Alt 2 (New PCC)']:
    if sh.hasByName(name):
        s = sh.getByName(name)
        for rr in range(28, 60):
            if s.getCellByPosition(1, rr).getString().strip() == 'Net Present Worth':
                print('%-18s %s' % (name, format(s.getCellByPosition(4, rr).getValue(), ',.2f')))
print('%d formulas recalculated, %d errors' % (n, err))
doc.close(False)
