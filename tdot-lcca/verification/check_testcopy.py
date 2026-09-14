"""Open the shipped test copy, recalculate, and confirm the visible TMP(NewHMA) is clean."""
import uno, os, subprocess, sys, time
from com.sun.star.beans import PropertyValue
WB, PORT = sys.argv[1], 2071
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
subprocess.run(['rm', '-rf', '/tmp/lo_tc'])
subprocess.Popen(['soffice', '-env:UserInstallation=file:///tmp/lo_tc', '--headless', '--invisible',
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
sheets = doc.Sheets
visible = [sheets.getByIndex(i).Name for i in range(sheets.getCount()) if sheets.getByIndex(i).IsVisible]
errs = {}
for i in range(sheets.getCount()):
    sh = sheets.getByIndex(i)
    cur = sh.createCursor(); cur.gotoEndOfUsedArea(False)
    n = 0
    for r in range(cur.RangeAddress.EndRow + 1):
        for c in range(cur.RangeAddress.EndColumn + 1):
            cell = sh.getCellByPosition(c, r)
            if cell.getError(): n += 1
    if n: errs[sh.Name] = n
sh = sheets.getByName('TMP(NewHMA)')
print('visible sheets:', visible)
print('TMP(NewHMA) is visible:', 'TMP(NewHMA)' in visible)
print('C11 =', repr(sh.getCellRangeByName('C11').getString()))
print('I11 =', sh.getCellRangeByName('I11').getValue())
print('A3  =', sh.getCellRangeByName('A3').getString()[:70])
print('error cells by sheet:', errs)
doc.close(True)
