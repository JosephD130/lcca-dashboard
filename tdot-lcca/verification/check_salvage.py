"""The salvage fractions must now restate themselves correctly at whatever analysis period is set."""
import uno, os, subprocess, sys, time
from com.sun.star.beans import PropertyValue
WB, PORT = sys.argv[1], 2091
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
subprocess.run(['rm', '-rf', '/tmp/lo_slv'])
subprocess.Popen(['soffice', '-env:UserInstallation=file:///tmp/lo_slv', '--headless', '--invisible',
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
mp = doc.Sheets.getByName('Maintenance Policies')
gi = doc.Sheets.getByName('General Information')
fails = []
def check(label, ok, detail=''):
    print(('PASS  ' if ok else 'FAIL  ') + label + (('   |   ' + str(detail)) if detail else ''))
    if not ok: fails.append(label)

overlay = mp.getCellRangeByName('E22').getValue()
check('the HMA overlay year is still a policy input', overlay == 20, overlay)
# (period, expected HMA fraction, expected PCC fraction)
for period, hma, pcc in [(30, 6 / 16.0, 10 / 40.0), (25, 11 / 16.0, 15 / 40.0),
                         (20, 1.0, 20 / 40.0), (36, 0.0, 4 / 40.0), (15, 0.0, 25 / 40.0)]:
    gi.getCellRangeByName('D33').setValue(period); doc.calculateAll()
    gh, gp = mp.getCellRangeByName('D32').getValue(), mp.getCellRangeByName('D46').getValue()
    check('at %d years HMA credits %.1f%% and PCC %.1f%%' % (period, hma * 100, pcc * 100),
          abs(gh - hma) < 1e-9 and abs(gp - pcc) < 1e-9, '%.4f / %.4f' % (gh, gp))
    check('  and the salvage year shown follows the period',
          mp.getCellRangeByName('E32').getValue() == period and mp.getCellRangeByName('E46').getValue() == period)
gi.getCellRangeByName('D33').setValue(30); doc.calculateAll()
print()
print('  HMA row reads:', mp.getCellRangeByName('C32').getString())
print('  PCC row reads:', mp.getCellRangeByName('C46').getString())
check('the HMA sentence states 6 of 16 and 37.5%',
      '6 of 16' in mp.getCellRangeByName('C32').getString() and '37.5%' in mp.getCellRangeByName('C32').getString())
check('the PCC sentence states 10 of 40 and 25.0%',
      '10 of 40' in mp.getCellRangeByName('C46').getString() and '25.0%' in mp.getCellRangeByName('C46').getString())
check('the two rehabilitation tables keep their stated zero salvage',
      mp.getCellRangeByName('D71').getValue() == 0 and mp.getCellRangeByName('D85').getValue() == 0)
doc.close(True)
print('\n%d failed' % len(fails))
print('ALL PASS' if not fails else 'FAILED: ' + '; '.join(fails))
sys.exit(1 if fails else 0)
