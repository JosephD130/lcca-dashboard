"""End-to-end: does the division picked on one alternative reach that alternative's net present
worth on the Summary, and leave the other alternatives alone?

Runs against a populated workbook (the MKL worked example, four alternatives), so it exercises the
whole chain: C11 -> I11 -> the unit cost lookups -> item cost -> subtotal -> initial construction ->
the NPW table -> the Summary results table.
"""
import uno, os, subprocess, sys, time
from com.sun.star.beans import PropertyValue

WB, PORT = sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 2073
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
subprocess.run(['rm', '-rf', '/tmp/lo_e2e'])
subprocess.Popen(['soffice', '-env:UserInstallation=file:///tmp/lo_e2e', '--headless', '--invisible',
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

fails = []
def check(label, ok, detail=''):
    print(('PASS  ' if ok else 'FAIL  ') + label + (('   |   ' + str(detail)) if detail else ''))
    if not ok: fails.append(label)

sm = doc.Sheets.getByName('Summary')
pay = doc.Sheets.getByName('Pay_Items')
T0 = 12                                        # first alternative row of the results table
alts = [txt(sm, 'G%d' % r) for r in range(T0, T0 + 4)]
alts = [a for a in alts if a]
check('the worked example still carries four alternatives', len(alts) == 4, alts)
npw0 = [num(sm, 'O%d' % (T0 + i)) for i in range(len(alts))]
init0 = [num(sm, 'J%d' % (T0 + i)) for i in range(len(alts))]

target = doc.Sheets.getByName(alts[0])
check('every alternative sheet ships priced at Regular',
      all(txt(doc.Sheets.getByName(a), 'C11') == 'Regular' for a in alts))

# double the Middle price of every item alternative 1 actually uses
used = [txt(target, 'C%d' % r) for r in range(13, 23) if txt(target, 'C%d' % r)]
bumped = 0
for r in range(3, 60):
    if txt(pay, 'D%d' % r) in used and num(pay, 'F%d' % r):
        pay.getCellRangeByName('G%d' % r).setValue(num(pay, 'F%d' % r) * 2)
        bumped += 1
doc.calculateAll()
check('filling the Middle column changes nothing while the sheet says Regular',
      abs(num(sm, 'O%d' % T0) - npw0[0]) < 0.005, '%d costs filled' % bumped)

target.getCellRangeByName('C11').setString('Middle')
doc.calculateAll()
check('switching that alternative to Middle moves its initial construction',
      num(sm, 'J%d' % T0) > init0[0] * 1.5, '%.2f -> %.2f' % (init0[0], num(sm, 'J%d' % T0)))
check('and its net present worth on the Summary',
      num(sm, 'O%d' % T0) > npw0[0], '%.2f -> %.2f' % (npw0[0], num(sm, 'O%d' % T0)))
check('the other three alternatives are untouched',
      all(abs(num(sm, 'O%d' % (T0 + i)) - npw0[i]) < 0.005 for i in range(1, len(alts))),
      [round(num(sm, 'O%d' % (T0 + i)), 2) for i in range(len(alts))])
check('the verdict line recomputes from the new numbers',
      txt(sm, 'G17').startswith('Lowest present worth:'), txt(sm, 'G17')[:70])

target.getCellRangeByName('C11').setString('Regular')
doc.calculateAll()
check('switching back restores every number exactly',
      all(abs(num(sm, 'O%d' % (T0 + i)) - npw0[i]) < 0.005 for i in range(len(alts))),
      [round(num(sm, 'O%d' % (T0 + i)), 2) for i in range(len(alts))])

doc.close(True)
print('\n%d checks, %d failed' % (8, len(fails)))
print('ALL PASS' if not fails else 'FAILED: ' + '; '.join(fails))
sys.exit(1 if fails else 0)
