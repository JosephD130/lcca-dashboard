"""Drive the pricing source end to end: copy an alternative template the way Alternative Setup does,
price it at Regular, fill a Middle cost for one item only, and prove the sheet follows C11 and falls
back to Unit Cost for every item the region leaves blank."""
import uno, os, subprocess, sys, time
from com.sun.star.beans import PropertyValue

WB, PORT = sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 2043
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
prof = '/tmp/lo_price'
subprocess.run(['rm', '-rf', prof])
subprocess.Popen(['soffice', '-env:UserInstallation=file://%s' % prof, '--headless', '--invisible',
                  '--norestore', '--nologo', '--accept=socket,host=localhost,port=%d;urp;' % PORT],
                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = None
for _ in range(90):
    try:
        local = uno.getComponentContext()
        res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = res.resolve('uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext' % PORT); break
    except Exception: time.sleep(1)
desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(WB)), '_blank', 0,
                                (pv('Hidden', True), pv('MacroExecutionMode', 0)))
num = lambda sh, r: sh.getCellRangeByName(r).getValue()
txt = lambda sh, r: sh.getCellRangeByName(r).getString()
def setv(sh, r, v):
    c = sh.getCellRangeByName(r)
    c.setString(v) if isinstance(v, str) else c.setValue(v)

fails = []
def check(label, got, want):
    ok = (abs(got - want) < 1e-6) if isinstance(want, float) else (got == want)
    print(('PASS  ' if ok else 'FAIL  ') + label + '   |   got %r want %r' % (got, want))
    if not ok: fails.append(label)

pay = doc.Sheets.getByName('Pay_Items')
# two items that both carry a statewide cost: row 6 Crack Repair Type I ($4.00), row 8 Crack Seal ($8.00)
A, B = 6, 8
a_name, b_name = txt(pay, 'D%d' % A), txt(pay, 'D%d' % B)
a_reg, b_reg = num(pay, 'F%d' % A), num(pay, 'F%d' % B)

doc.Sheets.copyByName('TMP(NewHMA)', 'PriceTest', doc.Sheets.getCount())
sh = doc.Sheets.getByName('PriceTest'); sh.IsVisible = True
setv(sh, 'C13', a_name); setv(sh, 'E13', 100.0)
setv(sh, 'C14', b_name); setv(sh, 'E14', 100.0)
doc.calculateAll()

check('the picker ships set to Regular', txt(sh, 'C11'), 'Regular')
check('and the hidden column index reads 1', num(sh, 'I11'), 1.0)
check('at Regular, item A prices from Unit Cost', num(sh, 'F13'), a_reg)
check('at Regular, item B prices from Unit Cost', num(sh, 'F14'), b_reg)

# fill a Middle cost for item A only, and leave item B's Middle cell blank
setv(pay, 'G%d' % A, a_reg * 2)
setv(sh, 'C11', 'Middle'); doc.calculateAll()
check('switching to Middle moves the index to 2', num(sh, 'I11'), 2.0)
check('item A now prices from the Middle column', num(sh, 'F13'), a_reg * 2)
check('item B, blank in Middle, falls back to Unit Cost', num(sh, 'F14'), b_reg)
check('the item cost follows the unit cost', num(sh, 'G13'), a_reg * 200)

setv(sh, 'C11', 'East'); doc.calculateAll()
check('East, with no regional costs filled, prices everything at Unit Cost',
      (num(sh, 'F13'), num(sh, 'F14')), (a_reg, b_reg))
setv(sh, 'C11', ''); doc.calculateAll()
check('an empty picker falls back to Regular rather than erroring', num(sh, 'F13'), a_reg)

doc.close(True)
print('\n%d checks, %d failed' % (9, len(fails)))
print('ALL PASS' if not fails else 'FAILED: ' + '; '.join(fails))
sys.exit(1 if fails else 0)
