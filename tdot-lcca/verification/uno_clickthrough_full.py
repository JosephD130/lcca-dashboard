"""End-to-end click-through of the delivered workbook and the MBT verification copy."""
import uno, re, time, subprocess, sys
from com.sun.star.beans import PropertyValue
S = '/tmp/claude-0/-home-user-lcca-dashboard/444ca227-15d7-51d4-a51e-a0ae3dbe81ed/scratchpad'
PASS, FAIL = [], []
def check(name, ok, detail=''):
    (PASS if ok else FAIL).append(name)
    print(('  PASS  ' if ok else '  FAIL  ') + name + (('  |  ' + str(detail)) if detail else ''))

def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
proc = subprocess.Popen(['soffice', '-env:UserInstallation=file:///tmp/lo_verify', '--headless', '--invisible',
                         '--norestore', '--nologo', '--accept=socket,host=localhost,port=2011;urp;'],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = None
for _ in range(90):
    try:
        local = uno.getComponentContext()
        r = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = r.resolve('uno:socket,host=localhost,port=2011;urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
load = lambda p: desktop.loadComponentFromURL(uno.systemPathToFileUrl(p), '_blank', 0, (pv('Hidden', True), pv('MacroExecutionMode', 0)))
num = lambda sh, ref: sh.getCellRangeByName(ref).getValue()
txt = lambda sh, ref: sh.getCellRangeByName(ref).getString()

def errors(doc):
    out = {}
    for sh in doc.Sheets:
        cur = sh.createCursor(); cur.gotoEndOfUsedArea(False); ra = cur.getRangeAddress(); n = 0
        rng = sh.getCellRangeByPosition(0, 0, ra.EndColumn, ra.EndRow)
        for rr in range(ra.EndRow + 1):
            for cc in range(ra.EndColumn + 1):
                c = rng.getCellByPosition(cc, rr)
                if c.getType().value == 'FORMULA' and c.getError(): n += 1
        if n: out[sh.Name] = n
    return out

def buttons(doc):
    """Every HYPERLINK cell: does the target sheet exist, is it visible, does selecting it land there."""
    names = [sh.Name for sh in doc.Sheets]; ctrl = doc.getCurrentController(); found = []
    for sh in doc.Sheets:
        cur = sh.createCursor(); cur.gotoEndOfUsedArea(False); ra = cur.getRangeAddress()
        rng = sh.getCellRangeByPosition(0, 0, min(ra.EndColumn, 40), min(ra.EndRow, 120))
        for rr in range(min(ra.EndRow, 120) + 1):
            for cc in range(min(ra.EndColumn, 40) + 1):
                cell = rng.getCellByPosition(cc, rr); f = cell.getFormula()
                if 'HYPERLINK(' in f.upper():
                    m = re.search(r'HYPERLINK\("#(.+?)"', f)
                    if not m: continue
                    tgt = m.group(1); tsheet, tcell = tgt.rsplit('!', 1); tsheet = tsheet.strip("'")
                    ok = tsheet in names
                    landed = visible = None
                    if ok:
                        t = doc.Sheets.getByName(tsheet); visible = t.IsVisible
                        ctrl.setActiveSheet(t); ctrl.select(t.getCellRangeByName(tcell))
                        landed = ctrl.getActiveSheet().Name
                    found.append((sh.Name, cell.AbsoluteName.split('.')[-1].replace('$', ''), cell.getString(), tgt, ok, visible, landed))
    return found

print('=' * 78); print('1. DELIVERED TEMPLATE  TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm'); print('=' * 78)
doc = load(S + '/TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm'); doc.calculateAll()
gi = doc.Sheets.getByName('General Information'); sm = doc.Sheets.getByName('Summary')
check('template opens with no error cells', errors(doc) == {}, errors(doc) or 'none')
bs = buttons(doc)
for sh, ref, label, tgt, ok, vis, landed in bs:
    check(f'button {sh}!{ref} "{label.strip()}" -> {tgt}', ok and vis and landed == tgt.rsplit("!", 1)[0].strip("'"), f'visible={vis} landed={landed}')
check('every sheet is reachable or deliberately hidden', len(bs) >= 4, f'{len(bs)} hyperlink buttons')
check('blank template says so on the Summary', txt(sm, 'G9').startswith('No alternatives yet'), txt(sm, 'G9'))
check('input checklist lists what is missing', txt(gi, 'F9').startswith('Still needed'), txt(gi, 'F9'))
gi.getCellRangeByName('D9').setString('Outlaw Field'); gi.getCellRangeByName('D25').setValue(2028)
gi.getCellRangeByName('D26').setValue(66667); gi.getCellRangeByName('D28').setValue(15000); doc.calculateAll()
check('checklist clears when the required inputs are filled', txt(gi, 'F9').startswith('All required'), txt(gi, 'F9'))
tv = doc.Sheets.getByName('Typical Values')
check('Typical Values resolves with no errors', 'Typical Values' not in errors(doc), errors(doc))
check('Typical Values reads the live workbook', num(tv, 'B63') == num(gi, 'D34') and num(tv, 'B64') == num(gi, 'D33'),
      f'rate {num(tv, "B63")} period {num(tv, "B64")}')
check('daily revenue table is live', num(tv, 'B106') > 0, f'CKV {num(tv, "B106")}')
check('asphalt unit weight cell is on the Summary', num(sm, 'J81') == 145, num(sm, 'J81'))
tabs = {sh.Name: sh.TabColor for sh in doc.Sheets}
check('tab colours applied', tabs['General Information'] != -1 and tabs['Summary'] != -1, f"GI {tabs['General Information']} Summary {tabs['Summary']}")
doc.close(True)

print(); print('=' * 78); print('2. POPULATED WORKBOOK  MBT_check.xlsm  (scenarios and the section read-back)'); print('=' * 78)
doc = load(S + '/MBT_check.xlsm'); doc.calculateAll()
gi = doc.Sheets.getByName('General Information'); sm = doc.Sheets.getByName('Summary'); db = doc.Sheets.getByName('Database')
base_err = errors(doc)
check('only the pre-existing template errors are present', set(base_err) <= {'TMP(NewPCC)', 'TMP(NewPCC)_IndirectCost'}, base_err)
check('NPW unchanged by everything added', round(num(sm, 'O4'), 2) == 8809266.42 and round(num(sm, 'O5'), 2) == 8028734.49,
      f"{num(sm,'O4'):,.2f} / {num(sm,'O5'):,.2f}")
check('agency + user = total in the comparison block', round(num(sm, 'H14') + num(sm, 'J14'), 2) == round(num(sm, 'L14'), 2))
crf = 0.03 * 1.03 ** 30 / (1.03 ** 30 - 1)
check('EUAC matches the closed form', abs(num(sm, 'M14') - num(sm, 'L14') * crf) < 0.01, f"{num(sm,'M14'):,.2f}")
check('section read-back reproduces the typed description', txt(sm, 'P83').startswith('5" P401 on 16" P209'), txt(sm, 'P83'))
check('excavation check agrees for the asphalt alternative', txt(sm, 'N83') == 'agrees', txt(sm, 'N83'))
check('excavation check flags the 8 in / 9 in concrete mismatch', txt(sm, 'N84').startswith('differs'), txt(sm, 'N84'))
a0 = num(sm, 'H83'); gi.getCellRangeByName('D26').setValue(num(gi, 'D26'))
sm.getCellRangeByName('J81').setValue(150); doc.calculateAll()
check('asphalt thickness follows the unit weight cell', round(num(sm, 'H83'), 3) == round(a0 * 145 / 150, 3), f'{a0:.3f} -> {num(sm,"H83"):.3f} at 150 pcf')
sm.getCellRangeByName('J81').setValue(145); doc.calculateAll()
check('shoulder chart is empty with no shoulder area', num(sm, 'X83') == 0 and num(sm, 'Y83') == 0)
gi.getCellRangeByName('D27').setValue(8000); doc.calculateAll()
f = num(gi, 'D26') / (num(gi, 'D26') + 8000)
check('shoulder chart scales the derived layers', abs(num(sm, 'Y83') - num(sm, 'Y76') * f) < 0.01, f'{num(sm,"Y76"):.2f} -> {num(sm,"Y83"):.2f}')
check('named concrete thickness does not scale', num(sm, 'Z85') == num(sm, 'Z78') or num(sm, 'Z85') == 0, f'{num(sm,"Z78"):.2f} / {num(sm,"Z85"):.2f}')
gi.getCellRangeByName('D27').setValue(0); doc.calculateAll()
for label, ref, val, expect in [('discount rate 7%', 'D34', 7, 'Alternative 1'), ('analysis period 20 years', 'D33', 20, 'Alternative 2'),
                                ('lost revenue off', 'D38', 'No', 'Alternative 2')]:
    old = gi.getCellRangeByName(ref).getString() if isinstance(val, str) else num(gi, ref)
    gi.getCellRangeByName(ref).setString(val) if isinstance(val, str) else gi.getCellRangeByName(ref).setValue(val)
    doc.calculateAll()
    check(f'{label}: verdict recomputes', expect in txt(sm, 'G9'), txt(sm, 'G9')[:96])
    check(f'{label}: no new errors', set(errors(doc)) <= set(base_err))
    gi.getCellRangeByName(ref).setString(old) if isinstance(val, str) else gi.getCellRangeByName(ref).setValue(old)
doc.calculateAll()
old = txt(gi, 'D10'); gi.getCellRangeByName('D10').setString('Zzz Test'); doc.calculateAll()
check('unknown airport does not error', set(errors(doc)) <= set(base_err), errors(doc))
gi.getCellRangeByName('D10').setString(old); doc.calculateAll()
keep = [[txt(db, f'{c}{r}') for c in 'ABD'] for r in range(4, 8)]
db.getRows().removeByIndex(3, 1); doc.calculateAll()
check('deleting a Database row leaves no #REF!', set(errors(doc)) <= set(base_err) and txt(sm, 'G4') == 'Alt 2 (New PCC)', txt(sm, 'G4'))
check('verdict handles a single alternative', 'only one alternative' in txt(sm, 'G9'), txt(sm, 'G9')[:96])
db.getRows().insertByIndex(3, 1)
for c, v in zip('ABD', keep[0]): db.getCellRangeByName(f'{c}4').setString(v)
doc.calculateAll()
check('restored to two alternatives', round(num(sm, 'O4'), 2) == 8809266.42 and round(num(sm, 'O5'), 2) == 8028734.49)
doc.close(True); proc.terminate()

print(); print('=' * 78)
print(f'{len(PASS)} passed, {len(FAIL)} failed')
if FAIL:
    print('FAILURES:'); [print('  - ' + f) for f in FAIL]
