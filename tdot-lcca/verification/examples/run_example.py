"""Parametrised end-to-end run of the TDOT LCCA framework on a made-up project.

Emulates what the Alternative Setup form does (copy the hidden templates, register the alternatives on
Database, write Summary A:E), fills General Information and the pay items, recalculates with LibreOffice,
reads every result block back, cross-checks the present-worth arithmetic independently and renders the
Summary and Alt sheets to PNG.

usage: python3 run_example.py <example: gkt|mkl> <workbook.xlsm> <out_dir> [port]
"""
import uno, time, subprocess, json, sys, os, re
from com.sun.star.beans import PropertyValue

EX, WB, OUT = sys.argv[1], sys.argv[2], os.path.abspath(sys.argv[3])
PORT = int(sys.argv[4]) if len(sys.argv) > 4 else 2021
os.makedirs(OUT, exist_ok=True)

# ---------------------------------------------------------------- scenarios
def cy(inches, area): return round(inches * area / 36)                     # C.Y. of a layer
def tons(inches, area, pcf=145): return round(inches * pcf * area / 2666.6667)  # tons of asphalt
MARK = 'Markings (prep, markings, reflective media)'

if EX == 'gkt':
    A = 45883                    # 5,506 x 75 ft runway
    SC = dict(airport='Gatlinburg-Pigeon Forge Airport', consultant='ARA', projno='6101', projname='Runway 10-28 Reconstruction',
              branch_type='Runway', branch='Runway 10-28', ptype='Reconstruction', desc='Full-depth reconstruction, 5,506 x 75 ft, no paved shoulders',
              year=2028, area=A, shoulder=0, markings=12000, mark_type='Reflective', period=30, rate=3, mob=10, eng=5, indirect='Yes')
    ALTS = [
        ('HMA', '4" P-401 surface + 4" P-401 base on 6" P-209 on 6" P-154',
         [('Unclassified Excavation', cy(20, A)), ('Subbase Course', cy(6, A)), ('Crushed Aggregate Base Course', cy(6, A)),
          ('Asphalt Base Course', tons(4, A)), ('Asphalt Surface Course', tons(4, A)), (MARK, 12000)]),
        ('PCC', '9" P-501 on 6" P-209 on 6" P-154',
         [('Unclassified Excavation', cy(21, A)), ('Subbase Course', cy(6, A)), ('Crushed Aggregate Base Course', cy(6, A)),
          ('Concrete Pavement, 9-inch', A), (MARK, 12000)]),
        ('PCC', '11" P-501 on 6" P-209 (no subbase)',
         [('Unclassified Excavation', cy(17, A)), ('Crushed Aggregate Base Course', cy(6, A)),
          ('Concrete Pavement, 11-inch', A), (MARK, 12000)]),
    ]
    EXPECT = {'section': [(8.0, 6.0, 6.0, 20.0), (9.0, 6.0, 6.0, 21.0), (11.0, 6.0, 0.0, 17.0)]}
elif EX == 'mkl':
    A, SH = 100083, 33361         # 6,005 x 150 ft runway with 25 ft paved shoulders each side
    C = A + SH                    # asphalt, base and subbase quantities are taken off the combined area
    SC = dict(airport='McKellar-Sipes Regional Airport', consultant='ARA', projno='6102', projname='Runway 2-20 Reconstruction',
              branch_type='Runway', branch='Runway 2-20', ptype='Reconstruction', desc='Full-depth reconstruction, 6,005 x 150 ft with 25 ft paved shoulders',
              year=2027, area=A, shoulder=SH, markings=25000, mark_type='Reflective', period=30, rate=3, mob=10, eng=5, indirect='Yes')
    ALTS = [
        ('HMA', '3" P-401 surface + 5" P-401 base on 6" P-209 on 6" P-154 (mainline and shoulders)',
         [('Unclassified Excavation', cy(20, C)), ('Subbase Course', cy(6, C)), ('Crushed Aggregate Base Course', cy(6, C)),
          ('Asphalt Base Course', tons(5, C)), ('Asphalt Surface Course', tons(3, C)), (MARK, 25000)]),
        ('PCC', '9" P-501 mainline on 6" P-209 on 6" P-154 (base and subbase under the shoulders too)',
         [('Unclassified Excavation', cy(21, A) + cy(12, SH)), ('Subbase Course', cy(6, C)), ('Crushed Aggregate Base Course', cy(6, C)),
          ('Concrete Pavement, 9-inch', A), (MARK, 25000)]),
        ('HMA', '4" P-401 surface + 6" P-401 base on 10" P-209 on lime-treated subgrade (no excavation item)',
         [('Lime Treated subgrade', C), ('Crushed Aggregate Base Course', cy(10, C)),
          ('Asphalt Base Course', tons(6, C)), ('Asphalt Surface Course', tons(4, C)), (MARK, 25000)]),
        ('PCC', '11" P-501 mainline on 6" P-209 on lime-treated subgrade',
         [('Unclassified Excavation', cy(17, A) + cy(6, SH)), ('Lime Treated subgrade', C), ('Crushed Aggregate Base Course', cy(6, C)),
          ('Concrete Pavement, 11-inch', A), (MARK, 25000)]),
    ]
    f = C / A                     # what a combined-area quantity reads back to on the mainline
    EXPECT = {'section': [(8 * f, 6 * f, 6 * f, 20 * f), (9.0, 6 * f, 6 * f, 25.0), (10 * f, 10 * f, 0.0, 20 * f), (11.0, 6 * f, 0.0, 19.0)],
              'shoulder': [(8, 6, 6), (9, 6, 6), (10, 10, 0), (11, 6, 0)]}
else:
    raise SystemExit('unknown example')
PAY_OVERRIDES = {'P-152-4.1': 12.0}   # project unit cost for an item the catalogue leaves blank ($/C.Y.)

# ---------------------------------------------------------------- LibreOffice
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
profile = f'/tmp/lo_ex_{EX}'
subprocess.run(['rm', '-rf', profile])
proc = subprocess.Popen(['soffice', f'-env:UserInstallation=file://{profile}', '--headless', '--invisible', '--norestore', '--nologo',
                         f'--accept=socket,host=localhost,port={PORT};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
for _ in range(90):
    try:
        local = uno.getComponentContext(); resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = resolver.resolve(f'uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(WB)), '_blank', 0, (pv('Hidden', True), pv('MacroExecutionMode', 0)))
num = lambda sh, ref: sh.getCellRangeByName(ref).getValue()
txt = lambda sh, ref: sh.getCellRangeByName(ref).getString()
def setv(sh, ref, v):
    c = sh.getCellRangeByName(ref)
    c.setString(v) if isinstance(v, str) else c.setValue(v)

# ---------------------------------------------------------------- inputs
gi = doc.Sheets.getByName('General Information')
for ref, key in [('D9', 'airport'), ('D14', 'consultant'), ('D15', 'projno'), ('D16', 'projname'), ('D21', 'branch_type'), ('D22', 'branch'),
                 ('D23', 'ptype'), ('D24', 'desc'), ('D25', 'year'), ('D26', 'area'), ('D27', 'shoulder'), ('D28', 'markings'), ('D29', 'mark_type'),
                 ('D33', 'period'), ('D34', 'rate'), ('D36', 'mob'), ('D37', 'eng'), ('D38', 'indirect')]:
    setv(gi, ref, SC[key])
pay = doc.Sheets.getByName('Pay_Items')
for r in range(3, 80):
    if txt(pay, f'C{r}') in PAY_OVERRIDES: setv(pay, f'F{r}', PAY_OVERRIDES[txt(pay, f'C{r}')])

# what the Alternative Setup form does
db = doc.Sheets.getByName('Database'); sm = doc.Sheets.getByName('Summary')
sheets, n0 = [], doc.Sheets.getCount()
for i, (kind, desc, items) in enumerate(ALTS):
    name = f'Alt {i+1} (New {kind})'
    doc.Sheets.copyByName(f'TMP(New{kind})_IndirectCost', name, n0 + i)
    sh = doc.Sheets.getByName(name); sh.IsVisible = True
    for k, (item, qty) in enumerate(items): setv(sh, f'C{13+k}', item); setv(sh, f'E{13+k}', qty)
    r = 4 + i; last = 52 if kind == 'HMA' else 43
    setv(db, f'A{r}', f'Alternative {i+1}'); setv(db, f'B{r}', f'{kind}-New'); setv(db, f'C{r}', f'New {kind}'); setv(db, f'D{r}', name)
    setv(sm, f'A{r}', f'Alt {i+1}'); setv(sm, f'B{r}', f'Alternative {i+1}')
    sm.getCellRangeByName(f'C{r}').setFormula(f"=$'{name}'.G27"); sm.getCellRangeByName(f'D{r}').setFormula(f"=$'{name}'.E{last}")
    setv(sm, f'E{r}', desc)
    sheets.append((sh, kind, last))
doc.calculateAll()

# ---------------------------------------------------------------- read back
N = len(ALTS)
out = {'example': EX, 'scenario': SC, 'alternatives': [], 'summary': {}, 'checks': []}
out['gi'] = {k: txt(gi, k) for k in ['D10', 'D11', 'D12', 'D13', 'D39', 'F9']}
def check(name, ok, detail=''):
    out['checks'].append({'check': name, 'ok': bool(ok), 'detail': str(detail)}); print(('PASS ' if ok else 'FAIL ') + name, detail if not ok else '')

r0 = SC['rate'] / 100
for i, (sh, kind, last) in enumerate(sheets):
    acts = [(txt(sh, f'B{r}'), num(sh, f'C{r}'), num(sh, f'D{r}'), num(sh, f'E{r}')) for r in range(36, last) if txt(sh, f'B{r}')]
    d = {'sheet': sh.Name, 'kind': kind, 'daily_revenue': num(sh, 'F2'),
         'closure_days': {txt(sh, f'E{r}'): num(sh, f'F{r}') for r in range(4, 11) if txt(sh, f'E{r}')},
         'items': [(txt(sh, f'B{r}'), txt(sh, f'C{r}'), txt(sh, f'D{r}'), num(sh, f'E{r}'), num(sh, f'F{r}'), num(sh, f'G{r}')) for r in range(13, 23) if txt(sh, f'C{r}')],
         'subtotal': num(sh, 'G24'), 'mobilization': num(sh, 'G25'), 'engineering': num(sh, 'G26'), 'initial': num(sh, 'G27'),
         'activities': acts, 'NPW': num(sh, f'E{last}')}
    out['alternatives'].append(d)
    # independent arithmetic: item costs, initial total, each activity's present worth, and the NPW total
    check(f'Alt {i+1} item costs = qty x unit cost', all(abs(q * u - c) < 0.5 for _, _, _, q, u, c in d['items']), d['items'])
    check(f'Alt {i+1} initial = subtotal x (1 + mob + eng)', abs(d['subtotal'] * (1 + SC['mob'] / 100 + SC['eng'] / 100) - d['initial']) < 1, (d['subtotal'], d['initial']))
    # column C is the year offset from construction (the initial row carries the calendar year, offset 0)
    pw_ok = all(abs(cost / (1 + r0) ** (yr if yr < 1000 else 0) - pw) < 1 for _, yr, cost, pw in acts)
    check(f'Alt {i+1} activity PW = cost / (1+r)^offset', pw_ok, [(a, yr, round(c), round(p)) for a, yr, c, p in acts][:12])
    check(f'Alt {i+1} NPW = sum of activity PW (initial included)', abs(sum(p for _, _, _, p in acts) - d['NPW']) < 1, (d['initial'], sum(p for *_, p in acts), d['NPW']))

S = out['summary']
S['results'] = [{c: (txt(sm, f'{c}{r}') if c in 'GHI' else num(sm, f'{c}{r}')) for c in 'GHIJKLMNOPQR'} for r in range(4, 4 + N)]
S['AtoE'] = [[txt(sm, f'{c}{r}') if c in 'ABE' else num(sm, f'{c}{r}') for c in 'ABCDE'] for r in range(4, 4 + N)]
S['G9'], S['G10'] = txt(sm, 'G9'), txt(sm, 'G10')
S['comparison'] = [{c: (txt(sm, f'{c}{r}') if c == 'G' else num(sm, f'{c}{r}')) for c in 'GHIJKLMNO'} for r in range(14, 14 + N)]
S['sensitivity'] = {num(sm, f'W{r}'): [num(sm, f'{c}{r}') for c in ['X', 'Y', 'Z', 'AA'][:N]] for r in range(12, 37)}
S['categories'] = {txt(sm, f'W{r}'): [num(sm, f'{c}{r}') for c in ['X', 'Y', 'Z', 'AA'][:N]] for r in range(5, 10)}
S['byyear_hdr'] = [txt(sm, f'{c}40') for c in ['W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK']]
S['byyear'] = [[num(sm, f'{c}{r}') for c in ['X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK']] for r in range(41, 72)]
S['section'] = [{c: (txt(sm, f'{c}{r}') if c in 'GKNPQ' else num(sm, f'{c}{r}')) for c in 'GHIJKLMNPQ'} for r in range(83, 83 + N)]
S['unit_weight'] = num(sm, 'J81')
S['chart_mainline'] = {txt(sm, f'W{r}'): [num(sm, f'{c}{r}') for c in ['X', 'Y', 'Z', 'AA'][:N]] for r in range(76, 80)}
S['chart_shoulder'] = {txt(sm, f'W{r}'): [num(sm, f'{c}{r}') for c in ['X', 'Y', 'Z', 'AA'][:N]] for r in range(82, 87)}

npws = [a['NPW'] for a in out['alternatives']]
best = min(range(N), key=lambda k: npws[k])
# results table agrees with the alt sheets and with the A:E block the form writes
for i, a in enumerate(out['alternatives']):
    R = S['results'][i]
    check(f'Summary row {i+1} initial / NPW match Alt sheet', abs(R['J'] - a['initial']) < 1 and abs(R['O'] - a['NPW']) < 1 and abs(S['AtoE'][i][3] - a['NPW']) < 1, (R['J'], a['initial'], R['O'], a['NPW']))
    cats = [R['K'], R['L'], R['M'], R['N']]
    check(f'Summary row {i+1} categories sum to NPW', abs(R['J'] + sum(cats) - R['O']) < 1, (R['J'], cats, R['O']))
    closure = sum(row[9 + i] for row in S['byyear'])          # by-year closure-day column for this alternative
    check(f'Summary row {i+1} closure days = sum of the by-year closure column', abs(R['Q'] - closure) < 0.5 and closure > 0, (R['Q'], closure))
    check(f'By-year cumulative PW for Alt {i+1} ends at its NPW', abs(S['byyear'][-1][5 + i] - a['NPW']) < 1, (S['byyear'][-1][5 + i], a['NPW']))
check('Verdict names the lowest-NPW alternative', f'Alternative {best+1}' in S['G9'], S['G9'])
check('vs. lowest column: zero for the winner, positive elsewhere', all((abs(S['results'][k]['P']) < 1) == (k == best) for k in range(N)), [S['results'][k]['P'] for k in range(N)])
crf = r0 * (1 + r0) ** SC['period'] / ((1 + r0) ** SC['period'] - 1)
check('EUAC = PW x CRF', all(abs(c['L'] * crf - c['M']) < 1 and abs(c['H'] + c['J'] - c['L']) < 1 for c in S['comparison']), [(round(c['L']), round(c['M'])) for c in S['comparison']])
s3 = S['sensitivity'].get(3.0) or S['sensitivity'].get(3)
check('Sensitivity curve at 3% reproduces the NPWs', s3 is not None and all(abs(s3[k] - npws[k]) < 1 for k in range(N)), (s3, npws))
rates = sorted(S['sensitivity'])
for k, a in enumerate(out['alternatives']):
    curve = [S['sensitivity'][r][k] for r in rates]
    if a['kind'] == 'HMA':   # net future cost: NPW must fall as the rate rises
        check(f'Sensitivity curve Alt {k+1} (HMA) falls as the rate rises', all(x > y for x, y in zip(curve[:-1], curve[1:])), [round(v) for v in curve[::4]])
    else:                    # a large year-30 salvage credit can offset the M&R stream: only require a smooth, bounded curve
        check(f'Sensitivity curve Alt {k+1} (PCC) is smooth and within 5% of the 3% value', all(abs(x - y) < 0.02 * curve[0] for x, y in zip(curve[:-1], curve[1:])) and max(curve) - min(curve) < 0.05 * npws[k], [round(v) for v in curve[::4]])
for i, exp in enumerate(EXPECT['section']):
    sec = S['section'][i]
    got = (sec['H'], sec['I'], sec['J'], sec['L'])
    check(f'Section read-back Alt {i+1} = surface {exp[0]:.2f} / base {exp[1]:.2f} / subbase {exp[2]:.2f} / total {exp[3]:.2f} in', all(abs(g - e) < 0.02 for g, e in zip(got, exp)), got)
    has_exc = any('Excavation' in it[1] for it in out['alternatives'][i]['items'])
    check(f'Section Alt {i+1} excavation check reads "{ "agrees" if has_exc else "no excavation item" }"', sec['N'] == ('agrees' if has_exc else 'no excavation item'), sec['N'])
    check(f'Section Alt {i+1} description string is populated', sec['P'].strip() != '' and '"' in sec['P'], sec['P'])
if 'shoulder' in EXPECT:
    for i, (asph, base, sub) in enumerate(EXPECT['shoulder']):
        layers = S['chart_shoulder']
        got = (layers['Asphalt'][i] + layers['Concrete'][i], layers['Aggregate base'][i], layers['Subbase (P-154)'][i])
        check(f'Shoulder chart Alt {i+1} = {asph} / {base} / {sub} in on the combined area', all(abs(g - e) < 0.02 for g, e in zip(got, (asph, base, sub))), got)
    for i, (asph, base, sub) in enumerate(EXPECT['shoulder']):
        q = S['section'][i]['Q']
        check(f'Shoulder description Alt {i+1} starts with {asph}" and reads the combined-area section', q.startswith(f'{asph}"') and (f'{base}" P2' in q) and ((f'{sub}" P154' in q) == (sub > 0)), q)
else:
    check('Shoulder description column stays blank without a shoulder area', all(s['Q'] == '' for s in S['section']), [s['Q'] for s in S['section']])
    check('Shoulder chart rows stay at zero without a shoulder area', all(v == 0 for vals in S['chart_shoulder'].values() for v in vals), S['chart_shoulder'])

errs = {}
for sh in doc.Sheets:
    cur = sh.createCursor(); cur.gotoEndOfUsedArea(False); ra = cur.getRangeAddress(); n = 0
    rng = sh.getCellRangeByPosition(0, 0, ra.EndColumn, ra.EndRow)
    for rr in range(ra.EndRow + 1):
        for cc in range(ra.EndColumn + 1):
            cell = rng.getCellByPosition(cc, rr)
            if cell.getType().value == 'FORMULA' and cell.getError(): n += 1
    if n: errs[sh.Name] = n
out['errors'] = errs
check('No formula errors on any visible sheet', not any(k for k in errs if not k.startswith('TMP(')), errs)

json.dump(out, open(f'{OUT}/results.json', 'w'), indent=1, default=str)
doc.storeToURL(uno.systemPathToFileUrl(f'{OUT}/{EX.upper()}_LCCA_run.xlsx'), (pv('FilterName', 'Calc MS Excel 2007 XML'),))
doc.getCurrentController().setActiveSheet(sm)
doc.storeToURL(uno.systemPathToFileUrl(f'{OUT}/{EX.upper()}_run.pdf'), (pv('FilterName', 'calc_pdf_Export'),))
doc.close(True); proc.terminate()

# ---------------------------------------------------------------- render
import pymupdf
pdf = pymupdf.open(f'{OUT}/{EX.upper()}_run.pdf'); made = []
for k, pg in enumerate(pdf):
    t = pg.get_text()
    tag = 'summary' if ('COMPARISON' in t or 'PAVEMENT SECTION' in t or 'Net present worth' in t) else ('alt1' if 'Initial Construction Costs' in t and 'alt1' not in ''.join(made) else None)
    if tag:
        f = f'{OUT}/{EX.upper()}_{tag}_p{k+1}.png'; pg.get_pixmap(dpi=110).save(f); made.append(f)
print('rendered', made)
print(json.dumps({'NPW': npws, 'verdict': S['G9'], 'G10': S['G10'], 'section': [(round(s['H'], 2), round(s['I'], 2), round(s['J'], 2), round(s['L'], 2), s['N'], s['P'], s['Q']) for s in S['section']],
                  'shoulder_chart': S['chart_shoulder'], 'fails': [c for c in out['checks'] if not c['ok']]}, indent=1, default=str))
