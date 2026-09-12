import uno, sys, re, time, subprocess, os, json
from com.sun.star.beans import PropertyValue
S='/tmp/claude-0/-home-user-lcca-dashboard/444ca227-15d7-51d4-a51e-a0ae3dbe81ed/scratchpad'
def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
proc=subprocess.Popen(['soffice','--headless','--invisible','--norestore','--nologo','--accept=socket,host=localhost,port=2002;urp;'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
ctx=None
for i in range(60):
    try:
        local=uno.getComponentContext(); resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
        ctx=resolver.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
smgr=ctx.ServiceManager; desktop=smgr.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
def load(path):
    return desktop.loadComponentFromURL(uno.systemPathToFileUrl(path),'_blank',0,(pv('Hidden',True),pv('MacroExecutionMode',0)))
def val(sh,ref):
    c=sh.getCellRangeByName(ref)
    if c.getError(): return '#ERR%d'%c.getError()
    return c.getString() if c.getType().value=='TEXT' or c.getFormula().startswith('=') and c.getType().value=='FORMULA' and c.getString() and not c.getString().replace('.','',1).replace('-','',1).replace(',','').replace('$','').replace('%','').replace('(','').replace(')','').strip().isdigit() else c.getValue()
def num(sh,ref):
    c=sh.getCellRangeByName(ref); return c.getValue()
def errs(doc):
    out={}
    for sh in doc.Sheets:
        n=0
        rng=sh.queryContentCells(16)  # FORMULA
        for cr in rng.getCells().createEnumeration() if False else []: pass
        cursor=sh.createCursor(); cursor.gotoEndOfUsedArea(False); er=cursor.getRangeAddress().EndRow; ec=cursor.getRangeAddress().EndColumn
        data=sh.getCellRangeByPosition(0,0,ec,er)
        for r in range(er+1):
            for c in range(ec+1):
                cell=data.getCellByPosition(c,r)
                if cell.getType().value=='FORMULA' and cell.getError(): n+=1
        if n: out[sh.Name]=n
    return out
report=[]
def say(*a):
    s=' '.join(str(x) for x in a); print(s); report.append(s)

for label,path in [('TEMPLATE',S+'/TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm'),('MBT',S+'/MBT_check.xlsm')]:
    doc=load(path); doc.calculateAll()
    say('=====',label, os.path.basename(path))
    ctrl=doc.getCurrentController()
    # 1. sheets, visibility, charts, shapes
    for sh in doc.Sheets:
        dp=sh.getDrawPage(); kinds={}
        for i in range(dp.getCount()):
            shp=dp.getByIndex(i); k=shp.getShapeType().split('.')[-1]; kinds[k]=kinds.get(k,0)+1
        say(f'  sheet {sh.Name!r:32} visible={sh.IsVisible} charts={sh.getCharts().getCount()} shapes={kinds}')
    # 2. hyperlink buttons: find and "click" (activate target)
    say('  -- hyperlink buttons')
    names=[sh.Name for sh in doc.Sheets]
    for sh in doc.Sheets:
        cursor=sh.createCursor(); cursor.gotoEndOfUsedArea(False); ra=cursor.getRangeAddress()
        data=sh.getCellRangeByPosition(0,0,ra.EndColumn,ra.EndRow)
        for r in range(ra.EndRow+1):
            for c in range(ra.EndColumn+1):
                cell=data.getCellByPosition(c,r); f=cell.getFormula()
                if 'HYPERLINK(' in f.upper():
                    m=re.search(r'HYPERLINK\("#(.+?)"', f); tgt=m.group(1) if m else '?'
                    tsheet,tcell=tgt.rsplit('!',1); tsheet=tsheet.strip("'")
                    ok=tsheet in names
                    if ok:
                        ctrl.setActiveSheet(doc.Sheets.getByName(tsheet)); ctrl.select(doc.Sheets.getByName(tsheet).getCellRangeByName(tcell))
                        landed=ctrl.getActiveSheet().Name
                    else: landed='MISSING SHEET'
                    say(f'    {sh.Name}!{cell.AbsoluteName.split(".")[-1]:5} "{cell.getString()}" -> {tgt}  target exists={ok} landed={landed} visible={doc.Sheets.getByName(tsheet).IsVisible if ok else None}')
    # 3. errors
    say('  -- error cells after full recalc:', errs(doc) or 'none')
    # 4. Summary state
    sm=doc.Sheets.getByName('Summary'); gi=doc.Sheets.getByName('General Information')
    def snap(tag):
        say(f'  [{tag}] D34={num(gi,"D34")} D33={num(gi,"D33")} D38={gi.getCellRangeByName("D38").getString()} D10={gi.getCellRangeByName("D10").getString()!r}')
        for r in range(4,8):
            if sm.getCellRangeByName(f'G{r}').getString():
                say(f'     row{r}: {sm.getCellRangeByName(f"H{r}").getString()} {sm.getCellRangeByName(f"I{r}").getString()} init={num(sm,f"J{r}"):,.0f} maint={num(sm,f"K{r}"):,.0f} rehab={num(sm,f"L{r}"):,.0f} lost={num(sm,f"M{r}"):,.0f} salv={num(sm,f"N{r}"):,.0f} NPW={num(sm,f"O{r}"):,.0f} delta={num(sm,f"P{r}"):,.0f} days={num(sm,f"Q{r}"):.0f} avail={num(sm,f"R{r}"):.4f}')
        say('     G9 :',sm.getCellRangeByName('G9').getString()); say('     G10:',sm.getCellRangeByName('G10').getString())
        say(f'     sens@3%: {num(sm,"X16"):,.0f} / {num(sm,"Y16"):,.0f}   sens@7%: {num(sm,"X32"):,.0f} / {num(sm,"Y32"):,.0f}   cat sums: {sum(num(sm,f"X{k}") for k in range(5,10)):,.0f} / {sum(num(sm,f"Y{k}") for k in range(5,10)):,.0f}')
        yrs=[(int(num(sm,f'X{r}')),num(sm,f'Y{r}'),num(sm,f'Z{r}'),num(sm,f'AG{r}'),num(sm,f'AH{r}')) for r in range(41,72)]
        nz=[y for y in yrs if y[1] or y[2] or y[3] or y[4]]
        say('     by-year nonzero rows:',len(nz),'first',nz[:2],'last',nz[-1:], '| cumPW end:',f'{num(sm,"AC71"):,.0f} / {num(sm,"AD71"):,.0f}')
        e=errs(doc); say('     errors:',e or 'none')
    snap('as loaded')
    if label=='TEMPLATE':
        gi.getCellRangeByName('D38').setString('Yes'); gi.getCellRangeByName('D10').setString('MBT'); doc.calculateAll()
        t=doc.Sheets.getByName('TMP(NewHMA)_IndirectCost'); say('     [D10=MBT, D38=Yes] template F2=',num(t,'F2'),'G2=',repr(t.getCellRangeByName('G2').getString()),'GI D39=',gi.getCellRangeByName('D39').getString())
        gi.getCellRangeByName('D10').setString('Zzz Test'); doc.calculateAll()
        say('     [D10=unknown, D38=Yes] template F2=',num(t,'F2'),'G2=',repr(t.getCellRangeByName('G2').getString()),'GI D39=',gi.getCellRangeByName('D39').getString(),'errors:',errs(doc) or 'none')
        gi.getCellRangeByName('D38').setString('No'); doc.calculateAll()
        say('     [D38=No] template F2=',num(t,'F2'),'G2=',repr(t.getCellRangeByName('G2').getString()))
    if label=='MBT':
        # scenario A: discount rate 7%
        gi.getCellRangeByName('D34').setValue(7); doc.calculateAll(); snap('D34=7')
        gi.getCellRangeByName('D34').setValue(3)
        # scenario B: 20-year period
        gi.getCellRangeByName('D33').setValue(20); doc.calculateAll(); snap('D33=20')
        gi.getCellRangeByName('D33').setValue(30)
        # scenario C: lost revenue off
        gi.getCellRangeByName('D38').setString('No'); doc.calculateAll(); snap('D38=No')
        gi.getCellRangeByName('D38').setString('Yes')
        # scenario D: airport outside the 17
        old=gi.getCellRangeByName('D10').getString(); gi.getCellRangeByName('D10').setString('Zzz Test Airport'); doc.calculateAll(); snap('D10=unknown airport')
        alt=doc.Sheets.getByName(doc.Sheets.getByName('Database').getCellRangeByName('D4').getString()); say('     alt1 G2 warning:',repr(alt.getCellRangeByName('G2').getString()), 'F2=',num(alt,'F2'))
        gi.getCellRangeByName('D10').setString(old)
        # scenario E: one alternative only, then none
        db=doc.Sheets.getByName('Database'); keep=[[db.getCellRangeByName(f'{c}{r}').getString() for c in 'ABD'] for r in range(4,8)]
        for c in 'ABD': db.getCellRangeByName(f'{c}5').setString('')
        doc.calculateAll(); snap('one alternative')
        for c in 'ABD': db.getCellRangeByName(f'{c}4').setString('')
        doc.calculateAll(); snap('no alternatives')
        for r,row in zip(range(4,8),keep):
            for c,v in zip('ABD',row): db.getCellRangeByName(f'{c}{r}').setString(v)
        doc.calculateAll(); snap('restored')
    doc.close(True)
open(S+'/uno/walk_report.txt','w').write('\n'.join(report))
proc.terminate()
