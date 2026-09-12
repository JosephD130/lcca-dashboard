import uno, re, time, subprocess, collections
from com.sun.star.beans import PropertyValue
S='/tmp/claude-0/-home-user-lcca-dashboard/444ca227-15d7-51d4-a51e-a0ae3dbe81ed/scratchpad'
def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
proc=subprocess.Popen(['soffice','-env:UserInstallation=file:///tmp/lo_uno_profile','--headless','--invisible','--norestore','--nologo','--accept=socket,host=localhost,port=2003;urp;'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
for i in range(60):
    try:
        local=uno.getComponentContext(); resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
        ctx=resolver.resolve('uno:socket,host=localhost,port=2003;urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
desktop=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
load=lambda p: desktop.loadComponentFromURL(uno.systemPathToFileUrl(p),'_blank',0,(pv('Hidden',True),pv('MacroExecutionMode',0)))
num=lambda sh,ref: sh.getCellRangeByName(ref).getValue()
txt=lambda sh,ref: sh.getCellRangeByName(ref).getString()
def errs(doc):
    out={}
    for sh in doc.Sheets:
        cur=sh.createCursor(); cur.gotoEndOfUsedArea(False); ra=cur.getRangeAddress(); n=0
        rng=sh.getCellRangeByPosition(0,0,ra.EndColumn,ra.EndRow)
        for r in range(ra.EndRow+1):
            for c in range(ra.EndColumn+1):
                cell=rng.getCellByPosition(c,r)
                if cell.getType().value=='FORMULA' and cell.getError(): n+=1
        if n: out[sh.Name]=n
    return out
doc=load(S+'/TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm'); doc.calculateAll()
print('=== A. text boxes on Overview / Instructions')
for name in ['Overview','Instructions']:
    dp=doc.Sheets.getByName(name).getDrawPage()
    for i in range(dp.getCount()):
        shp=dp.getByIndex(i)
        if 'CustomShape' in shp.getShapeType():
            t=shp.getString(); print(f'  [{name}] {len(t)} chars')
            for kw in ['Create Alternatives','D38','D39','four chart','10 year','10-year','Output','Summary','Chart','View Summary','ActiveX','RevenueData','17 airport','closure','Neel','ARA','2022','2026']:
                for m in re.finditer(re.escape(kw),t): print(f'     "{kw}": ...{t[max(0,m.start()-70):m.end()+90].replace(chr(10)," ")}...'); break
print('=== B. General Information dropdown and inputs')
gi=doc.Sheets.getByName('General Information')
lst=[txt(gi,f'L{r}') for r in range(10,89)]
print('  L10:L88 airports:',len(lst),'blank:',sum(1 for x in lst if not x.strip()),'first:',lst[0],'last:',lst[-1], 'dups:',[k for k,v in collections.Counter(lst).items() if v>1][:5])
print('  D9 formula/value:',repr(gi.getCellRangeByName('D9').getFormula()),'| D10:',repr(gi.getCellRangeByName('D10').getFormula())[:80])
for r in range(9,40):
    c=gi.getCellRangeByName(f'D{r}')
    v=c.Validation
    if v.Type.value!='ANY': print(f'  D{r} validation {v.Type.value} {v.getFormula1()[:60]!r} err={v.ShowErrorMessage}')
print('=== C. Pay_Items duplicate descriptions')
pi=doc.Sheets.getByName('Pay_Items'); cur=pi.createCursor(); cur.gotoEndOfUsedArea(False); er=cur.getRangeAddress().EndRow
hdr=[txt(pi,f'{c}1') for c in 'ABCDEFGH']; print('  headers',hdr)
descs=[txt(pi,f'D{r}') for r in range(2,er+1) if txt(pi,f'D{r}')]
d=collections.Counter(descs); print('  rows',len(descs),'duplicate descriptions:',[k for k,v in d.items() if v>1])
print('=== E. Summary page setup')
sm=doc.Sheets.getByName('Summary'); pa=sm.getPrintAreas(); print('  print areas:',[(a.StartColumn,a.StartRow,a.EndColumn,a.EndRow) for a in pa])
ps=doc.StyleFamilies.getByName('PageStyles').getByName(sm.PageStyle); print('  landscape:',ps.IsLandscape,'ScaleToPagesX:',ps.ScaleToPagesX,'ScaleToPagesY:',ps.ScaleToPagesY,'PageScale:',ps.PageScale)
print('  Summary col widths (A..R, 1/100 mm):',[sm.getColumns().getByIndex(i).Width for i in range(18)])
print('  row1 height',sm.getRows().getByIndex(0).Height,'row9',sm.getRows().getByIndex(8).Height)
print('  errors:',errs(doc) or 'none')
doc.close(True)
print('=== D. four alternatives / tie on MBT copy')
doc=load(S+'/MBT_check.xlsm'); doc.calculateAll(); db=doc.Sheets.getByName('Database'); sm=doc.Sheets.getByName('Summary')
print('  Database A4:D7 before:',[[txt(db,f'{c}{r}') for c in 'ABCD'] for r in range(4,8)])
db.getCellRangeByName('A6').setString('Alternative 3'); db.getCellRangeByName('B6').setString('HMA-New'); db.getCellRangeByName('D6').setString('Alt 1 (New HMA)')
db.getCellRangeByName('A7').setString('Alternative 4'); db.getCellRangeByName('B7').setString('PCC-New'); db.getCellRangeByName('D7').setString('Alt 2 (New PCC)')
doc.calculateAll()
for r in range(4,8): print(f'   row{r}:',txt(sm,f'H{r}'),txt(sm,f'I{r}'),f'NPW={num(sm,f"O{r}"):,.0f} delta={num(sm,f"P{r}"):,.0f} days={num(sm,f"Q{r}"):.0f}')
print('   G9:',txt(sm,'G9')); print('   G10:',txt(sm,'G10'))
print('   chart cat block row5 (X..AA):',[round(num(sm,f'{c}5')) for c in ['X','Y','Z','AA']],'| sens row32:',[round(num(sm,f'{c}32')) for c in ['X','Y','Z','AA']],'| by-year 2027 spend:',[round(num(sm,f'{c}41')) for c in ['Y','Z','AA','AB']])
print('   errors:',errs(doc) or 'none')
doc.close(True); proc.terminate()
