"""New-project end-to-end test: Outlaw Field (CKV), Runway 17-35 reconstruction, HMA vs PCC.
Emulates what the Alternative Setup form does (copy hidden templates, register on Database, write Summary A:E),
fills General Information and the pay items, recalculates, and reads the results."""
import uno, time, subprocess, json, math
from com.sun.star.beans import PropertyValue
S='/tmp/claude-0/-home-user-lcca-dashboard/444ca227-15d7-51d4-a51e-a0ae3dbe81ed/scratchpad'
def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
proc=subprocess.Popen(['soffice','-env:UserInstallation=file:///tmp/lo_uno_profile5','--headless','--invisible','--norestore','--nologo','--accept=socket,host=localhost,port=2005;urp;'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
for i in range(60):
    try:
        local=uno.getComponentContext(); resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
        ctx=resolver.resolve('uno:socket,host=localhost,port=2005;urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
desktop=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
doc=desktop.loadComponentFromURL(uno.systemPathToFileUrl(S+'/TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm'),'_blank',0,(pv('Hidden',True),pv('MacroExecutionMode',0)))
num=lambda sh,ref: sh.getCellRangeByName(ref).getValue(); txt=lambda sh,ref: sh.getCellRangeByName(ref).getString()
def setv(sh,ref,v):
    c=sh.getCellRangeByName(ref)
    if isinstance(v,str): c.setString(v)
    else: c.setValue(v)
# ---- scenario
SC=dict(airport='Outlaw Field', id='CKV', region='Middle', consultant='ARA', projno='5936', projname='Runway 17-35 Reconstruction',
        branch_type='Runway', branch='Runway 17-35', ptype='Reconstruction', desc='Full-depth reconstruction, 6,000 x 100 ft',
        year=2028, area=66667, shoulder=0, markings=15000, mark_type='Reflective', period=30, rate=3, mob=10, eng=5, indirect='Yes')
HMA=[('Lime Treated subgrade',66667),('Subbase Course',14815),('Crushed Aggregate Base Course',11111),('Asphalt Base Course',18333),('Asphalt Surface Course',14667),('Markings (prep, markings, reflective media)',15000)]
PCC=[('Lime Treated subgrade',66667),('Subbase Course',14815),('Crushed Aggregate Base Course',11111),('Concrete Pavement, 9-inch',66667),('Markings (prep, markings, reflective media)',15000)]
gi=doc.Sheets.getByName('General Information')
for ref,v in [('D9',SC['airport']),('D14',SC['consultant']),('D15',SC['projno']),('D16',SC['projname']),('D21',SC['branch_type']),('D22',SC['branch']),('D23',SC['ptype']),('D24',SC['desc']),
              ('D25',SC['year']),('D26',SC['area']),('D27',SC['shoulder']),('D28',SC['markings']),('D29',SC['mark_type']),('D33',SC['period']),('D34',SC['rate']),('D36',SC['mob']),('D37',SC['eng']),('D38',SC['indirect'])]:
    setv(gi,ref,v)
# ---- emulate Alternative Setup: copy the hidden indirect-cost templates (D38 = Yes) to Alt sheets
n=doc.Sheets.getCount()
doc.Sheets.copyByName('TMP(NewHMA)_IndirectCost','Alt 1 (New HMA)',n); doc.Sheets.copyByName('TMP(NewPCC)_IndirectCost','Alt 2 (New PCC)',n+1)
a1=doc.Sheets.getByName('Alt 1 (New HMA)'); a2=doc.Sheets.getByName('Alt 2 (New PCC)'); a1.IsVisible=True; a2.IsVisible=True
for sh,items in [(a1,HMA),(a2,PCC)]:
    for k,(desc,qty) in enumerate(items): setv(sh,f'C{13+k}',desc); setv(sh,f'E{13+k}',qty)
db=doc.Sheets.getByName('Database'); sm=doc.Sheets.getByName('Summary')
for r,(name,typ,ws) in zip((4,5),[('Alternative 1','HMA-New','Alt 1 (New HMA)'),('Alternative 2','PCC-New','Alt 2 (New PCC)')]):
    setv(db,f'A{r}',name); setv(db,f'B{r}',typ); setv(db,f'C{r}','New '+typ.split('-')[0]); setv(db,f'D{r}',ws)
    setv(sm,f'A{r}',f'Alt {r-3}'); setv(sm,f'B{r}',name); sm.getCellRangeByName(f'C{r}').setFormula(f"=$'{ws}'.G27"); sm.getCellRangeByName(f'D{r}').setFormula(f"=$'{ws}'.E{52 if r==4 else 43}"); setv(sm,f'E{r}','New '+typ.split('-')[0])
doc.calculateAll()
# ---- read back
out={'scenario':SC,'gi':{k:txt(gi,k) for k in ['D10','D11','D12','D13','D39']},'alts':{},'summary':{}}
for sh,tag,last in [(a1,'HMA',52),(a2,'PCC',43)]:
    d={'F2':num(sh,'F2'),'G2':txt(sh,'G2'),'durations':{f'F{r}':num(sh,f'F{r}') for r in range(4,11) if txt(sh,f'E{r}')},
       'items':[(txt(sh,f'B{r}'),txt(sh,f'C{r}'),txt(sh,f'D{r}'),num(sh,f'E{r}'),num(sh,f'F{r}'),num(sh,f'G{r}')) for r in range(13,23) if txt(sh,f'C{r}')],
       'subtotal':num(sh,'G24'),'mob':num(sh,'G25'),'eng':num(sh,'G26'),'total':num(sh,'G27'),
       'activities':[(txt(sh,f'B{r}'),num(sh,f'C{r}'),num(sh,f'D{r}'),num(sh,f'E{r}')) for r in range(36,last) if txt(sh,f'B{r}')],
       'NPW':num(sh,f'E{last}'), 'errors':sum(1 for r in range(1,90) for c in range(1,60) if sh.getCellByPosition(c-1,r-1).getError())}
    out['alts'][tag]=d
for r in (4,5):
    out['summary'][f'row{r}']={c:(num(sm,f'{c}{r}') if c not in 'GHI' else txt(sm,f'{c}{r}')) for c in 'GHIJKLMNOPQR'}
out['summary']['G9']=txt(sm,'G9'); out['summary']['G10']=txt(sm,'G10')
out['summary']['sens']={rate:[num(sm,f'X{r}'),num(sm,f'Y{r}')] for r,rate in [(12,2),(16,3),(24,5),(32,7),(36,8)]}
out['summary']['byyear']=[(int(num(sm,f'X{r}')),num(sm,f'Y{r}'),num(sm,f'Z{r}'),num(sm,f'AC{r}'),num(sm,f'AD{r}'),num(sm,f'AG{r}'),num(sm,f'AH{r}')) for r in range(41,72)]
out['section']=[[ (sm.getCellRangeByName(f'{c}{r}').getString() if c in 'GKNO' else num(sm,f'{c}{r}')) for c in 'GHIJKLMNO'] for r in (83,84)]
out['section_unitweight']=num(sm,'J81')
out['section_chart']=[[ (txt(sm,f'{c}{r}') if c=='W' else num(sm,f'{c}{r}')) for c in ['W','X','Y']] for r in range(75,87)]
gi.getCellRangeByName('D27').setValue(8000); doc.calculateAll()
out['section_shoulder']=[[ (txt(sm,f'{c}{r}') if c=='W' else num(sm,f'{c}{r}')) for c in ['W','X','Y']] for r in range(82,87)]
out['section_shoulder_table']=[[ (sm.getCellRangeByName(f'{c}{r}').getString() if c in 'GKNO' else num(sm,f'{c}{r}')) for c in 'GHIJKLMNO'] for r in (83,84)]
gi.getCellRangeByName('D27').setValue(0); doc.calculateAll()
out['summary']['catsum']=[sum(num(sm,f'{c}{k}') for k in range(5,10)) for c in 'XY']
out['summary']['AtoE']=[[txt(sm,f'{c}{r}') if c in 'ABE' else num(sm,f'{c}{r}') for c in 'ABCDE'] for r in (4,5)]
errs={}
for sh in doc.Sheets:
    cur=sh.createCursor(); cur.gotoEndOfUsedArea(False); ra=cur.getRangeAddress(); n=0
    rng=sh.getCellRangeByPosition(0,0,ra.EndColumn,ra.EndRow)
    for rr in range(ra.EndRow+1):
        for cc in range(ra.EndColumn+1):
            cell=rng.getCellByPosition(cc,rr)
            if cell.getType().value=='FORMULA' and cell.getError(): n+=1
    if n: errs[sh.Name]=n
out['errors']=errs
json.dump(out,open(S+'/newproj/ckv_results.json','w'),indent=1,default=str)
# ---- save a viewable copy and a PDF
doc.storeToURL(uno.systemPathToFileUrl(S+'/newproj/CKV_Runway17-35_LCCA_run.xlsx'),(pv('FilterName','Calc MS Excel 2007 XML'),))
doc.getCurrentController().setActiveSheet(sm)
doc.storeToURL(uno.systemPathToFileUrl(S+'/newproj/CKV_run.pdf'),(pv('FilterName','calc_pdf_Export'),))
doc.close(True); proc.terminate()
print(json.dumps({k:v for k,v in out.items() if k!='summary'},indent=1,default=str)[:6000])
print('SUMMARY rows:'); [print('  ',k,{kk:(round(vv) if isinstance(vv,float) else vv) for kk,vv in v.items()}) for k,v in out['summary'].items() if k.startswith('row')]
print('  G9:',out['summary']['G9']); print('  G10:',out['summary']['G10'])
print('  sens:',{k:[round(x) for x in v] for k,v in out['summary']['sens'].items()})
print('  catsum:',[round(x) for x in out['summary']['catsum']]); print('  A:E',out['summary']['AtoE'])
print('  by-year nonzero:',[(y,round(a),round(b),round(e),round(f)) for y,a,b,c,d,e,f in out['summary']['byyear'] if a or b or e or f]); print('  cum end:',[round(out['summary']['byyear'][-1][3]),round(out['summary']['byyear'][-1][4])])
