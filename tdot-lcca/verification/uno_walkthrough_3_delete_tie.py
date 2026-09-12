import uno, time, subprocess
from com.sun.star.beans import PropertyValue
S='/tmp/claude-0/-home-user-lcca-dashboard/444ca227-15d7-51d4-a51e-a0ae3dbe81ed/scratchpad'
def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
proc=subprocess.Popen(['soffice','-env:UserInstallation=file:///tmp/lo_uno_profile3','--headless','--invisible','--norestore','--nologo','--accept=socket,host=localhost,port=2004;urp;'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
for i in range(60):
    try:
        local=uno.getComponentContext(); resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
        ctx=resolver.resolve('uno:socket,host=localhost,port=2004;urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
desktop=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
doc=desktop.loadComponentFromURL(uno.systemPathToFileUrl(S+'/MBT_check.xlsm'),'_blank',0,(pv('Hidden',True),pv('MacroExecutionMode',0))); doc.calculateAll()
num=lambda sh,ref: sh.getCellRangeByName(ref).getValue(); txt=lambda sh,ref: sh.getCellRangeByName(ref).getString()
db=doc.Sheets.getByName('Database'); sm=doc.Sheets.getByName('Summary')
def rows(tag):
    print(f'[{tag}]')
    for r in range(4,8):
        c=sm.getCellRangeByName(f'G{r}'); print(f'   row{r}: err={c.getError()} G={txt(sm,f"G{r}")!r} H={txt(sm,f"H{r}")!r} NPW={num(sm,f"O{r}"):,.0f} days={num(sm,f"Q{r}"):.0f}')
    print('   G9:',txt(sm,'G9')); print('   G4 formula:',sm.getCellRangeByName('G4').getFormula()[:80])
rows('as loaded')
# emulate VBA Workbook_SheetBeforeDelete: wsDatabase.Rows(4).Delete  (delete Alt 1)
db.getRows().removeByIndex(3,1); doc.calculateAll(); rows('after deleting Database row 4 (Alt 1 removed)')
# emulate re-adding: insert row and write Alt 1 back at row 5
db.getRows().insertByIndex(4,1)
for c,v in zip('ABD',['Alternative 1','HMA-New','Alt 1 (New HMA)']): db.getCellRangeByName(f'{c}5').setString(v)
doc.calculateAll(); rows('after re-adding Alt 1 in row 5')
# tie: point row 6 at the same sheet as row 5
for c,v in zip('ABD',['Alternative 3','PCC-New','Alt 2 (New PCC)']): db.getCellRangeByName(f'{c}6').setString(v)
doc.calculateAll(); rows('tie (two rows on the PCC sheet)')
doc.close(True); proc.terminate()
