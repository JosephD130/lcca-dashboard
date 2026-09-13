"""Re-test the shoulder scaling against cells that actually carry a value."""
import uno, time, subprocess
from com.sun.star.beans import PropertyValue
S='/tmp/claude-0/-home-user-lcca-dashboard/444ca227-15d7-51d4-a51e-a0ae3dbe81ed/scratchpad'
def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
proc=subprocess.Popen(['soffice','-env:UserInstallation=file:///tmp/lo_shoulder','--headless','--invisible','--norestore','--nologo','--accept=socket,host=localhost,port=2012;urp;'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
for _ in range(90):
    try:
        local=uno.getComponentContext(); r=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
        ctx=r.resolve('uno:socket,host=localhost,port=2012;urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
desktop=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
doc=desktop.loadComponentFromURL(uno.systemPathToFileUrl(S+'/MBT_check.xlsm'),'_blank',0,(pv('Hidden',True),pv('MacroExecutionMode',0))); doc.calculateAll()
num=lambda sh,ref: sh.getCellRangeByName(ref).getValue(); txt=lambda sh,ref: sh.getCellRangeByName(ref).getString()
gi=doc.Sheets.getByName('General Information'); sm=doc.Sheets.getByName('Summary')
rows={76:'Subbase',77:'Aggregate base',78:'Asphalt',79:'Concrete'}
print('mainline chart data (rows 76-79), alt 1 = X, alt 2 = Y')
for r,lab in rows.items(): print(f'   {lab:16} X{r}={num(sm,f"X{r}"):6.2f}  Y{r}={num(sm,f"Y{r}"):6.2f}')
print('shoulder block with no shoulder area (rows 83-86):')
for r,lab in zip(range(83,87), rows.values()): print(f'   {lab:16} X{r}={num(sm,f"X{r}"):6.2f}  Y{r}={num(sm,f"Y{r}"):6.2f}')
area=num(gi,'D26'); gi.getCellRangeByName('D27').setValue(8000); doc.calculateAll()
f=area/(area+8000)
print(f'\nwith 8,000 S.Y. of shoulder (scale factor {f:.4f}):')
ok=True
for r,lab in zip(range(83,87), rows.values()):
    src=r-7
    for col in ('X','Y'):
        base=num(sm,f'{col}{src}'); got=num(sm,f'{col}{r}')
        want=base if lab=='Concrete' else base*f
        good=abs(got-want)<0.01
        ok&=good
        if base: print(f'   {lab:16} {col}: {base:6.2f} -> {got:6.2f}  expected {want:6.2f}  {"ok" if good else "MISMATCH"}')
print('\nsection table under the shoulder case (the table itself stays on mainline area):')
# the section block rows for the first two alternatives
for r in (93, 94):
    print('   ', [round(v,2) if isinstance(v,float) else v for v in [sm.getCellRangeByName(f'{c}{r}').getValue() if c in 'HIJLM' else sm.getCellRangeByName(f'{c}{r}').getString() for c in 'GHIJKLMNP']])
print('\nRESULT:', 'shoulder scaling verified' if ok else 'FAILED')
gi.getCellRangeByName('D27').setValue(0); doc.close(True); proc.terminate()
