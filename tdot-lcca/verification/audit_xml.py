import zipfile, re, sys, collections
import xml.etree.ElementTree as ET
path=sys.argv[1]; z=zipfile.ZipFile(path); names=set(z.namelist())
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','ct':'http://schemas.openxmlformats.org/package/2006/content-types','pr':'http://schemas.openxmlformats.org/package/2006/relationships'}
issues=[]
def col2n(c):
    n=0
    for ch in c: n=n*26+ord(ch)-64
    return n
# content types
ct=ET.fromstring(z.read('[Content_Types].xml'))
overrides={o.get('PartName') for o in ct.findall('ct:Override',NS)}; defaults={d.get('Extension').lower() for d in ct.findall('ct:Default',NS)}
for n in names:
    if n.endswith('/') : continue
    ext=n.rsplit('.',1)[-1].lower()
    if '/'+n not in overrides and ext not in defaults: issues.append(f'no content type for {n}')
for o in overrides:
    if o[1:] not in names: issues.append(f'content type override for missing part {o}')
if 'xl/calcChain.xml' in names: issues.append('calcChain present')
# rels targets exist
for n in [x for x in names if x.endswith('.rels')]:
    base=n.replace('_rels/','').replace('.rels','')
    d='/'.join(base.split('/')[:-1])
    for rel in ET.fromstring(z.read(n)).findall('pr:Relationship',NS):
        if rel.get('TargetMode')=='External': continue
        t=rel.get('Target'); 
        if t.startswith('/'): p=t[1:]
        else:
            parts=(d+'/'+t).split('/'); out=[]
            for s in parts:
                if s=='..': out.pop()
                elif s and s!='.': out.append(s)
            p='/'.join(out)
        if p not in names: issues.append(f'{n}: target {t} -> {p} missing')
# workbook
wb=ET.fromstring(z.read('xl/workbook.xml'))
sheets=wb.find('m:sheets',NS).findall('m:sheet',NS)
bv=wb.find('m:bookViews/m:workbookView',NS); active=int(bv.get('activeTab','0')) if bv is not None else 0
print('sheets',len(sheets),'activeTab',active,'->',sheets[active].get('name'))
wrels={r.get('Id'):r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels')).findall('pr:Relationship',NS)}
dn=collections.Counter((d.get('name'),d.get('localSheetId')) for d in wb.findall('m:definedNames/m:definedName',NS))
for k,v in dn.items():
    if v>1: issues.append(f'duplicate definedName {k}')
for d in wb.findall('m:definedNames/m:definedName',NS):
    if d.get('name')=='_xlnm.Print_Area': print('  Print_Area localSheetId',d.get('localSheetId'),'=',sheets[int(d.get('localSheetId'))].get('name'),':',d.text)
    if '[' in (d.text or '') : issues.append(f'definedName {d.get("name")} still external: {d.text[:60]}')
# styles
st=ET.fromstring(z.read('xl/styles.xml'))
for tag in ['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs']:
    e=st.find('m:'+tag,NS)
    if e is not None and e.get('count') is not None and int(e.get('count'))!=len(list(e)): issues.append(f'styles {tag} count={e.get("count")} actual={len(list(e))}')
ncellxfs=len(list(st.find('m:cellXfs',NS)))
numfmt_ids={int(n.get('numFmtId')) for n in st.findall('m:numFmts/m:numFmt',NS)}
for xf in st.findall('m:cellXfs/m:xf',NS):
    nf=int(xf.get('numFmtId',0))
    if nf>=164 and nf not in numfmt_ids: issues.append(f'xf numFmtId {nf} undefined')
    if int(xf.get('fontId',0))>=len(list(st.find('m:fonts',NS))): issues.append('xf fontId out of range')
    if int(xf.get('fillId',0))>=len(list(st.find('m:fills',NS))): issues.append('xf fillId out of range')
    if int(xf.get('borderId',0))>=len(list(st.find('m:borders',NS))): issues.append('xf borderId out of range')
# shared strings
ss=ET.fromstring(z.read('xl/sharedStrings.xml')); nss=len(ss.findall('m:si',NS))
if ss.get('uniqueCount') and int(ss.get('uniqueCount'))!=nss: issues.append(f'sharedStrings uniqueCount {ss.get("uniqueCount")} vs {nss}')
# worksheets
selected=[]
for sh in sheets:
    part='xl/'+wrels[sh.get('{%s}id'%NS['r'])]
    x=ET.fromstring(z.read(part))
    sv=x.find('m:sheetViews/m:sheetView',NS)
    if sv is not None and sv.get('tabSelected')=='1': selected.append(sh.get('name'))
    lastr=0
    for row in x.findall('m:sheetData/m:row',NS):
        r=int(row.get('r'))
        if r<=lastr: issues.append(f'{sh.get("name")}: row order {lastr}->{r}')
        lastr=r; lastc=0
        for c in row.findall('m:c',NS):
            ref=c.get('r'); m=re.match(r'([A-Z]+)(\d+)$',ref)
            if not m or int(m.group(2))!=r: issues.append(f'{sh.get("name")}: cell {ref} in row {r}'); continue
            cn=col2n(m.group(1))
            if cn<=lastc: issues.append(f'{sh.get("name")}: cell order at {ref}')
            lastc=cn
            s=c.get('s')
            if s and int(s)>=ncellxfs: issues.append(f'{sh.get("name")} {ref}: style {s} >= {ncellxfs}')
            if c.get('t')=='s':
                v=c.find('m:v',NS)
                if v is None or int(v.text)>=nss: issues.append(f'{sh.get("name")} {ref}: shared string index bad')
            f=c.find('m:f',NS)
            if f is not None and f.get('t')=='shared' and f.get('ref') is None and f.get('si') is None: issues.append(f'{sh.get("name")} {ref}: shared formula without si')
    # merged cells overlapping? skip. data validations sqref sanity
    for dv in x.findall('m:dataValidations/m:dataValidation',NS):
        f1=dv.find('m:formula1',NS)
        if f1 is not None and '#REF' in (f1.text or ''): issues.append(f'{sh.get("name")} dv #REF')
print('tabSelected sheets:',selected)
if len(selected)>1: issues.append(f'multiple sheets tabSelected (Excel opens them grouped): {selected}')
print('ISSUES:' if issues else 'no structural issues'); [print('  -',i) for i in issues]
