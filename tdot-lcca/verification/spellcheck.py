"""Spell-check every user-visible string in the delivery.

Sources: the workbook (shared strings and inline strings on every visible sheet), the KML macro, the
change notes, the report and deck builders, the design artboards and the verification READMEs.
Domain vocabulary is whitelisted below; anything else is reported with the file it came from.

usage: python3 spellcheck.py <workbook.xlsm> [more files ...]
"""
import sys, re, zipfile, warnings, os, json
from collections import defaultdict
from openpyxl import load_workbook
from spellchecker import SpellChecker
warnings.filterwarnings('ignore')

_OK_SRC = """
lcca tdot faa ara aeronautics npias aptech tasp omb fhwa caltrans realcost aip pgl ndbc
hma pcc npw euac crf pw sy syd sft cy lf ton tons pcf sqft
aashto asphalt asphaltic subbase subgrade geotextile bituminous hma's
mobilization unhide unhidden workbook workbooks worksheet worksheets spreadsheet recalculation
recalculate recalculates recalculated dropdown dropdowns lookup lookups vlookup xlookup sumif sumifs
countif indirect iferror roundup sumproduct activex vba xlsm xlsx kml kmz uno libreoffice openpyxl
json csv xml html svg png pdf docx pptx api cli url urls
gkt mkl mbt srb ckv tys bna mem cha tri mqy jwn dkx syi gcy tha mor nqa xnx
murfreesboro smyrna tullahoma greeneville shelbyville gatlinburg sevierville clarksville
mckellar sipes jackson madison lauderdale nashville knoxville chattanooga memphis
tennessee tennessee's runway runways taxiway apron aprons airfield airfields
lifecycle salvage salvaged discounting discounted undiscounted
pavement pavements rehabilitation rehabilitations
milling mill overlay overlays reclamation slurry emulsified sealcoat
striping markings marking
kpi kpis dashboard dashboards artboard artboards
prefill prefilled prepopulated
sensitivity sensitivities
dcterms datastore
ourairports airportsdata basemap gshhs wdb nasr lid icao iata
ogc kmz lookat timespan timestamp placemark placemarks styleurl
douglas peucker
excel excel's google google's microsoft windows bing
neel schaffer brynick
arial consolas helvetica
lon lat elev bbox
colour colours coloured colouring behaviour favour
analyse analysed organisation organisations recognise
centre metre metres
sqref dxf dxfs
# --- Tennessee place, county and airport names carried by the airport list
blount bomar centerville chilhowee collegedale copperhill covington crossville dekalb dyersburg
eagleville elizabethton fayette fentress gainesboro gallatin hamblen hardeman hassell hohenwald
humphreys jacksboro kingsport lawrenceburg lewisburg madisonville mcghee mckinnon mcminn mcminnville
mcnairy millington morristown murrell obion overton reelfoot rockwood rogersville rossville selmer
sevier sewanee sibley skypark smithville somerville stiner tazewell tipton tiptonville waverly
whitehurst whitson gliderport ebenezer mittlesteadt
# --- standards, programs and reference documents
aaptp acrp apmm apms astm atpb capm ctpb faarfield ncat duah nasr
# --- domain shorthand that appears in headings and notes
centerline sitework constructability tdoa incl verif evid
# --- names the builders write into file and defined-name text
hmarehab newhma newpcc salvagebase revenuedata lostrevenue directcost indirectcost dailyrevenue unitcostgrid payitemkeys pricesources listwidth datablock morphdatacontrol oforms fillable acpa airnav mndot nist pdfs sasp geotextiles
closuredays exportlccakml altchart imagerun xlfn xlnm datamashup sharepoint unpatched overridable
worths befores diff calc stat repo intro const dirname screenshot screenshots hashfile certutil
pypi rels workflow pptxgen pptxgenjs
# --- VBA, KML and OOXML tokens quoted in comments and code
byval cdbl clng cstr elseif freefile instr isnumeric msgbox ucase thisworkbook cdata xmlns bgcolor
cellpadding cellspacing iconstyle labelstyle linearring styleurl combobox comboboxes docname
kmlfilename legendrow pathseparator runwaybearing shortalt writeairport writealternatives
writeevents writelegend writestyle writestyles unparseable isnumber isnumeric recalc ctrl
alignmenttype borderstyle gridlines headinglevel levelformat pagebreak pagenumber shadingtype
tablecell tableofcontents tablerow tabstoptype textrun widthtype valign
# --- the misspellings themselves, quoted in the change notes and the technical record
adminimstration clossures associeted intial
# --- HTML entities used in the design artboards
nbsp ndash mdash ldquo rdquo rsquo minus times deg prime
"""
OK = {w for ln in _OK_SRC.splitlines() if not ln.lstrip().startswith('#') for w in ln.split()}

WORD = re.compile(r"[A-Za-z][A-Za-z']*")
URLISH = re.compile(r'https?://|www\.|\.gov|\.com|\.org|\.edu|\.pdf|\.html|\.xml|\.js|\.py|\.bas')
IDENT = re.compile(r'^[a-z]+[A-Z]|_')          # camelCase or snake_case: code, not prose
sp = SpellChecker()
found = defaultdict(set)


def check(text, where):
    text = str(text or '')
    if not text: return
    for chunk in re.split(r'\s+', text):
        if URLISH.search(chunk): continue                      # links and file names are not prose
        for w in WORD.findall(chunk):
            if IDENT.search(w): continue
            core = w.strip("'").lower()
            if core.endswith("'s"): core = core[:-2]
            if len(core) < 4 or core in OK: continue
            if sp.unknown([core]):
                found[core].add(where)


def from_workbook(path):
    z = zipfile.ZipFile(path)
    wb = load_workbook(path, keep_vba=True)
    for sh in wb.worksheets:
        if sh.sheet_state != 'visible': continue
        for row in sh.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and not v.startswith('='):
                    check(v, f'{os.path.basename(path)}:{sh.title}!{c.coordinate}')
                elif isinstance(v, str):
                    for lit in re.findall(r'"([^"]{4,})"', v):       # text inside formulas
                        check(lit, f'{os.path.basename(path)}:{sh.title}!{c.coordinate} (formula text)')
    for n in z.namelist():                                          # chart titles and axis titles
        if re.match(r'xl/(charts|drawings)/', n) and n.endswith('.xml'):
            for t in re.findall(r'<a:t>([^<]+)</a:t>', z.read(n).decode('utf-8', 'replace')):
                check(t, f'{os.path.basename(path)}:{n}')


def from_text(path):
    s = open(path, encoding='utf-8', errors='replace').read()
    base = os.path.basename(path)
    if path.endswith('.bas'):
        for m in re.findall(r"^\s*'(.*)$", s, re.M) + re.findall(r'"([^"]{6,})"', s):
            check(m, base)
    elif path.endswith('.js'):
        for m in re.findall(r"'((?:[^'\\]|\\.){6,})'", s) + re.findall(r'"((?:[^"\\]|\\.){6,})"', s):
            check(m, base)
    elif path.endswith(('.html', '.dc.html')):
        body = re.sub(r'<style.*?</style>', ' ', s, flags=re.S)
        body = re.sub(r'<svg.*?</svg>', ' ', body, flags=re.S)
        body = re.sub(r'<[^>]+>', ' ', body)
        check(body, base)
    else:
        check(s, base)


for p in sys.argv[1:]:
    (from_workbook if p.endswith(('.xlsm', '.xlsx')) else from_text)(p)

if not found:
    print('no spelling issues found')
else:
    print(f'{len(found)} words to look at\n')
    for w in sorted(found):
        print('%-22s %s' % (w, '; '.join(sorted(found[w])[:3])))
