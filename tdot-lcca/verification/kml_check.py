"""Lint a KML file the way Google Earth and a strict GIS reader would, and simulate the time slider.

No network: the OGC schema is not reachable from here, so the rules that matter are implemented
directly, including the element order the KML 2.2 sequence requires inside a Feature.

usage: python3 kml_check.py <file.kml>
"""
import sys, re, math
from lxml import etree

K = 'http://www.opengis.net/kml/2.2'
NS = {'k': K}
# kml:AbstractFeatureType sequence, in order
FEATURE_ORDER = ['name', 'visibility', 'open', 'author', 'link', 'address', 'AddressDetails', 'phoneNumber',
                 'Snippet', 'snippet', 'description', 'LookAt', 'Camera', 'TimeStamp', 'TimeSpan', 'styleUrl',
                 'Style', 'StyleMap', 'Region', 'Metadata', 'ExtendedData']
GEOM = {'Point', 'LineString', 'LinearRing', 'Polygon', 'MultiGeometry', 'Model', 'Track'}
CONTAINER_EXTRA = {'Document': ['Schema'], 'Folder': [], 'Placemark': list(GEOM), 'GroundOverlay': ['altitude', 'altitudeMode', 'LatLonBox', 'Icon', 'color', 'drawOrder'],
                   'ScreenOverlay': ['overlayXY', 'screenXY', 'rotationXY', 'size', 'rotation', 'Icon', 'color', 'drawOrder']}
FEATURES = {'Document', 'Folder', 'Placemark', 'GroundOverlay', 'ScreenOverlay', 'PhotoOverlay', 'NetworkLink'}

issues, notes = [], []
def bad(msg): issues.append(msg)
def note(msg): notes.append(msg)


def tag(e): return etree.QName(e).localname


def check_feature_order(e):
    seen, last = [], -1
    for c in e:
        if not isinstance(c.tag, str): continue
        t = tag(c)
        if t in FEATURES or t in CONTAINER_EXTRA.get(tag(e), []): break_pos = True
        if t in FEATURE_ORDER:
            i = FEATURE_ORDER.index(t)
            if i < last:
                bad(f'{tag(e)} "{name_of(e)}": <{t}> comes after <{FEATURE_ORDER[last]}>; the KML 2.2 sequence '
                    f'wants {" then ".join(x for x in FEATURE_ORDER if x in seen + [t])}')
            last = max(last, i); seen.append(t)


def name_of(e):
    n = e.find('k:name', NS)
    return (n.text or '')[:40] if n is not None else ''


def coords_of(text):
    out = []
    for tok in text.split():
        p = tok.split(',')
        if len(p) < 2: bad(f'coordinate tuple "{tok}" has fewer than two numbers'); continue
        try: out.append(tuple(float(x) for x in p))
        except ValueError: bad(f'coordinate tuple "{tok}" is not numeric')
    return out


def main(path):
    raw = open(path, 'rb').read()
    try:
        root = etree.fromstring(raw)
    except etree.XMLSyntaxError as e:
        print('FAIL  not well-formed XML:', e); return 1
    print(f'file            {path}  ({len(raw):,} bytes)')
    print(f'root            {{{etree.QName(root).namespace}}}{tag(root)}')
    if etree.QName(root).namespace != K: bad('root namespace is not the KML 2.2 namespace')

    # ---- structure
    feats = [e for e in root.iter() if isinstance(e.tag, str) and tag(e) in FEATURES]
    for e in feats: check_feature_order(e)
    pms = root.findall('.//k:Placemark', NS)
    for pm in pms:
        g = [c for c in pm if isinstance(c.tag, str) and tag(c) in GEOM]
        if not g: bad(f'Placemark "{name_of(pm)}" has no geometry')

    # ---- styles resolve
    ids = {s.get('id') for s in root.findall('.//k:Style', NS) + root.findall('.//k:StyleMap', NS)}
    for u in root.findall('.//k:styleUrl', NS):
        ref = (u.text or '').strip()
        if ref.startswith('#') and ref[1:] not in ids: bad(f'styleUrl {ref} points at a style that is not in the file')
    print(f'styles          {len(ids)} defined: {", ".join(sorted(x for x in ids if x))}')

    # ---- colours are aabbggrr hex
    for c in root.findall('.//k:color', NS):
        if not re.fullmatch(r'[0-9a-fA-F]{8}', (c.text or '').strip()):
            bad(f'colour "{c.text}" is not 8 hex digits (aabbggrr)')

    # ---- geometry
    rings = root.findall('.//k:LinearRing', NS)
    for r in rings:
        cs = coords_of(r.find('k:coordinates', NS).text)
        if len(cs) < 4: bad(f'LinearRing has only {len(cs)} points')
        elif cs[0][:2] != cs[-1][:2]: bad('LinearRing is not closed (first point != last point)')
    pts = [coords_of(c.text)[0] for c in root.findall('.//k:Point/k:coordinates', NS)]
    allc = pts + [c for r in rings for c in coords_of(r.find('k:coordinates', NS).text)]
    for lon, lat, *rest in allc:
        if not -180 <= lon <= 180 or not -90 <= lat <= 90: bad(f'coordinate out of range: {lon},{lat}')
    lons = [c[0] for c in allc]; lats = [c[1] for c in allc]
    print(f'geometry        {len(pts)} points, {len(rings)} rings, extent lon {min(lons):.5f}..{max(lons):.5f} '
          f'lat {min(lats):.5f}..{max(lats):.5f}')
    span_m = (max(lats) - min(lats)) * 111320
    if span_m > 20000: bad(f'features spread over {span_m/1000:.1f} km, which is too wide for one airport')
    else: print(f'                north-south spread {span_m:,.0f} m')

    # ---- polygons that stand up
    for p in root.findall('.//k:Polygon', NS):
        ex = p.find('k:extrude', NS)
        if ex is not None and (ex.text or '').strip() == '1':
            am = p.find('k:altitudeMode', NS)
            if am is None or (am.text or '').strip() not in ('relativeToGround', 'absolute'):
                bad('an extruded polygon has no altitudeMode of relativeToGround or absolute, so it will lie flat')
            alts = {c[2] for c in coords_of(p.find('.//k:coordinates', NS).text) if len(c) > 2}
            if alts == {0.0}: bad('an extruded polygon sits at altitude 0, so it has no height')

    # ---- the time slider
    tps = root.findall('.//k:TimeSpan', NS) + root.findall('.//k:TimeStamp', NS)
    for t in tps:
        parent = t.getparent()
        if tag(parent) not in FEATURES: bad(f'a {tag(t)} hangs off <{tag(parent)}>, which the slider ignores')
    years = []
    timed = []
    for pm in pms:
        ts = pm.find('k:TimeSpan', NS)
        if ts is None: continue
        b = (ts.findtext('k:begin', namespaces=NS) or '').strip()
        e = (ts.findtext('k:end', namespaces=NS) or '').strip()
        for v in (b, e):
            if not re.fullmatch(r'\d{4}(-\d{2}(-\d{2})?)?(T.*)?', v): bad(f'time value "{v}" is not an ISO 8601 date')
        if b and e and b > e: bad(f'TimeSpan begin {b} is after end {e}')
        years.append(int(b[:4])); timed.append((int(b[:4]), name_of(pm)))
    untimed = len(pms) - len(timed)
    if not years:
        bad('no feature carries a time, so Google Earth shows no time slider')
    else:
        print(f'time            {len(timed)} timed placemarks, slider spans {min(years)} to {max(years)}; '
              f'{untimed} features always visible')

    # ---- what the slider shows, year by year
    if years:
        print('\nTIME SLIDER WALK (what appears as you scroll)')
        for y in range(min(years), max(years) + 1):
            here = [n for yy, n in timed if yy == y]
            if here: print(f'  {y}  ' + '; '.join(here))

    # ---- a legend?
    doc = root.find('k:Document', NS)
    desc = (doc.findtext('k:description', namespaces=NS) or '') if doc is not None else ''
    if 'legend' not in desc.lower() and not root.findall('.//k:ScreenOverlay', NS):
        note('no legend: nothing in the file tells the reader what the colours mean')
    if doc is not None and doc.find('k:LookAt', NS) is None and doc.find('k:Camera', NS) is None:
        note('no LookAt on the Document, so opening the file does not fly anywhere')

    print()
    for n in notes: print('NOTE  ' + n)
    for i in issues: print('ISSUE ' + i)
    print(f'\n{len(issues)} issues, {len(notes)} notes')
    return 1 if issues else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1]))
