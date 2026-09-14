"""Re-parse every ActiveX stream with an independent implementation of [MS-OFORMS] (oletools'
oleform) and confirm it consumes the stream exactly: a MorphDataControl whose cbMorphData or
DataBlock padding were wrong would either raise or stop short of the TextProps at the end."""
import io, os, struct, sys, zipfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ol'))
from oletools.oleform import ExtendedStream, consume_MorphDataControl

src = sys.argv[1]
ok = bad = 0
widths = set()
with zipfile.ZipFile(src) as z:
    for n in sorted(x for x in z.namelist() if x.startswith('xl/activeX/') and x.endswith('.bin')):
        d = z.read(n)
        if d[:16] != bytes.fromhex('301dd28b42ecce119e0d00aa006002f3'):
            continue
        mask = struct.unpack('<Q', d[20:28])[0]
        assert mask & (1 << 10), '%s has no fListWidth' % n
        widths.add(struct.unpack('<i', d[36:40])[0])
        st = ExtendedStream(io.BytesIO(d[16:]), n)
        try:
            consume_MorphDataControl(st)
        except Exception as e:
            print('FAIL', n, e); bad += 1; continue
        left = st._stream.read()
        if left:
            print('FAIL', n, 'trailing bytes:', len(left)); bad += 1
        else:
            ok += 1
print('combo streams re-parsed clean: %d, failed: %d' % (ok, bad))
print('ListWidth values (HIMETRIC):', sorted(widths), '=', [round(w / 2540.0, 2) for w in sorted(widths)], 'inches')
sys.exit(1 if bad else 0)
