"""Confirm the New Study macro is built into the workbook, cleanly.

Run: python3 verification/check_newstudy.py <workbook.xlsm>

Checks that the VBA project is a coherent purged project (no stale performance
cache left to fight the source), that the new standard module is present with
its three procedures, and that nothing else in the project moved: the two
UserForms and every ActiveX combo box are still there.
"""
import sys
import olefile
from oletools.olevba import VBA_Parser

WB = sys.argv[1] if len(sys.argv) > 1 else 'TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm'

checks, fails = 0, []


def check(name, ok, detail=''):
    global checks
    checks += 1
    status = 'PASS' if ok else 'FAIL'
    print(f'{status}  {name}   |   {detail}')
    if not ok:
        fails.append(name)


# --- the module and its source -----------------------------------------------
vp = VBA_Parser(WB)
mods = {name: code for _fn, _sp, name, code in vp.extract_macros()
        if not name.endswith('P-code.txt')}
vp.close()

check('the project still has all 24 original modules plus the new one',
      len(mods) == 25, f'{len(mods)} modules')
check('LCCA_NewStudy is one of them',
      'LCCA_NewStudy.bas' in mods, sorted(m for m in mods if 'NewStudy' in m))

code = mods.get('LCCA_NewStudy.bas', '')
for sub in ('Sub NewStudy', 'Sub ClearStudy', 'Function IsAlternativeSheet'):
    check(f'the module defines {sub.split()[-1]}', sub in code)
check('the two UserForms are still in the project',
      'frmAlternativeSetup.frm' in mods and 'frmSelectAltType.frm' in mods)

# --- the project is a clean purge, not a corrupt hybrid ----------------------
import zipfile
with zipfile.ZipFile(WB) as z:
    vba = z.read('xl/vbaProject.bin')
o = olefile.OleFileIO(vba)
paths = ['/'.join(e) for e in o.listdir(streams=True, storages=True)]
srp = [p for p in paths if '__SRP' in p]
check('no stale __SRP performance-cache streams remain', len(srp) == 0, f'{len(srp)} found')
vbaproj = o.openstream('VBA/_VBA_PROJECT').read()
check('_VBA_PROJECT is truncated to its 7-byte header (forces recompile from source)',
      vbaproj == bytes.fromhex('cc61ffff000000'), vbaproj.hex())
# every module offset in dir must be zero for a source-only project
from oletools.olevba import decompress_stream
import struct
dir_dec = bytes(decompress_stream(bytearray(o.openstream('VBA/dir').read())))
offsets, i, nonzero = [], 0, 0
while i + 6 <= len(dir_dec):
    rid, sz = struct.unpack_from('<HI', dir_dec, i)
    if rid == 0x0009:
        i += 12
        continue
    if rid == 0x0031 and sz == 4:
        off = struct.unpack_from('<I', dir_dec, i + 6)[0]
        offsets.append(off)
        nonzero += 1 if off else 0
    i += 6 + sz
    if rid == 0x0010:
        break
check('every module offset in dir is zero (source-only)',
      offsets and nonzero == 0, f'{len(offsets)} modules, {nonzero} nonzero')
o.close()

# --- the ActiveX controls are untouched --------------------------------------
with zipfile.ZipFile(WB) as z:
    ax = [n for n in z.namelist() if n.startswith('xl/activeX/') and n.endswith('.bin')]
combos = 0
for n in ax:
    with zipfile.ZipFile(WB) as z:
        b = z.read(n)
    # a Forms ComboBox control stream carries this CLSID
    if bytes.fromhex('301dd28b42ecce119e0d00aa006002f3') in b:
        combos += 1
check('all 105 ActiveX combo-box controls are still present',
      combos == 105, f'{combos} combos of {len(ax)} controls')

print('\n' + '=' * 78)
print(f'{checks} checks, {len(fails)} failed')
print('ALL PASS' if not fails else 'FAILED: ' + '; '.join(fails))
sys.exit(1 if fails else 0)
