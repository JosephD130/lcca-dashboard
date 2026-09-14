"""Add the LCCA_NewStudy standard module to vbaProject.bin, self-contained.

Method: VBA purge (Mandiant/Didier Stevens), then append the new module as
source-only. Purging removes the version-locked performance cache from every
module, deletes the __SRP_* streams, and truncates _VBA_PROJECT to its 7-byte
header, so Office recompiles all source (existing + new) from scratch on open.
It never touches the UserForm design storages or anything outside this file,
which is why the 105 ActiveX controls and both forms come through intact.
"""
import struct, sys
import cfb, dirparse, ovba
from oletools.olevba import decompress_stream

SRC_BIN = sys.argv[1]                 # vbaProject.bin to read
BAS = sys.argv[2]                     # LCCA_NewStudy.bas
OUT_BIN = sys.argv[3]                 # vbaProject.bin to write
MODULE = 'LCCA_NewStudy'

PURGED_VBA_PROJECT = bytes.fromhex('cc61ffff000000')   # 7-byte header, FFFF version


def main():
    root = cfb.read(SRC_BIN)
    vba = root.find('VBA')

    # --- parse the dir stream -------------------------------------------------
    dir_node = vba.find('dir')
    dir_dec = bytes(decompress_stream(bytearray(dir_node.data)))
    recs, tail = dirparse.parse(dir_dec)
    assert dirparse.build(recs, tail) == dir_dec

    # module stream name -> its MODULEOFFSET, walked from the record list
    offsets, cur = {}, None
    for rid, body in recs:
        if rid == 0x0019:
            cur = None
        elif rid == 0x001A:
            cur = body.decode('latin-1')
        elif rid == 0x0031 and cur is not None:
            offsets[cur] = struct.unpack('<I', body)[0]

    # --- purge every module stream to source-only ----------------------------
    module_streams = {c.name for c in vba.children if c.kind == cfb.STREAM}
    purged = 0
    for name, off in offsets.items():
        node = vba.find(name)
        if off:
            src = bytes(decompress_stream(bytearray(node.data[off:])))
            # re-emit source so the split point is provably clean
            node.data = ovba.compress_checked(src)
            purged += 1
        else:
            # already source-only; normalise through the round-trip anyway
            src = bytes(decompress_stream(bytearray(node.data)))
            node.data = ovba.compress_checked(src)

    # --- delete __SRP_* performance-cache streams -----------------------------
    srp = [c for c in vba.children if c.name.startswith('__SRP_')]
    for c in srp:
        vba.children.remove(c)

    # --- truncate _VBA_PROJECT to its header ----------------------------------
    vba.find('_VBA_PROJECT').data = PURGED_VBA_PROJECT

    # --- zero every MODULEOFFSET in dir ---------------------------------------
    for rec in recs:
        if rec[0] == 0x0031:
            rec[1] = struct.pack('<I', 0)

    # --- bump PROJECTMODULES count --------------------------------------------
    for rec in recs:
        if rec[0] == 0x000F:
            rec[1] = struct.pack('<H', struct.unpack('<H', rec[1])[0] + 1)
            break

    # --- build and insert the new module's dir records ------------------------
    nm = MODULE.encode('latin-1')
    nmu = MODULE.encode('utf-16-le')
    block = [
        [0x0019, nm],                       # MODULENAME
        [0x0047, nmu],                      # MODULENAME unicode
        [0x001A, nm],                       # MODULESTREAMNAME
        [0x0032, nmu],                      # streamname unicode (reserved 0x0032)
        [0x001C, b''],                      # MODULEDOCSTRING
        [0x0048, b''],                      # MODULEDOCSTRING unicode
        [0x0031, struct.pack('<I', 0)],     # MODULEOFFSET = 0 (source-only)
        [0x001E, struct.pack('<I', 0)],     # MODULEHELPCONTEXT
        [0x002C, struct.pack('<H', 0xFFFF)],# MODULECOOKIE (none)
        [0x0021, b''],                      # MODULETYPE = procedural
        [0x002B, b''],                      # MODULE terminator
    ]
    # insert just before the project terminator (0x0010)
    idx = next(i for i, r in enumerate(recs) if r[0] == 0x0010)
    recs[idx:idx] = block

    # recompress dir
    new_dir = dirparse.build(recs, tail)
    dir_node.data = ovba.compress_checked(new_dir)

    # --- add the module source stream ----------------------------------------
    src_text = open(BAS, 'rb').read().replace(b'\r\n', b'\n').replace(b'\n', b'\r\n')
    if not src_text.endswith(b'\r\n'):
        src_text += b'\r\n'
    node = cfb.Node(MODULE, cfb.STREAM, data=ovba.compress_checked(src_text))
    vba.add(node)

    # --- update PROJECT (declarations + workspace) ----------------------------
    proj = root.find('PROJECT').data.decode('latin-1')
    lines = proj.split('\r\n')
    out, added_mod, added_ws, in_ws = [], False, False, False
    for line in lines:
        if line.strip() == '[Workspace]':
            in_ws = True
        if not added_mod and line.startswith('Module=Output'):
            out.append(line)
            out.append(f'Module={MODULE}')
            added_mod = True
            continue
        out.append(line)
    # workspace entry appended at end of file (order there is not significant)
    if not added_ws:
        # place it right after the last existing workspace entry
        for i in range(len(out) - 1, -1, -1):
            if '=' in out[i] and out[i].split('=')[0].strip() and not out[i].startswith('['):
                out.insert(i + 1, f'{MODULE}=0, 0, 0, 0, C')
                added_ws = True
                break
    assert added_mod and added_ws, (added_mod, added_ws)
    root.find('PROJECT').data = '\r\n'.join(out).encode('latin-1')

    # --- update PROJECTwm (name map) ------------------------------------------
    wm = bytearray(root.find('PROJECTwm').data)
    # stream is a sequence of  <mbcs name>\0 <utf16 name>\0\0 , terminated by \0\0
    assert wm.endswith(b'\x00\x00')
    entry = nm + b'\x00' + nmu + b'\x00\x00'
    wm[-2:-2] = entry                        # splice before the final terminator
    root.find('PROJECTwm').data = bytes(wm)

    cfb.write(root, OUT_BIN)
    print(f'purged {purged} modules, removed {len(srp)} __SRP streams, added {MODULE}')


if __name__ == '__main__':
    main()
