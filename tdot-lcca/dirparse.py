"""Parse and rebuild the MS-OVBA 'dir' stream as a flat record list.

Every record is (id, bytes). Rebuilding is just concatenation, so edits are
made by splicing the record list, never by blind byte-poking.
"""
import struct

# Records whose size field the writer of this project used; the reader trusts
# the on-disk size field regardless, so this table is only for readable dumps.
NAMES = {
    0x0001: 'SYSKIND', 0x004A: 'COMPATVERSION', 0x0002: 'LCID',
    0x0014: 'LCIDINVOKE', 0x0003: 'CODEPAGE', 0x0004: 'NAME',
    0x0005: 'DOCSTRING', 0x0040: 'DOCSTRING_U', 0x0006: 'HELPFILE1',
    0x003D: 'HELPFILE2', 0x0007: 'HELPCONTEXT', 0x0008: 'LIBFLAGS',
    0x0009: 'VERSION', 0x000C: 'CONSTANTS', 0x003C: 'CONSTANTS_U',
    0x000F: 'PROJECTMODULES', 0x0013: 'PROJECTCOOKIE', 0x0016: 'REFNAME',
    0x0019: 'MODULENAME', 0x0047: 'MODULENAME_U', 0x001A: 'STREAMNAME',
    0x0032: 'STREAMNAME_U', 0x001C: 'MODULEDOCSTRING', 0x0048: 'MODULEDOCSTRING_U',
    0x0031: 'MODULEOFFSET', 0x001E: 'MODULEHELPCONTEXT', 0x002C: 'MODULECOOKIE',
    0x0021: 'MODULETYPE_STD', 0x0022: 'MODULETYPE_DOC', 0x0025: 'MODULEREADONLY',
    0x0028: 'MODULEPRIVATE', 0x002B: 'MODULE_TERM', 0x0010: 'TERMINATOR',
}
# The VERSION record (0x0009) has a fixed 4-byte reserved size and its size
# field reads 4 but the layout is special; treat it like any other 6+size.


def parse(d):
    """Return list of [id, body]. Faithful: b''.join rebuilds d exactly."""
    records = []
    i = 0
    n = len(d)
    while i + 6 <= n:
        rid, sz = struct.unpack_from('<HI', d, i)
        if rid == 0x0009:
            # PROJECTVERSION: its Size field reads 4 but the record carries a
            # 4-byte major + 2-byte minor after it, so the real body is 6 bytes.
            body = d[i + 6:i + 12]
            records.append([rid, body])
            i += 12
            continue
        body = d[i + 6:i + 6 + sz]
        records.append([rid, body])
        i += 6 + sz
        if rid == 0x0010:                       # project terminator
            break
    # Anything trailing the terminator (reserved) is kept verbatim.
    tail = d[i:]
    return records, tail


def build(records, tail=b''):
    out = bytearray()
    for rid, body in records:
        if rid == 0x0009:
            # Reserved size field is always 4, regardless of the 6-byte body.
            out += struct.pack('<HI', rid, 4) + bytes(body)
        else:
            out += struct.pack('<HI', rid, len(body)) + bytes(body)
    return bytes(out) + tail
