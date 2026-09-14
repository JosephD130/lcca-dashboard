"""MS-OVBA 2.4.1 CompressedContainer: the compressor Office expects.

oletools ships a decompressor but no compressor, so this is the other half.
Every output of compress() is checked against oletools' decompress_stream()
before it is written anywhere.
"""
import struct


# Share the decoder's own bit-split so encode and decode cannot drift apart.
from oletools.olevba import copytoken_help as _copytoken_help


def _match(data, chunk_start, pos, end, max_len):
    """Longest match for data[pos:] inside the current chunk. Greedy is plenty
    here: the streams are a few kilobytes and Office only has to read them."""
    best_len, best_off = 0, 0
    limit = min(max_len, end - pos)
    if limit < 3:
        return 0, 0
    for off in range(1, pos - chunk_start + 1):
        src = pos - off
        n = 0
        # Overlapping copies need no special case: the bytes the decompressor
        # would read back are the ones already standing at those positions.
        while n < limit and data[src + n] == data[pos + n]:
            n += 1
        if n > best_len:
            best_len, best_off = n, off
            if n == limit:
                break
    return (best_len, best_off) if best_len >= 3 else (0, 0)


def _compress_chunk(data, start, end):
    """Token-encode one chunk of decompressed data. Returns (bytes, consumed)."""
    out = bytearray()
    pos = start
    while pos < end and (pos - start) < 4096:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= end or (pos - start) >= 4096:
                break
            if pos == start:
                # Nothing behind us in this chunk yet, so the first byte of a
                # chunk is always a literal.
                length = offset = 0
            else:
                _, _, bit_count, max_len = _copytoken_help(pos, start)
                length, offset = _match(data, start, pos,
                                        min(end, start + 4096), max_len)
            if length:
                token = ((offset - 1) << (16 - bit_count)) | (length - 3)
                out += struct.pack('<H', token)
                flags |= 1 << bit
                pos += length
            else:
                out.append(data[pos])
                pos += 1
        out[flag_pos] = flags
    return bytes(out), pos - start


def compress(data):
    """Whole-stream compression. Signature byte then chunk headers + tokens."""
    data = bytes(data)
    out = bytearray(b'\x01')
    pos = 0
    while pos < len(data):
        body, consumed = _compress_chunk(data, pos, len(data))
        raw = data[pos:pos + consumed]
        if len(body) < len(raw):
            header = 0xB000 | ((len(body) - 1) & 0x0FFF)
            out += struct.pack('<H', header) + body
        else:
            # An uncompressed chunk is always a full 4096 bytes, so it is only
            # usable when the chunk really is full.
            if len(raw) == 4096:
                out += struct.pack('<H', 0x3FFF) + raw
            else:
                header = 0xB000 | ((len(body) - 1) & 0x0FFF)
                out += struct.pack('<H', header) + body
        pos += consumed
    return bytes(out)


def compress_checked(data):
    """compress(), but refuse to hand back anything that does not round-trip."""
    from oletools.olevba import decompress_stream
    blob = compress(data)
    back = decompress_stream(bytearray(blob))
    if bytes(back) != bytes(data):
        raise AssertionError('compressed stream does not round-trip')
    return blob
