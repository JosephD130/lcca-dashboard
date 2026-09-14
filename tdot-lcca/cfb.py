"""A minimal Compound File Binary writer (MS-CFB, version 3).

olefile reads these but cannot add a stream, and adding a stream is exactly
what putting a new VBA module into vbaProject.bin requires. Storages nest, so
the directory is a real tree, not a flat list.
"""
import struct

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF

SECTOR = 512
MINISECTOR = 64
MINI_CUTOFF = 4096

STORAGE, STREAM, ROOT = 1, 2, 5


class Node:
    def __init__(self, name, kind, clsid=b'\x00' * 16, data=None,
                 create_time=0, modify_time=0):
        self.name = name
        self.kind = kind
        self.clsid = clsid
        self.data = data if data is not None else b''
        self.create_time = create_time
        self.modify_time = modify_time
        self.children = []
        self.sid = NOSTREAM
        self.left = NOSTREAM
        self.right = NOSTREAM
        self.child = NOSTREAM

    def add(self, node):
        self.children.append(node)
        return node

    def find(self, path):
        node = self
        for part in path.split('/'):
            node = next(c for c in node.children if c.name == part)
        return node


def _sort_key(node):
    """MS-CFB directory order: shorter names first, then uppercased."""
    return (len(node.name), node.name.upper())


def read(path_or_bytes):
    """Read a compound file into a Node tree using olefile."""
    import olefile
    ole = olefile.OleFileIO(path_or_bytes)
    root = Node('Root Entry', ROOT)
    by_path = {(): root}
    entries = sorted(ole.listdir(streams=True, storages=True), key=len)
    for entry in entries:
        parent = by_path[tuple(entry[:-1])]
        full = '/'.join(entry)
        direntry = ole.direntries[ole._find(full)]
        clsid = getattr(direntry, 'clsid', None)
        raw_clsid = _clsid_bytes(clsid)
        if ole.get_type(full) == olefile.STGTY_STREAM:
            node = Node(entry[-1], STREAM, raw_clsid, ole.openstream(full).read())
        else:
            node = Node(entry[-1], STORAGE, raw_clsid)
        by_path[tuple(entry)] = parent.add(node)
    rootdir = ole.direntries[0]
    root.clsid = _clsid_bytes(getattr(rootdir, 'clsid', None))
    ole.close()
    return root


def _clsid_bytes(clsid):
    """olefile hands back a CLSID string; the file wants the 16 raw bytes."""
    if not clsid or set(clsid) <= {'0', '-'}:
        return b'\x00' * 16
    import uuid
    return uuid.UUID(clsid).bytes_le


def _flatten(root):
    order = []

    def walk(node):
        node.sid = len(order)
        order.append(node)
        for child in node.children:
            walk(child)

    walk(root)

    def link(children):
        if not children:
            return NOSTREAM
        mid = len(children) // 2
        node = children[mid]
        node.left = link(children[:mid])
        node.right = link(children[mid + 1:])
        return node.sid

    for node in order:
        node.child = link(sorted(node.children, key=_sort_key))
    return order


def _chain(fat, sectors):
    for a, b in zip(sectors, sectors[1:]):
        fat[a] = b
    if sectors:
        fat[sectors[-1]] = ENDOFCHAIN


def write(root, path):
    order = _flatten(root)

    # Split the streams by the mini-stream cutoff.
    big = [n for n in order if n.kind == STREAM and len(n.data) >= MINI_CUTOFF]
    small = [n for n in order if n.kind == STREAM and 0 < len(n.data) < MINI_CUTOFF]

    mini = bytearray()
    mini_start = {}
    for node in small:
        mini_start[node.sid] = len(mini) // MINISECTOR
        mini += node.data
        if len(mini) % MINISECTOR:
            mini += b'\x00' * (MINISECTOR - len(mini) % MINISECTOR)

    n_mini_sectors = len(mini) // MINISECTOR
    minifat = [FREESECT] * n_mini_sectors
    for node in small:
        first = mini_start[node.sid]
        count = (len(node.data) + MINISECTOR - 1) // MINISECTOR
        _chain(minifat, list(range(first, first + count)))

    def n_sectors(length):
        return (length + SECTOR - 1) // SECTOR

    minifat_bytes = b''.join(struct.pack('<I', v) for v in minifat)
    n_data = sum(n_sectors(len(n.data)) for n in big)
    n_ministream = n_sectors(len(mini))
    n_minifat = n_sectors(len(minifat_bytes))
    n_dir = (len(order) + 3) // 4

    # FAT size depends on the total, which depends on the FAT size.
    n_fat, n_difat = 1, 0
    for _ in range(64):
        total = n_data + n_ministream + n_minifat + n_dir + n_fat + n_difat
        need_fat = max(1, (total + 127) // 128)
        need_difat = 0 if need_fat <= 109 else (need_fat - 109 + 126) // 127
        if (need_fat, need_difat) == (n_fat, n_difat):
            break
        n_fat, n_difat = need_fat, need_difat
    total = n_data + n_ministream + n_minifat + n_dir + n_fat + n_difat

    fat = [FREESECT] * (n_fat * 128)
    body = bytearray(total * SECTOR)
    cursor = 0

    def place(blob):
        nonlocal cursor
        count = n_sectors(len(blob))
        sectors = list(range(cursor, cursor + count))
        body[cursor * SECTOR:cursor * SECTOR + len(blob)] = blob
        cursor += count
        _chain(fat, sectors)
        return sectors[0] if sectors else ENDOFCHAIN

    for node in big:
        node._start = place(node.data)
    ministream_start = place(bytes(mini)) if mini else ENDOFCHAIN
    minifat_start = place(minifat_bytes) if minifat_bytes else ENDOFCHAIN

    dirdata = bytearray()
    for node in order:
        name = node.name.encode('utf-16-le') + b'\x00\x00'
        entry = bytearray(128)
        entry[0:len(name)] = name
        struct.pack_into('<H', entry, 64, len(name))
        entry[66] = node.kind
        entry[67] = 1                                   # black; a legal colouring
        struct.pack_into('<III', entry, 68, node.left, node.right, node.child)
        entry[80:96] = node.clsid
        struct.pack_into('<I', entry, 96, 0)            # state bits
        struct.pack_into('<Q', entry, 100, node.create_time)
        struct.pack_into('<Q', entry, 108, node.modify_time)
        if node.kind == ROOT:
            struct.pack_into('<I', entry, 116, ministream_start)
            struct.pack_into('<Q', entry, 120, len(mini))
        elif node.kind == STREAM and node.data:
            start = node._start if node in big else mini_start[node.sid]
            struct.pack_into('<I', entry, 116, start)
            struct.pack_into('<Q', entry, 120, len(node.data))
        else:
            struct.pack_into('<I', entry, 116, ENDOFCHAIN)
            struct.pack_into('<Q', entry, 120, 0)
        dirdata += entry
    dir_start = place(bytes(dirdata))

    fat_sectors = list(range(cursor, cursor + n_fat))
    cursor += n_fat
    difat_sectors = list(range(cursor, cursor + n_difat))
    cursor += n_difat
    for s in fat_sectors:
        fat[s] = FATSECT
    for s in difat_sectors:
        fat[s] = DIFSECT
    assert cursor == total, (cursor, total)

    fatdata = b''.join(struct.pack('<I', v) for v in fat)
    body[fat_sectors[0] * SECTOR:fat_sectors[0] * SECTOR + len(fatdata)] = fatdata

    header = bytearray(SECTOR)
    header[0:8] = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
    struct.pack_into('<HHH', header, 24, 0x003E, 3, 0xFFFE)
    struct.pack_into('<HH', header, 30, 9, 6)
    struct.pack_into('<I', header, 40, 0)               # dir sector count: 0 in v3
    struct.pack_into('<I', header, 44, n_fat)
    struct.pack_into('<I', header, 48, dir_start)
    struct.pack_into('<I', header, 56, MINI_CUTOFF)
    struct.pack_into('<I', header, 60, minifat_start)
    struct.pack_into('<I', header, 64, n_minifat)
    struct.pack_into('<I', header, 68, difat_sectors[0] if difat_sectors else ENDOFCHAIN)
    struct.pack_into('<I', header, 72, n_difat)
    for i in range(109):
        struct.pack_into('<I', header, 76 + 4 * i,
                         fat_sectors[i] if i < len(fat_sectors) else FREESECT)
    for i, sector in enumerate(difat_sectors):
        entries = fat_sectors[109 + i * 127:109 + (i + 1) * 127]
        blob = b''.join(struct.pack('<I', v) for v in entries)
        blob += b'\xff\xff\xff\xff' * (127 - len(entries))
        nxt = difat_sectors[i + 1] if i + 1 < len(difat_sectors) else ENDOFCHAIN
        body[sector * SECTOR:(sector + 1) * SECTOR] = blob + struct.pack('<I', nxt)

    blob = bytes(header) + bytes(body)
    if path is None:
        return blob
    with open(path, 'wb') as fh:
        fh.write(blob)
    return blob
