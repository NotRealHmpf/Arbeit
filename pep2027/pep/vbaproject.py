"""Erzeugt ein VBA-Projekt (vbaProject.bin) aus Quelltext – ohne Excel.

Umsetzung nach [MS-OVBA] (VBA File Format) und [MS-CFB] (Compound File Binary):
- Modulquellen werden mit dem MS-OVBA-Verfahren komprimiert (2.4.1),
- der _VBA_PROJECT-Stream enthaelt keinen P-Code-Cache (Version 0xFFFF, 2.3.4.1), Excel kompiliert beim
  Oeffnen aus dem Quelltext,
- PROJECT/PROJECTwm/dir-Streams wie in 2.3.1, 2.3.3, 2.3.4.2 beschrieben,
- alles in einer CFB-Datei (Version 3, 512-Byte-Sektoren, Mini-Stream fuer kleine Streams).
"""
import random
import struct
import uuid

CODEPAGE = 1252
ENC = "cp1252"


# =============================================================================== Kompression (2.4.1)
def _copy_token_help(difference):
    bit_count = 4
    while (1 << bit_count) < difference:
        bit_count += 1
    length_mask = 0xFFFF >> bit_count
    max_length = length_mask + 3
    return bit_count, length_mask, max_length


def compress(data: bytes) -> bytes:
    """MS-OVBA-Kompression (2.4.1.3.6) mit Hash-Ketten fuer die Suche nach Wiederholungen."""
    out = bytearray(b"\x01")
    pos = 0
    n = len(data)
    while pos < n:
        chunk_start = pos
        chunk_end = min(pos + 4096, n)
        chunk = bytearray()
        heads = {}

        def remember(p):
            if p + 3 <= chunk_end:
                heads.setdefault(data[p:p + 3], []).append(p)

        while pos < chunk_end:
            flag_pos = len(chunk)
            chunk.append(0)
            flags = 0
            for bit in range(8):
                if pos >= chunk_end:
                    break
                best_len, best_off = 0, 0
                difference = pos - chunk_start
                if difference > 0 and pos + 3 <= chunk_end:
                    bit_count, _, max_len = _copy_token_help(difference)
                    max_len = min(max_len, chunk_end - pos)
                    cands = heads.get(data[pos:pos + 3], [])
                    for cand in reversed(cands[-64:]):
                        ln = 3
                        while ln < max_len and data[cand + ln] == data[pos + ln]:
                            ln += 1
                        if ln > best_len:
                            best_len, best_off = ln, pos - cand
                            if ln == max_len:
                                break
                if best_len >= 3:
                    bit_count, _, _ = _copy_token_help(difference)
                    token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                    chunk += struct.pack("<H", token)
                    flags |= 1 << bit
                    for p in range(pos, pos + best_len):
                        remember(p)
                    pos += best_len
                else:
                    chunk.append(data[pos])
                    remember(pos)
                    pos += 1
            chunk[flag_pos] = flags
        raw_len = chunk_end - chunk_start
        if len(chunk) > 4096 and raw_len == 4096:
            header = (4096 + 2 - 3) | (0b011 << 12)            # unkomprimiert
            out += struct.pack("<H", header) + data[chunk_start:chunk_end]
        else:
            if len(chunk) + 2 > 4098:
                raise ValueError("Chunk zu gross")
            header = (len(chunk) + 2 - 3) | (0b011 << 12) | 0x8000
            out += struct.pack("<H", header) + chunk
    return bytes(out)


def decompress(data: bytes) -> bytes:
    if data[0] != 1:
        raise ValueError("Signatur fehlt")
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        flag = header >> 15
        if (header >> 12) & 0b111 != 0b011:
            raise ValueError("Chunk-Signatur falsch")
        chunk_end = pos + size
        pos += 2
        dstart = len(out)
        if not flag:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if flags & (1 << bit):
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    bit_count, length_mask, _ = _copy_token_help(len(out) - dstart)
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
                else:
                    out.append(data[pos])
                    pos += 1
    return bytes(out)


# =============================================================================== dir-Stream (2.3.4.2)
def _rec(rid, payload: bytes):
    return struct.pack("<HI", rid, len(payload)) + payload


def _mbcs(s):
    return s.encode(ENC)


def _utf16(s):
    return s.encode("utf-16-le")


REFERENCES = [
    ("stdole", "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation"),
    ("Office", "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\Microsoft Shared\\"
               "OFFICE16\\MSO.DLL#Microsoft Office 16.0 Object Library"),
]


def build_dir(modules, project_name="VBAProject"):
    b = bytearray()
    b += _rec(0x0001, struct.pack("<I", 0x00000001))                    # SYSKIND Win32
    b += _rec(0x0002, struct.pack("<I", 0x0409))                        # LCID
    b += _rec(0x0014, struct.pack("<I", 0x0409))                        # LCIDINVOKE
    b += _rec(0x0003, struct.pack("<H", CODEPAGE))                      # CODEPAGE
    b += _rec(0x0004, _mbcs(project_name))                              # NAME
    b += _rec(0x0005, b"") + _rec(0x0040, b"")                          # DOCSTRING + Unicode
    b += _rec(0x0006, b"") + _rec(0x003D, b"")                          # HELPFILEPATH 1 + 2
    b += _rec(0x0007, struct.pack("<I", 0))                             # HELPCONTEXT
    b += _rec(0x0008, struct.pack("<I", 0))                             # LIBFLAGS
    b += struct.pack("<HIIH", 0x0009, 0x00000004, 1, 0)                 # VERSION (Reserved=4, Major, Minor)
    b += _rec(0x000C, b"") + _rec(0x003C, b"")                          # CONSTANTS + Unicode
    for name, libid in REFERENCES:
        b += _rec(0x0016, _mbcs(name)) + _rec(0x003E, _utf16(name))      # REFERENCENAME
        lib = _mbcs(libid)
        payload = struct.pack("<I", len(lib)) + lib + struct.pack("<IH", 0, 0)
        b += _rec(0x000D, payload)                                      # REFERENCEREGISTERED
    b += _rec(0x000F, struct.pack("<H", len(modules)))                  # MODULES count
    b += _rec(0x0013, struct.pack("<H", 0xFFFF))                        # PROJECTCOOKIE
    for m in modules:
        b += _rec(0x0019, _mbcs(m["name"]))                             # MODULENAME
        b += _rec(0x0047, _utf16(m["name"]))                            # MODULENAMEUNICODE
        b += _rec(0x001A, _mbcs(m["name"])) + _rec(0x0032, _utf16(m["name"]))   # STREAMNAME
        b += _rec(0x001C, b"") + _rec(0x0048, b"")                      # DOCSTRING
        b += _rec(0x0031, struct.pack("<I", 0))                         # OFFSET (kein P-Code)
        b += _rec(0x001E, struct.pack("<I", 0))                         # HELPCONTEXT
        b += _rec(0x002C, struct.pack("<H", 0xFFFF))                    # COOKIE
        b += struct.pack("<HI", 0x0022 if m["document"] else 0x0021, 0)  # TYPE
        b += struct.pack("<HI", 0x002B, 0)                              # Terminator
    b += struct.pack("<HI", 0x0010, 0)                                  # dir Terminator
    return bytes(b)


# =============================================================================== PROJECT-Stream (2.3.1)
def encrypt(data: bytes, project_id: str, seed=None) -> str:
    """Datenverschluesselung 2.4.3.2 (fuer CMG, DPB, GC)."""
    rnd = random.Random(seed)
    seed_b = rnd.randrange(256) if seed is None else seed & 0xFF
    version = 2
    proj_key = sum(project_id.encode(ENC)) & 0xFF
    out = bytearray([seed_b, seed_b ^ version, seed_b ^ proj_key])
    unenc1, enc1, enc2 = proj_key, out[2], out[1]
    ignored_len = (seed_b & 6) // 2
    for _ in range(ignored_len):
        temp = rnd.randrange(256)
        b = temp ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, temp
    for byte in struct.pack("<I", len(data)) + data:
        b = byte ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, byte
    return out.hex().upper()


def decrypt(hexstr: str) -> bytes:
    d = bytes.fromhex(hexstr)
    seed_b, version_enc, proj_key_enc = d[0], d[1], d[2]
    if seed_b ^ version_enc != 2:
        raise ValueError("Version")
    proj_key = seed_b ^ proj_key_enc
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc
    pos = 3
    for _ in range((seed_b & 6) // 2):
        b = d[pos]; pos += 1
        temp = b ^ ((enc2 + unenc1) & 0xFF)
        enc2, enc1, unenc1 = enc1, b, temp
    plain = bytearray()
    for b in d[pos:]:
        byte = b ^ ((enc2 + unenc1) & 0xFF)
        plain.append(byte)
        enc2, enc1, unenc1 = enc1, b, byte
    length = struct.unpack_from("<I", plain, 0)[0]
    return bytes(plain[4:4 + length])


def build_project(modules, project_id):
    lines = [f'ID="{project_id}"']
    for m in modules:
        if m["document"]:
            lines.append(f"Document={m['name']}/&H00000000")
    for m in modules:
        if not m["document"]:
            lines.append(f"Module={m['name']}")
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              f'CMG="{encrypt(struct.pack("<I", 0), project_id, 0x35)}"',
              f'DPB="{encrypt(bytes([0]), project_id, 0x6B)}"',
              f'GC="{encrypt(bytes([0xFF]), project_id, 0x9A)}"',
              "", "[Host Extender Info]", "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
              "", "[Workspace]"]
    for m in modules:
        lines.append(f"{m['name']}=0, 0, 0, 0, C")
    return ("\r\n".join(lines) + "\r\n").encode(ENC)


def build_projectwm(modules):
    b = bytearray()
    for m in modules:
        b += _mbcs(m["name"]) + b"\x00" + _utf16(m["name"]) + b"\x00\x00"
    b += b"\x00\x00"
    return bytes(b)


# =============================================================================== CFB-Writer
SECT = 512
MINI = 64
CUTOFF = 4096
FREESECT, ENDOFCHAIN, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


class Node:
    def __init__(self, name, data=None, children=None, root=False):
        self.name, self.data, self.children, self.root = name, data, children, root
        self.sid = None
        self.left = self.right = self.child = NOSTREAM
        self.color = 1
        self.start, self.size = ENDOFCHAIN, 0

    @property
    def is_storage(self):
        return self.children is not None


def _cmp_key(name):
    return (len(name), name.upper())


def _build_tree(nodes, depth=0, maxdepth=None):
    """Ausgeglichener Binaerbaum; unterste unvollstaendige Ebene rot, sonst schwarz (gueltiger RB-Baum)."""
    if not nodes:
        return NOSTREAM
    mid = len(nodes) // 2
    n = nodes[mid]
    n.left = _build_tree(nodes[:mid], depth + 1, maxdepth)
    n.right = _build_tree(nodes[mid + 1:], depth + 1, maxdepth)
    n._depth = depth
    return n.sid


def _assign_colors(nodes_sorted):
    import math
    count = len(nodes_sorted)
    if count == 0:
        return
    full = int(math.floor(math.log2(count + 1)))          # vollstaendige Ebenen
    for n in nodes_sorted:
        n.color = 0 if (n._depth >= full and (2 ** full - 1) != count) else 1


def write_cfb(root: Node) -> bytes:
    # 1. Knoten durchnummerieren (Root = 0)
    allnodes = []

    def walk(n):
        n.sid = len(allnodes)
        allnodes.append(n)
        if n.is_storage:
            for ch in n.children:
                walk(ch)
    walk(root)
    # 2. Baeume je Storage
    for n in allnodes:
        if n.is_storage and n.children:
            kids = sorted(n.children, key=lambda x: _cmp_key(x.name))
            n.child = _build_tree(kids)
            _assign_colors(kids)
    root.color = 1
    # 3. Streams verteilen: klein -> Mini-Stream, gross -> normale Sektoren
    mini = bytearray()
    minifat = []
    big = []
    for n in allnodes:
        if n.is_storage:
            continue
        n.size = len(n.data)
        if n.size == 0:
            n.start = ENDOFCHAIN
        elif n.size < CUTOFF:
            first = len(mini) // MINI
            nsec = (n.size + MINI - 1) // MINI
            n.start = first
            mini += n.data + b"\x00" * (nsec * MINI - n.size)
            for i in range(nsec):
                minifat.append(first + i + 1 if i < nsec - 1 else ENDOFCHAIN)
        else:
            big.append(n)
    # Sektorplanung: [FAT][Verzeichnis][MiniFAT][Mini-Stream][grosse Streams]
    dir_bytes_len = len(allnodes) * 128
    n_dir = (dir_bytes_len + SECT - 1) // SECT
    minifat_bytes = struct.pack(f"<{len(minifat)}I", *minifat) if minifat else b""
    n_minifat = (len(minifat_bytes) + SECT - 1) // SECT
    n_ministream = (len(mini) + SECT - 1) // SECT
    n_big = sum((len(b.data) + SECT - 1) // SECT for b in big)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_ministream + n_big
        if n_fat * (SECT // 4) >= total:
            break
        n_fat += 1
    if n_fat > 109:
        raise ValueError("DIFAT nicht unterstuetzt")
    fat = [FREESECT] * (n_fat * (SECT // 4))
    sec = 0
    fat_secs = list(range(sec, sec + n_fat))
    for s in fat_secs:
        fat[s] = FATSECT
    sec += n_fat

    def chain(start, count):
        for i in range(count):
            fat[start + i] = start + i + 1 if i < count - 1 else ENDOFCHAIN

    dir_start = sec
    chain(sec, n_dir)
    sec += n_dir
    minifat_start = sec if n_minifat else ENDOFCHAIN
    chain(sec, n_minifat)
    sec += n_minifat
    ministream_start = sec if n_ministream else ENDOFCHAIN
    chain(sec, n_ministream)
    sec += n_ministream
    for b in big:
        cnt = (len(b.data) + SECT - 1) // SECT
        b.start = sec
        chain(sec, cnt)
        sec += cnt
    root.start = ministream_start
    root.size = len(mini)
    # Verzeichnis schreiben
    dir_bytes = bytearray()
    for n in allnodes:
        name16 = n.name.encode("utf-16-le")
        if len(name16) > 62:
            raise ValueError("Name zu lang: " + n.name)
        e = bytearray(128)
        e[0:len(name16)] = name16
        struct.pack_into("<H", e, 64, len(name16) + 2)
        e[66] = 5 if n.root else (1 if n.is_storage else 2)
        e[67] = n.color
        struct.pack_into("<III", e, 68, n.left, n.right, n.child if n.is_storage else NOSTREAM)
        # CLSID, State, Zeiten = 0
        struct.pack_into("<I", e, 116, n.start if (n.root or not n.is_storage) else 0)
        struct.pack_into("<Q", e, 120, n.size if (n.root or not n.is_storage) else 0)
        dir_bytes += e
    dir_bytes += b"\x00" * (n_dir * SECT - len(dir_bytes))
    # leere Verzeichniseintraege muessen "frei" sein: Links/Child = NOSTREAM
    for i in range(len(allnodes), n_dir * (SECT // 128)):
        off = i * 128
        struct.pack_into("<III", dir_bytes, off + 68, NOSTREAM, NOSTREAM, NOSTREAM)
    # Header
    h = bytearray(SECT)
    h[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", h, 24, 0x003E, 0x0003, 0xFFFE, 9, 6)
    struct.pack_into("<I", h, 40, 0)                 # Anzahl Verzeichnissektoren (v3: 0)
    struct.pack_into("<I", h, 44, n_fat)
    struct.pack_into("<I", h, 48, dir_start)
    struct.pack_into("<I", h, 52, 0)
    struct.pack_into("<I", h, 56, CUTOFF)
    struct.pack_into("<I", h, 60, minifat_start)
    struct.pack_into("<I", h, 64, n_minifat)
    struct.pack_into("<I", h, 68, ENDOFCHAIN)        # erster DIFAT-Sektor
    struct.pack_into("<I", h, 72, 0)
    difat = fat_secs + [FREESECT] * (109 - len(fat_secs))
    struct.pack_into("<109I", h, 76, *difat)
    body = bytearray()
    body += struct.pack(f"<{len(fat)}I", *fat)
    body += dir_bytes
    if n_minifat:
        mf = minifat_bytes + struct.pack("<I", FREESECT) * ((n_minifat * SECT - len(minifat_bytes)) // 4)
        body += mf
    if n_ministream:
        body += mini + b"\x00" * (n_ministream * SECT - len(mini))
    for b in big:
        cnt = (len(b.data) + SECT - 1) // SECT
        body += b.data + b"\x00" * (cnt * SECT - len(b.data))
    return bytes(h + body)


# =============================================================================== Projekt zusammenbauen
DOC_ATTR_WB = ['Attribute VB_Name = "{name}"', 'Attribute VB_Base = "0{{00020819-0000-0000-C000-000000000046}}"',
               "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False",
               "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = True",
               "Attribute VB_TemplateDerived = False", "Attribute VB_Customizable = True"]
DOC_ATTR_WS = [DOC_ATTR_WB[0], 'Attribute VB_Base = "0{{00020820-0000-0000-C000-000000000046}}"'] + DOC_ATTR_WB[2:]


def module_source(name, body, kind):
    """kind: 'workbook', 'sheet' oder 'module'."""
    if kind == "module":
        head = [f'Attribute VB_Name = "{name}"']
    else:
        head = [a.format(name=name) for a in (DOC_ATTR_WB if kind == "workbook" else DOC_ATTR_WS)]
    text = "\r\n".join(head) + "\r\n" + body.replace("\r\n", "\n").replace("\n", "\r\n")
    if not text.endswith("\r\n"):
        text += "\r\n"
    return text.encode(ENC)


def build(modules_src, project_id=None):
    """modules_src: Liste (Name, Quelltext ohne Attribute, kind)."""
    project_id = project_id or "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, "pep-planung")).upper() + "}"
    modules = [{"name": n, "document": k != "module", "src": module_source(n, b, k)} for n, b, k in modules_src]
    vba_children = [Node("_VBA_PROJECT", bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])),
                    Node("dir", compress(build_dir(modules)))]
    for m in modules:
        vba_children.append(Node(m["name"], compress(m["src"])))
    root = Node("Root Entry", children=[
        Node("VBA", children=vba_children),
        Node("PROJECT", build_project(modules, project_id)),
        Node("PROJECTwm", build_projectwm(modules)),
    ], root=True)
    return write_cfb(root)
