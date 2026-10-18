"""Build a genuine Excel VBA project (vbaProject.bin) from a .bas source file,
and inject it into an .xlsx to produce a macro-enabled .xlsm.

Why this exists: LibreOffice cannot author Excel VBA projects — it runs VBA
macros but never exports them as vbaProject.bin. So this module constructs
the VBA project by hand, following MS-OVBA (dir stream + module streams +
PROJECT stream) and MS-CFB (the OLE compound container), using the record
layout validated against oletools' battle-tested parser.

Validation gate (see TEST-REPORT.md): the produced .xlsm must
  1. parse with oletools.olevba and yield the exact VBA source, and
  2. open in LibreOffice with the macro present and runnable.

100% local. No network. by Muhammad Bilal.
"""
import math
import struct
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple

# ---------------------------------------------------------------------------
# MS-OVBA compression (2.4.1) — mirror of oletools' decompress_stream
# ---------------------------------------------------------------------------


def _copy_token_params(difference: int):
    """bit_count / max length the decoder will compute at this chunk position."""
    bit_count = max(math.ceil(math.log2(difference)), 4) if difference > 1 else 4
    max_len = (0xFFFF >> bit_count) + 3
    return bit_count, max_len


def _encode_chunk(data: bytes) -> bytes:
    """Encode one chunk (<=4096 input bytes) as a compressed chunk payload
    (flag bytes + tokens, WITHOUT the 2-byte chunk header)."""
    out = bytearray()
    n = len(data)
    # index 3-byte sequences -> positions, for longest-match search
    index: Dict[bytes, List[int]] = {}
    for p in range(max(0, n - 2)):
        index.setdefault(data[p:p + 3], []).append(p)

    i = 0
    while i < n:
        flag_pos = len(out)
        out.append(0)  # placeholder flag byte
        flag = 0
        for bit in range(8):
            if i >= n:
                break
            difference = i  # bytes emitted in this chunk so far
            best_len, best_off = 0, 0
            if difference >= 1 and i + 3 <= n:
                bit_count, max_len = _copy_token_params(difference)
                max_len = min(max_len, n - i)
                for cand in reversed(index.get(data[i:i + 3], [])):
                    if cand >= i:
                        continue
                    off = i - cand
                    if off > difference:
                        continue
                    # A copy token (off, len) is valid iff for every l:
                    # data[i+l] == data[i+l-off] (decoder copies byte-by-byte
                    # from its own output, so overlapping/periodic matches work).
                    l = 0
                    while l < max_len and data[i + l] == data[i + l - off]:
                        l += 1
                    if l > best_len:
                        best_len, best_off = l, off
                        if l >= max_len:
                            break
            if best_len >= 3:
                bit_count, _ = _copy_token_params(difference)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flag |= (1 << bit)
                i += best_len
            else:
                out.append(data[i])
                i += 1
        out[flag_pos] = flag
    return bytes(out)


def compress_stream(data: bytes) -> bytes:
    """Compress a whole stream per MS-OVBA 2.4.1. Returns container bytes
    (0x01 signature + chunk headers + payloads)."""
    out = bytearray()
    out.append(0x01)
    pos = 0
    n = len(data)
    while pos < n:
        # try a full 4096-byte chunk; shrink until the payload fits (<=4096)
        size = min(4096, n - pos)
        while True:
            payload = _encode_chunk(data[pos:pos + size])
            if len(payload) <= 4096 or size <= 1:
                break
            size //= 2
        total = len(payload) + 2
        header = 0xB000 | (total - 3)  # 0b1011 signature + compressed flag + size-3
        out += struct.pack("<H", header)
        out += payload
        pos += size
    return bytes(out)


# ---------------------------------------------------------------------------
# MS-OVBA dir stream (decompressed form)
# ---------------------------------------------------------------------------

def _rec_u16(rec_id: int, value: int) -> bytes:
    return struct.pack("<HI", rec_id, 2) + struct.pack("<H", value)


def _rec_u32(rec_id: int, value: int) -> bytes:
    return struct.pack("<HI", rec_id, 4) + struct.pack("<I", value)


def _rec_bytes(rec_id: int, payload: bytes) -> bytes:
    return struct.pack("<HI", rec_id, len(payload)) + payload


def _str_rec(rec_id: int, text: str, reserved: int) -> bytes:
    """Length-prefixed MBCS string + reserved(2B) + length-prefixed UTF-16LE string."""
    mbcs = text.encode("cp1252")
    uni = text.encode("utf-16-le")
    return (struct.pack("<HI", rec_id, len(mbcs)) + mbcs +
            struct.pack("<H", reserved) +
            struct.pack("<I", len(uni)) + uni)


def _module_record(module_name: str, module_type: int) -> bytes:
    """One MODULE record. module_type: 0x0021 procedural, 0x0022 document."""
    out = bytearray()
    name_b = module_name.encode("cp1252")
    name_u = module_name.encode("utf-16-le")
    out += struct.pack("<HI", 0x0019, len(name_b)) + name_b            # MODULENAME
    out += struct.pack("<HI", 0x0047, len(name_u)) + name_u            # MODULENAMEUNICODE
    out += struct.pack("<HI", 0x001A, len(name_b)) + name_b            # MODULESTREAMNAME
    out += struct.pack("<H", 0x0032)                                   # reserved
    out += struct.pack("<I", len(name_u)) + name_u                     # stream name unicode
    out += _str_rec(0x001C, "", 0x0048)                                # MODULEDOCSTRING (empty)
    out += _rec_u32(0x0031, 0)                                          # MODULEOFFSET = 0
    out += _rec_u32(0x001E, 0)                                          # MODULEHELPCONTEXT
    out += struct.pack("<HIH", 0x002C, 2, 0)                           # MODULECOOKIE
    out += struct.pack("<HI", module_type, 0)                          # MODULETYPE
    out += struct.pack("<HI", 0x002B, 0)                               # MODULETERMINATOR
    return bytes(out)


def build_dir_stream(module_name: str, project_name: str = "VBAProject",
                     doc_modules: Tuple[str, ...] = ("ThisWorkbook", "Sheet1")) -> bytes:
    out = bytearray()
    # Header records
    out += _rec_u32(0x0001, 0x00000001)          # PROJECTSYSKIND: 32-bit Windows
    out += _rec_u32(0x0002, 0x00000409)          # PROJECTLCID: en-US
    out += _rec_u32(0x0014, 0x00000409)          # PROJECTLCIDINVOKE
    out += _rec_u16(0x0003, 0x04E4)              # PROJECTCODEPAGE: 1252
    out += _rec_bytes(0x0004, project_name.encode("cp1252"))  # PROJECTNAME
    out += _str_rec(0x0005, "", 0x0040)          # PROJECTDOCSTRING (empty)
    out += _str_rec(0x0006, "", 0x003D)          # PROJECTHELPFILEPATH (empty)
    out += _rec_u32(0x0007, 0)                    # PROJECTHELPCONTEXT
    out += _rec_u32(0x0008, 0)                    # PROJECTLIBFLAGS
    out += struct.pack("<HIIH", 0x0009, 4, 6, 0)  # PROJECTVERSION: reserved=4, major=6, minor=0
    out += _str_rec(0x000C, "", 0x003C)          # PROJECTCONSTANTS (empty)
    # Standard type-library references (REFERENCEREGISTERED 0x000D)
    libids = [
        r"*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation",
        r"*\\G{000204EF-0000-0000-C000-000000000046}#4.0#0#C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL#Visual Basic For Applications",
        r"*\\G{00020813-0000-0000-C000-000000000046}#1.9#0#C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE#Microsoft Excel 16.0 Object Library",
        r"*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library",
    ]
    for libid in libids:
        lb = libid.encode("cp1252")
        out += struct.pack("<HI", 0x000D, 4 + len(lb) + 4 + 2)
        out += struct.pack("<I", len(lb)) + lb
        out += struct.pack("<IH", 0, 0)
    # PROJECTMODULES terminator of the reference array
    all_modules = [(m, 0x0022) for m in doc_modules] + [(module_name, 0x0021)]
    out += struct.pack("<H", 0x000F)             # PROJECTMODULES Id
    out += struct.pack("<I", 2)                  # Size = 2
    out += struct.pack("<H", len(all_modules))   # ModulesCount
    out += struct.pack("<HIH", 0x0013, 2, 0)     # PROJECTCOOKIE record
    # ---- MODULE records: document modules first (as Excel writes them),
    # then the procedural (standard) module ----
    for m_name, m_type in all_modules:
        out += _module_record(m_name, m_type)
    return bytes(out)


def build_project_stream(module_name: str,
                         doc_modules: Tuple[str, ...] = ("ThisWorkbook", "Sheet1")) -> bytes:
    lines = [
        'ID="{3F1B5C2A-7E4A-4B2C-9D3A-5F6A7B8C9D0E}"',
    ]
    for dm in doc_modules:
        lines.append(f"Document={dm}/&H00000000")
    lines += [
        f"Module={module_name}",
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "&H00000002={00020813-0000-0000-C000-000000000046};Excel;&H00000000",
        "",
    ]
    return ("\r\n".join(lines)).encode("cp1252")


# ---------------------------------------------------------------------------
# MS-CFB (OLE compound file) writer — minimal, sector size 512
# ---------------------------------------------------------------------------

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF
SECTOR = 512
MINI_SECTOR = 64
MINI_CUTOFF = 4096


def _dir_entry(name: str, obj_type: int, color: int, left: int, right: int,
               child: int, start_sector: int, size: int) -> bytes:
    name_u16 = name.encode("utf-16-le") + b"\x00\x00"
    name_field = name_u16.ljust(64, b"\x00")
    return (name_field +
            struct.pack("<H", len(name_u16)) +
            struct.pack("<B", obj_type) +
            struct.pack("<B", color) +
            struct.pack("<III", left, right, child) +
            b"\x00" * 16 +                       # CLSID
            struct.pack("<I", 0) +               # state bits
            struct.pack("<Q", 0) +               # creation time
            struct.pack("<Q", 0) +               # modified time
            struct.pack("<I", start_sector) +
            struct.pack("<I", size) + struct.pack("<I", 0))  # size (v3: 4B + unused)


def _sort_key(name: str):
    # CFB directory comparison: (name length, uppercase name)
    return (len(name), name.upper())


def _build_tree(ids_names: List[Tuple[int, str]]) -> Tuple[int, Dict[int, Tuple[int, int]]]:
    """Build a balanced binary tree over sorted children.
    Returns (root_entry_id, {entry_id: (left_id, right_id)})."""
    ordered = sorted(ids_names, key=lambda t: _sort_key(t[1]))
    links: Dict[int, Tuple[int, int]] = {}

    def build(items):
        if not items:
            return NOSTREAM
        mid = len(items) // 2
        eid, _ = items[mid]
        left = build(items[:mid])
        right = build(items[mid + 1:])
        links[eid] = (left, right)
        return eid

    root = build(ordered)
    return root, links


def build_cfb(streams: List[Tuple[str, bytes]]) -> bytes:
    """Build a CFB file. streams: list of (path, data) where path uses '/'
    separators, e.g. '_VBA_PROJECT_CUR/VBA/dir'. Storages are created implicitly.
    Returns the complete file bytes."""
    # --- collect storages and assign entry ids ---
    # entries: 0 = root, then storages, then streams (stable order)
    paths = {}
    for path, _ in streams:
        parts = path.split("/")
        for i in range(1, len(parts)):
            paths["/".join(parts[:i])] = "storage"
    for path, _ in streams:
        paths[path] = "stream"

    entry_ids = {"": 0}
    names = {0: ""}  # entry id -> short name
    kinds = {0: "root"}
    next_id = 1
    # storages first (sorted for determinism), then streams
    for p in sorted(paths):
        if paths[p] == "storage":
            entry_ids[p] = next_id
            names[next_id] = p.split("/")[-1]
            kinds[next_id] = "storage"
            next_id += 1
    for p in sorted(paths):
        if paths[p] == "stream":
            entry_ids[p] = next_id
            names[next_id] = p.split("/")[-1]
            kinds[next_id] = "stream"
            next_id += 1
    n_entries = next_id

    # --- children lists per storage ---
    children: Dict[int, List[int]] = {i: [] for i in range(n_entries)}
    for p, eid in entry_ids.items():
        if p == "":
            continue
        parent_path = "/".join(p.split("/")[:-1])
        parent_id = entry_ids.get(parent_path, 0)
        children[parent_id].append(eid)

    # --- build sibling trees + colors (tree root black, others red) ---
    links: Dict[int, Tuple[int, int]] = {}
    colors: Dict[int, int] = {0: 1}
    for parent_id, kids in children.items():
        if not kids:
            continue
        pairs = [(k, names[k]) for k in kids]
        root_id, tree_links = _build_tree(pairs)
        links.update(tree_links)
        colors[root_id] = 1
        for k in kids:
            colors.setdefault(k, 0)
        children[parent_id] = [root_id]  # store tree root as the child pointer

    # --- split streams into mini (<4096) vs regular ---
    mini_streams = []   # (entry_id, data)
    regular = []        # (entry_id, data)
    for path, data in streams:
        eid = entry_ids[path]
        if len(data) < MINI_CUTOFF:
            mini_streams.append((eid, data))
        else:
            regular.append((eid, data))

    # --- mini stream layout ---
    mini_alloc: Dict[int, Tuple[int, int]] = {}  # entry -> (mini_start, mini_len)
    mini_data = bytearray()
    for eid, data in mini_streams:
        start = len(mini_data) // MINI_SECTOR
        padded = data + b"\x00" * ((MINI_SECTOR - len(data) % MINI_SECTOR) % MINI_SECTOR)
        mini_data += padded
        mini_alloc[eid] = (start, len(data))
    n_mini_sectors = len(mini_data) // MINI_SECTOR

    # --- sector layout ---
    # sector 0 = FAT; then directory sectors; then regular stream sectors;
    # then MiniFAT sector; then ministream sectors.
    n_dir_sectors = (n_entries * 128 + SECTOR - 1) // SECTOR
    fat_sector = 0
    dir_start = 1
    cursor = 1 + n_dir_sectors
    regular_alloc: Dict[int, Tuple[int, int]] = {}
    for eid, data in regular:
        nsec = (len(data) + SECTOR - 1) // SECTOR
        regular_alloc[eid] = (cursor, nsec)
        cursor += nsec
    minifat_sector = cursor
    cursor += 1
    ministream_start = cursor
    n_ministream_sectors = (len(mini_data) + SECTOR - 1) // SECTOR if mini_data else 0
    total_sectors = cursor + n_ministream_sectors

    # --- FAT ---
    fat = [FREESECT] * total_sectors
    fat[fat_sector] = FATSECT
    for s in range(dir_start, dir_start + n_dir_sectors - 1):
        fat[s] = s + 1
    if n_dir_sectors:
        fat[dir_start + n_dir_sectors - 1] = ENDOFCHAIN
    for eid, (start, nsec) in regular_alloc.items():
        for k in range(nsec - 1):
            fat[start + k] = start + k + 1
        fat[start + nsec - 1] = ENDOFCHAIN
    fat[minifat_sector] = ENDOFCHAIN
    for s in range(ministream_start, ministream_start + n_ministream_sectors - 1):
        fat[s] = s + 1
    if n_ministream_sectors:
        fat[ministream_start + n_ministream_sectors - 1] = ENDOFCHAIN

    # --- MiniFAT ---
    minifat = [FREESECT] * n_mini_sectors
    offset = 0
    for eid, data in mini_streams:
        nsec = (len(data) + MINI_SECTOR - 1) // MINI_SECTOR
        for k in range(nsec - 1):
            minifat[offset + k] = offset + k + 1
        if nsec:
            minifat[offset + nsec - 1] = ENDOFCHAIN
        offset += nsec

    # --- directory entries ---
    entries = []
    for eid in range(n_entries):
        name = names[eid]
        kind = kinds[eid]
        if kind == "root":
            obj_type, child = 5, children[0][0] if children[0] else NOSTREAM
            start, size = ministream_start if mini_data else ENDOFCHAIN, len(mini_data)
            left = right = NOSTREAM
            color = 1
        elif kind == "storage":
            obj_type = 1
            child = children[eid][0] if children[eid] else NOSTREAM
            start, size = ENDOFCHAIN, 0
            left, right = links.get(eid, (NOSTREAM, NOSTREAM))
            color = colors.get(eid, 0)
        else:
            obj_type = 2
            child = NOSTREAM
            if eid in mini_alloc:
                start, size = mini_alloc[eid]
            else:
                start, size = regular_alloc[eid][0], len(dict(streams)[_path_of(entry_ids, eid)])
            left, right = links.get(eid, (NOSTREAM, NOSTREAM))
            color = colors.get(eid, 0)
        entries.append(_dir_entry(name, obj_type, color, left, right, child, start, size))

    # --- assemble file ---
    out = bytearray()
    header = bytearray(SECTOR)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<H", header, 24, 0x003E)   # minor version
    struct.pack_into("<H", header, 26, 0x0003)   # major version (3)
    struct.pack_into("<H", header, 28, 0xFFFE)   # byte order
    struct.pack_into("<H", header, 30, 9)        # sector shift (512)
    struct.pack_into("<H", header, 32, 6)        # mini sector shift (64)
    struct.pack_into("<I", header, 40, 0)        # num dir sectors (v3: 0)
    struct.pack_into("<I", header, 44, 1)        # num FAT sectors
    struct.pack_into("<I", header, 48, dir_start)  # first dir sector
    struct.pack_into("<I", header, 52, 0)        # transaction signature
    struct.pack_into("<I", header, 56, MINI_CUTOFF)
    struct.pack_into("<I", header, 60, minifat_sector)
    struct.pack_into("<I", header, 64, 1)        # num MiniFAT sectors
    struct.pack_into("<I", header, 68, ENDOFCHAIN)  # first DIFAT sector
    struct.pack_into("<I", header, 72, 0)        # num DIFAT sectors
    for i in range(109):
        struct.pack_into("<I", header, 76 + 4 * i, fat_sector if i == 0 else FREESECT)
    out += header

    # FAT sector
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    out += fat_bytes.ljust(SECTOR, b"\xFF")
    # directory sectors
    dir_bytes = b"".join(entries)
    out += dir_bytes.ljust(n_dir_sectors * SECTOR, b"\x00")
    # regular stream sectors
    for eid, data in regular:
        padded = data + b"\x00" * ((SECTOR - len(data) % SECTOR) % SECTOR)
        out += padded
    # MiniFAT sector
    minifat_bytes = b"".join(struct.pack("<I", v) for v in minifat)
    out += minifat_bytes.ljust(SECTOR, b"\xFF")
    # ministream sectors
    out += bytes(mini_data).ljust(n_ministream_sectors * SECTOR, b"\x00")
    return bytes(out)


def _path_of(entry_ids: Dict[str, int], eid: int) -> str:
    for p, i in entry_ids.items():
        if i == eid:
            return p
    raise KeyError(eid)


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def build_vba_project_bin(bas_path: str, module_name: str = "TaskDashboard") -> bytes:
    """Build vbaProject.bin from a .bas file (Attribute lines stripped,
    line endings normalized to CRLF)."""
    with open(bas_path, encoding="utf-8") as f:
        src = f.read()
    lines = [ln for ln in src.splitlines() if not ln.startswith("Attribute ")]
    code = ("\r\n".join(lines) + "\r\n").encode("cp1252")
    dir_raw = build_dir_stream(module_name)
    project_raw = build_project_stream(module_name)
    streams = [
        ("_VBA_PROJECT_CUR/PROJECT", project_raw),
        # MS-OVBA 2.2.1: the VBA storage MUST contain a _VBA_PROJECT stream.
        # It holds a regenerable compiled-project cache; hosts rebuild it from
        # source, so an empty stream is the safe "no cache" representation.
        ("_VBA_PROJECT_CUR/VBA/_VBA_PROJECT", b""),
        ("_VBA_PROJECT_CUR/VBA/dir", compress_stream(dir_raw)),
        (f"_VBA_PROJECT_CUR/VBA/{module_name}", compress_stream(code)),
    ]
    # Empty streams for the document modules (ThisWorkbook / Sheet1), as a
    # fresh Excel file has no code behind them.
    for dm in ("ThisWorkbook", "Sheet1"):
        streams.append((f"_VBA_PROJECT_CUR/VBA/{dm}", compress_stream(b"")))
    return build_cfb(streams)


def inject_into_xlsm(xlsx_path: str, vba_bin: bytes, out_path: str) -> None:
    """Take an .xlsx (zip) and produce an .xlsm with the VBA project wired in."""
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        items = {n: zin.read(n) for n in zin.namelist()}

    # [Content_Types].xml: register .bin as vbaProject AND mark the workbook
    # part as macro-enabled, otherwise hosts will not load the VBA project.
    ct = items["[Content_Types].xml"].decode("utf-8")
    if "vbaProject" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
    ct = ct.replace(
        'PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"',
        'PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"')
    items["[Content_Types].xml"] = ct.encode("utf-8")

    # xl/_rels/workbook.xml.rels: add vbaProject relationship with a fresh rId
    rels = items["xl/_rels/workbook.xml.rels"].decode("utf-8")
    if "vbaProject" not in rels:
        import re
        ids = [int(m.group(1)) for m in re.finditer(r'Id="rId(\d+)"', rels)]
        new_id = f"rId{(max(ids) + 1) if ids else 1}"
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="{new_id}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
    items["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    # Each worksheet's sheetPr gets a VBA codeName (Sheet1, Sheet2, ...) in
    # WORKBOOK order, so the document modules declared in the VBA project
    # (ThisWorkbook/SheetN) associate with the sheets, exactly as Excel does.
    import re
    NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    wb_root = ET.fromstring(items["xl/workbook.xml"])
    rels_root = ET.fromstring(items["xl/_rels/workbook.xml.rels"])
    rel_map = {rel.get("Id"): rel.get("Target")
               for rel in rels_root.findall("rel:Relationship", {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})}
    # (rels use the package namespace; parse leniently if needed)
    if not rel_map:
        rel_map = {el.get("Id"): el.get("Target") for el in rels_root.iter() if el.get("Id")}
    for idx, sheet_el in enumerate(wb_root.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"), start=1):
        rid = sheet_el.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rel_map.get(rid, "")
        part = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
        if part not in items:
            continue
        xml = items[part].decode("utf-8")
        if "codeName" not in xml:
            xml = re.sub(r"<sheetPr(?=[ />])", f'<sheetPr codeName="Sheet{idx}"', xml, count=1)
            if f'codeName="Sheet{idx}"' not in xml:
                xml = re.sub(r"(<worksheet[^>]*>)", rf"\1<sheetPr codeName=\"Sheet{idx}\"/>", xml, count=1)
            items[part] = xml.encode("utf-8")

    items["xl/vbaProject.bin"] = vba_bin
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)


if __name__ == "__main__":
    import sys
    bas = sys.argv[1]
    out = sys.argv[2]
    with open(out, "wb") as f:
        f.write(build_vba_project_bin(bas))
    print(f"Wrote {out}")
