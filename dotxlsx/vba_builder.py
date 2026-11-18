"""Pure-Python MS-OVBA vbaProject.bin builder for auto-playing .xlsm files."""

from __future__ import annotations

import struct
from typing import Sequence

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

THISWORKBOOK_CODE = (
    'Attribute VB_Name = "ThisWorkbook"\r\n'
    'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\n'
    'Attribute VB_GlobalNameSpace = False\r\n'
    'Attribute VB_Creatable = False\r\n'
    'Attribute VB_PredeclaredId = True\r\n'
    'Attribute VB_Exposed = True\r\n'
    'Attribute VB_TemplateDerived = False\r\n'
    'Attribute VB_Customizable = True\r\n'
    'Private Sub Workbook_Open()\r\n'
    '    StartDotXLSXAnim\r\n'
    'End Sub\r\n'
    'Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n'
    '    StopDotXLSXAnim\r\n'
    'End Sub\r\n'
)

MODULE1_CODE = (
    'Attribute VB_Name = "Module1"\r\n'
    'Option Explicit\r\n'
    'Private NextTickTime As Double\r\n'
    'Private IsRunning As Boolean\r\n'
    '\r\n'
    'Public Sub StartDotXLSXAnim()\r\n'
    '    On Error Resume Next\r\n'
    '    If ThisWorkbook.Sheets("Engine").Range("B4").Value <= 1 Then Exit Sub\r\n'
    '    IsRunning = True\r\n'
    '    ScheduleNextFrame\r\n'
    'End Sub\r\n'
    '\r\n'
    'Public Sub StopDotXLSXAnim()\r\n'
    '    On Error Resume Next\r\n'
    '    IsRunning = False\r\n'
    '    Application.OnTime NextTickTime, "AdvanceDotXLSXFrame", , False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Public Sub AdvanceDotXLSXFrame()\r\n'
    '    On Error GoTo ErrHandler\r\n'
    '    If Not IsRunning Then Exit Sub\r\n'
    '    If ActiveWorkbook Is Nothing Then Exit Sub\r\n'
    '    If ActiveWorkbook.Name <> ThisWorkbook.Name Then\r\n'
    '        ScheduleNextFrame\r\n'
    '        Exit Sub\r\n'
    '    End If\r\n'
    '    Dim wsEngine As Worksheet\r\n'
    '    Set wsEngine = ThisWorkbook.Sheets("Engine")\r\n'
    '    Dim totalFrames As Long\r\n'
    '    totalFrames = CLng(wsEngine.Range("B4").Value)\r\n'
    '    If totalFrames <= 1 Then Exit Sub\r\n'
    '    wsEngine.Range("B2").Value = (CLng(wsEngine.Range("B2").Value) + 1) Mod totalFrames\r\n'
    '    ScheduleNextFrame\r\n'
    '    Exit Sub\r\n'
    'ErrHandler:\r\n'
    '    IsRunning = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub ScheduleNextFrame()\r\n'
    '    Dim delayMs As Long\r\n'
    '    On Error Resume Next\r\n'
    '    delayMs = CLng(ThisWorkbook.Sheets("Engine").Range("B5").Value)\r\n'
    '    If delayMs < 30 Then delayMs = 100\r\n'
    '    NextTickTime = Now + (delayMs / 86400000#)\r\n'
    '    Application.OnTime NextTickTime, "AdvanceDotXLSXFrame"\r\n'
    'End Sub\r\n'
)


def _ovba_compress(data: bytes) -> bytes:
    """Compress bytes using valid MS-OVBA literal-token chunks."""
    out = bytearray([0x01])
    pos = 0
    while pos < len(data):
        block = data[pos : pos + 4096]
        pos += len(block)
        chunk = bytearray()
        i = 0
        while i < len(block):
            sub = block[i : i + 8]
            chunk.append(0x00)
            chunk.extend(sub)
            i += 8
        header = 0xB000 | (len(chunk) - 1)
        out.extend(struct.pack("<H", header))
        out.extend(chunk)
    return bytes(out)


def _rec(rec_id: int, payload: bytes) -> bytes:
    return struct.pack("<HI", rec_id, len(payload)) + payload


def _build_dir_stream() -> bytes:
    """Build MS-OVBA 2.3.4.2 dir stream (uncompressed, then OVBA-compressed)."""
    buf = bytearray()
    # PROJECTSYSKIND (Win32=1, or Win64=3; 1 works on both 32/64-bit Excel)
    buf.extend(_rec(0x0001, struct.pack("<I", 1)))
    # PROJECTLCID
    buf.extend(_rec(0x0002, struct.pack("<I", 0x0409)))
    # PROJECTLCIDINVOKE
    buf.extend(_rec(0x0014, struct.pack("<I", 0x0409)))
    # PROJECTCODEPAGE (1252)
    buf.extend(_rec(0x0003, struct.pack("<H", 1252)))
    # PROJECTNAME
    buf.extend(_rec(0x0004, b"VBAProject"))
    # PROJECTDOCSTRING + Unicode
    buf.extend(_rec(0x0005, b""))
    buf.extend(_rec(0x0040, b""))
    # PROJECTHELPFILEPATH + 2
    buf.extend(_rec(0x0006, b""))
    buf.extend(_rec(0x003D, b""))
    # PROJECTHELPCONTEXT
    buf.extend(_rec(0x0007, struct.pack("<I", 0)))
    # PROJECTLIBFLAGS
    buf.extend(_rec(0x0008, struct.pack("<I", 0)))
    # PROJECTVERSION (Major=1, Minor=0 -> size 4 + 2? Spec: Reserved 0x0009, Size=4, Major uint32, Minor uint16)
    buf.extend(struct.pack("<HI", 0x0009, 4) + struct.pack("<IH", 1, 0))
    # PROJECTCONSTANTS + Unicode
    buf.extend(_rec(0x000C, b""))
    buf.extend(_rec(0x003C, b""))

    # REFERENCE stdole
    stdole_lib = b"*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation"
    buf.extend(_rec(0x0016, b"stdole"))
    buf.extend(_rec(0x003E, "stdole".encode("utf-16le")))
    buf.extend(_rec(0x000D, struct.pack("<I", len(stdole_lib)) + stdole_lib + struct.pack("<IH", 0, 0)))

    # REFERENCE Office
    office_lib = b"*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL#Microsoft Office 16.0 Object Library"
    buf.extend(_rec(0x0016, b"Office"))
    buf.extend(_rec(0x003E, "Office".encode("utf-16le")))
    buf.extend(_rec(0x000D, struct.pack("<I", len(office_lib)) + office_lib + struct.pack("<IH", 0, 0)))

    # PROJECTMODULES
    buf.extend(_rec(0x000F, struct.pack("<H", 2)))
    # PROJECTCOOKIE
    buf.extend(_rec(0x0013, struct.pack("<H", 0xFFFF)))

    # Module 1: ThisWorkbook (Document module: 0x0022)
    buf.extend(_module_records("ThisWorkbook", is_doc=True))
    # Module 2: Module1 (Procedural module: 0x0021)
    buf.extend(_module_records("Module1", is_doc=False))

    # Terminator
    buf.extend(struct.pack("<HI", 0x0010, 0))
    return _ovba_compress(bytes(buf))


def _module_records(name: str, is_doc: bool) -> bytes:
    b_name = name.encode("ascii")
    u_name = name.encode("utf-16le")
    out = bytearray()
    out.extend(_rec(0x0019, b_name))           # MODULENAME
    out.extend(_rec(0x0047, u_name))           # MODULENAMEUNICODE
    out.extend(_rec(0x001A, b_name))           # MODULESTREAMNAME
    out.extend(_rec(0x0032, u_name))           # MODULESTREAMNAMEUNICODE
    out.extend(_rec(0x001C, b""))              # MODULEDOCSTRING
    out.extend(_rec(0x0048, b""))              # MODULEDOCSTRINGUNICODE
    out.extend(_rec(0x0031, struct.pack("<I", 0)))  # MODULEOFFSET = 0 (source starts at 0)
    out.extend(_rec(0x001E, struct.pack("<I", 0)))  # MODULEHELPCONTEXT
    out.extend(_rec(0x002C, struct.pack("<H", 0xFFFF)))  # MODULECOOKIE
    out.extend(struct.pack("<HI", 0x0022 if is_doc else 0x0021, 0))  # MODULETYPE
    out.extend(struct.pack("<HI", 0x002B, 0))  # Terminator
    return bytes(out)


def _build_project_stream() -> bytes:
    txt = (
        'ID="{A1B2C3D4-E5F6-4789-ABCD-0123456789AB}"\r\n'
        'Document=ThisWorkbook/&H00000000\r\n'
        'Module=Module1\r\n'
        'Name="VBAProject"\r\n'
        'HelpContextID="0"\r\n'
        'VersionCompatible32="393222000"\r\n'
        'CMG="0000"\r\n'
        'DPB="0000"\r\n'
        'GC="0000"\r\n'
        '\r\n'
        '[Host Extender Info]\r\n'
        '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n'
        '\r\n'
        '[Workspace]\r\n'
        'ThisWorkbook=0, 0, 0, 0, C\r\n'
        'Module1=26, 26, 400, 300, C\r\n'
    )
    return txt.encode("ascii")


def _build_projectwm_stream() -> bytes:
    out = bytearray()
    for name in ("ThisWorkbook", "Module1"):
        out.extend(name.encode("ascii") + b"\x00")
        out.extend(name.encode("utf-16le") + b"\x00\x00")
    out.extend(b"\x00\x00")
    return bytes(out)


def _dir_entry(
    name: str,
    obj_type: int,
    color: int,
    left_id: int,
    right_id: int,
    child_id: int,
    start_sector: int,
    size: int,
) -> bytes:
    u_name = (name.encode("utf-16le") + b"\x00\x00") if name else b""
    name_len = len(u_name)
    padded_name = u_name.ljust(64, b"\x00")[:64]
    return (
        padded_name
        + struct.pack(
            "<HBBIII",
            name_len,
            obj_type,
            color,
            left_id,
            right_id,
            child_id,
        )
        + b"\x00" * 16  # CLSID
        + struct.pack("<I", 0)  # StateBits
        + b"\x00" * 16  # Creation/Modified times
        + struct.pack("<IIQ", start_sector, size & 0xFFFFFFFF, 0)[0:8]  # wait: start_sect(4) + size(8 in v3=4+4)
    )


def _pack_dir_entry(
    name: str,
    obj_type: int,
    color: int,
    left_id: int,
    right_id: int,
    child_id: int,
    start_sector: int,
    size: int,
) -> bytes:
    u_name = (name.encode("utf-16le") + b"\x00\x00") if name else b""
    name_len = len(u_name)
    padded_name = u_name.ljust(64, b"\x00")[:64]
    return (
        padded_name
        + struct.pack("<HBBIII", name_len, obj_type, color, left_id, right_id, child_id)
        + b"\x00" * 16
        + struct.pack("<I", 0)
        + b"\x00" * 16
        + struct.pack("<IQ", start_sector, size)
    )


def _build_cfb(streams: Sequence[tuple[str, bytes]]) -> bytes:
    """Build a minimal v3 OLE2 Compound File (512-byte sectors, no mini-stream).

    Pad every stream to >= 4096 bytes so Excel reads directly from regular FAT sectors.
    """
    sector_size = 512
    # Pad streams to 4096 bytes so MiniStreamCutoff (4096) is avoided cleanly.
    # Actually CFB v3 allows storing >=4096 streams in regular FAT; we store actual size
    # in dir entry? No, if dir entry size < 4096, reader looks in MiniStream!
    # So we store streams in MiniStream OR pad? Padding VBA/dir or PROJECT breaks parsing!
    # Instead, let's put ALL streams < 4096 into MiniStream cleanly (super easy!).
    mini_sector_size = 64
    mini_stream = bytearray()
    mini_fat: list[int] = []
    stream_meta: list[tuple[str, int, int]] = []  # (name, start_mini_sect, actual_size)

    for name, data in streams:
        start_ms = len(mini_stream) // mini_sector_size
        n_ms = max(1, (len(data) + mini_sector_size - 1) // mini_sector_size)
        padded = data.ljust(n_ms * mini_sector_size, b"\x00")
        mini_stream.extend(padded)
        for idx in range(n_ms):
            mini_fat.append(start_ms + idx + 1 if idx < n_ms - 1 else ENDOFCHAIN)
        stream_meta.append((name, start_ms, len(data)))

    # Sectors layout:
    # Sector 0: MiniFAT (1 sector covers up to 128 mini-sectors = 8192 bytes)
    # Sectors 1..K: MiniStream
    # Sectors K+1..K+2: Directory (2 sectors = 8 entries)
    # Sector K+3: FAT
    n_minifat_sects = max(1, (len(mini_fat) * 4 + sector_size - 1) // sector_size)
    n_ministream_sects = max(1, (len(mini_stream) + sector_size - 1) // sector_size)
    n_dir_sects = 2  # 8 directory entries (128 bytes each = 1024 bytes)

    minifat_start = 0
    ministream_start = minifat_start + n_minifat_sects
    dir_start = ministream_start + n_ministream_sects
    fat_start = dir_start + n_dir_sects
    total_sects = fat_start + 1

    fat: list[int] = [FREESECT] * (sector_size // 4)
    for i in range(n_minifat_sects):
        fat[minifat_start + i] = minifat_start + i + 1 if i < n_minifat_sects - 1 else ENDOFCHAIN
    for i in range(n_ministream_sects):
        fat[ministream_start + i] = (
            ministream_start + i + 1 if i < n_ministream_sects - 1 else ENDOFCHAIN
        )
    for i in range(n_dir_sects):
        fat[dir_start + i] = dir_start + i + 1 if i < n_dir_sects - 1 else ENDOFCHAIN
    fat[fat_start] = FATSECT

    # Directory entries (MS-CFB requires BST sorted by (len(name), name.upper())):
    # Root children: "VBA" (len 3) < "PROJECT" (len 7) < "PROJECTwm" (len 9)
    #   root=1 (VBA, left=NOSTREAM, right=6(PROJECT))
    #   6 (PROJECT, left=NOSTREAM, right=7(PROJECTwm))
    # VBA children: "dir" (3) < "Module1" (7) < "ThisWorkbook" (12, 'T') < "_VBA_PROJECT" (12, '_')
    #   child=3 (dir, left=NOSTREAM, right=5(Module1))
    #   5 (Module1, left=NOSTREAM, right=4(ThisWorkbook))
    #   4 (ThisWorkbook, left=NOSTREAM, right=2(_VBA_PROJECT))
    meta_map = {m[0]: (m[1], m[2]) for m in stream_meta}

    entries = [
        _pack_dir_entry("Root Entry", 5, 1, NOSTREAM, NOSTREAM, 1, ministream_start, len(mini_stream)),
        _pack_dir_entry("VBA", 1, 1, NOSTREAM, 6, 3, 0, 0),
        _pack_dir_entry("_VBA_PROJECT", 2, 1, NOSTREAM, NOSTREAM, NOSTREAM, *meta_map["_VBA_PROJECT"]),
        _pack_dir_entry("dir", 2, 1, NOSTREAM, 5, NOSTREAM, *meta_map["dir"]),
        _pack_dir_entry("ThisWorkbook", 2, 1, NOSTREAM, 2, NOSTREAM, *meta_map["ThisWorkbook"]),
        _pack_dir_entry("Module1", 2, 1, NOSTREAM, 4, NOSTREAM, *meta_map["Module1"]),
        _pack_dir_entry("PROJECT", 2, 1, NOSTREAM, 7, NOSTREAM, *meta_map["PROJECT"]),
        _pack_dir_entry("PROJECTwm", 2, 1, NOSTREAM, NOSTREAM, NOSTREAM, *meta_map["PROJECTwm"]),
    ]

    header = (
        b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"  # Signature
        + b"\x00" * 16                       # CLSID
        + struct.pack("<HHHHHH", 0x003E, 0x0003, 0xFFFE, 0x0009, 0x0006, 0x0000)
        + b"\x00" * 4                        # Reserved
        + struct.pack("<III", 0, 1, dir_start)  # numDirSects(0 in v3), numFATSects=1, firstDirSect
        + struct.pack("<IIII", 0, 4096, minifat_start, n_minifat_sects)
        + struct.pack("<II", ENDOFCHAIN, 0)  # DIFAT
        + struct.pack("<I", fat_start)
        + struct.pack("<" + "I" * 108, *([FREESECT] * 108))
    )

    minifat_bytes = struct.pack("<" + "I" * len(mini_fat), *mini_fat).ljust(
        n_minifat_sects * sector_size, b"\xFF"
    )
    ministream_bytes = bytes(mini_stream).ljust(n_ministream_sects * sector_size, b"\x00")
    dir_bytes = b"".join(entries).ljust(n_dir_sects * sector_size, b"\x00")
    fat_bytes = struct.pack("<" + "I" * len(fat), *fat)

    assert len(header) == 512
    assert total_sects * sector_size == len(minifat_bytes) + len(ministream_bytes) + len(dir_bytes) + len(fat_bytes)
    return header + minifat_bytes + ministream_bytes + dir_bytes + fat_bytes


def build_vba_project_bin() -> bytes:
    """Build a complete vbaProject.bin binary with auto-play timer macros."""
    streams = [
        ("_VBA_PROJECT", b"\xCC\x61\xFF\xFF\x00\x00\x00"),
        ("dir", _build_dir_stream()),
        ("ThisWorkbook", _ovba_compress(THISWORKBOOK_CODE.encode("ascii"))),
        ("Module1", _ovba_compress(MODULE1_CODE.encode("ascii"))),
        ("PROJECT", _build_project_stream()),
        ("PROJECTwm", _build_projectwm_stream()),
    ]
    return _build_cfb(streams)
