"""Unit tests for the OLE Compound File (CFB) reader.

The tests build minimal in-memory OLE files with a single ``Equation Native``
stream (large enough to use the regular FAT chain) and assert the reader can
extract it.
"""

from __future__ import annotations

import struct

import pytest

from tianshang_scribe.rendering.mtef.ole_util import (
    OleCompoundFile,
    OleError,
    extract_native_stream,
)

SECTOR = 512
_FREESECT = 0xFFFFFFFF
_ENDOFCHAIN = 0xFFFFFFFE
_FATSECT = 0xFFFFFFFD
_ROOT = 5
_STREAM = 2


def _make_cfb(stream_name: str, payload: bytes) -> bytes:
    """Build a minimal OLE compound file with one regular-chain stream."""
    num_payload = (len(payload) + SECTOR - 1) // SECTOR
    name_utf16 = stream_name.encode('utf-16-le') + b'\x00\x00'
    # Header (sector 0). 109-entry DIFAT at offset 76.
    header = bytearray(SECTOR)
    header[0:8] = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
    struct.pack_into('<HH', header, 0x18, 3, 0x3E)  # major 3, minor 0x3E
    struct.pack_into('<H', header, 0x1C, 0xFFFE)  # little-endian
    struct.pack_into('<H', header, 0x1E, 9)  # sector shift -> 512
    struct.pack_into('<H', header, 0x20, 6)  # mini sector shift -> 64
    struct.pack_into('<I', header, 0x2C, 1)  # number of FAT sectors
    struct.pack_into('<I', header, 0x30, num_payload)  # first directory sector
    struct.pack_into('<I', header, 0x38, 4096)  # mini stream cutoff
    struct.pack_into('<I', header, 0x3C, _FREESECT)  # first mini FAT
    struct.pack_into('<I', header, 0x40, 0)  # number of mini FAT sectors
    struct.pack_into('<I', header, 0x44, _FREESECT)  # first DIFAT sector
    struct.pack_into('<I', header, 0x48, 0)  # number of DIFAT sectors
    struct.pack_into('<I', header, 0x4C + 4 * 0, num_payload + 1)  # DIFAT[0] -> FAT sector
    for i in range(1, 109):
        struct.pack_into('<I', header, 0x4C + 4 * i, _FREESECT)

    # Layout (sector index — header is a reserved "sector", data starts at 0):
    #   0 .. k-1 : payload sectors
    #   k        : directory sector
    #   k+1      : FAT sector
    dir_sector = num_payload
    fat_sector = num_payload + 1

    # FAT entries.
    fat = bytearray(SECTOR)
    entries = [_FREESECT] * (SECTOR // 4)
    # payload chain
    for i in range(num_payload):
        entries[i] = (i + 1) if i + 1 < num_payload else _ENDOFCHAIN
    # directory chain
    entries[dir_sector] = _ENDOFCHAIN
    # FAT sector self-reference
    entries[fat_sector] = _FATSECT
    for i, v in enumerate(entries):
        struct.pack_into('<I', fat, 4 * i, v)

    # Directory entry (128 bytes) for the stream.
    direntry = bytearray(128)
    direntry[0 : len(name_utf16)] = name_utf16
    struct.pack_into('<H', direntry, 64, len(name_utf16))  # name size incl NUL
    direntry[66] = _STREAM  # object type
    struct.pack_into('<I', direntry, 116, 0)  # start sector (payload starts at 0)
    struct.pack_into('<Q', direntry, 120, len(payload))  # size

    # Directory sector.
    dir_sector_bytes = bytearray(SECTOR)
    dir_sector_bytes[0:128] = direntry

    out = bytearray(header)
    out += payload
    # pad payload to sector boundary
    pad = (SECTOR - (len(payload) % SECTOR)) % SECTOR
    out += b'\x00' * pad
    out += dir_sector_bytes
    out += fat
    return bytes(out)


def test_reads_stream_from_minimal_cfb() -> None:
    payload = b'MTEF-synthetic-data-' * 300  # ~6600 bytes -> regular chain
    blob = _make_cfb('Equation Native', payload)
    ole = OleCompoundFile(blob)
    assert ole.get_stream('Equation Native') == payload
    assert ole.list_streams() == ['Equation Native']


def test_extract_native_stream_helper() -> None:
    payload = b'\x03\x01\x00MTEF!' * 700  # 5600 bytes -> regular chain
    blob = _make_cfb('Equation Native', payload)
    assert extract_native_stream(blob) == payload


def test_bad_magic_raises() -> None:
    with pytest.raises(OleError):
        OleCompoundFile(b'not an ole file at all' + b'\x00' * 500)


def test_missing_stream_raises_keyerror() -> None:
    blob = _make_cfb('Other Stream', b'x' * 5000)
    ole = OleCompoundFile(blob)
    with pytest.raises(KeyError):
        ole.get_stream('Equation Native')


def test_small_stream_via_mini_chain() -> None:
    """Streams smaller than the mini-stream cutoff use the mini FAT chain.

    This is the common case for MathType equations, so build a CFB with a
    root storage whose mini stream holds the payload, and a stream directory
    entry pointing into it.
    """
    payload = b'mini-equation-data' * 25  # 425 bytes < 512, single mini sector
    stream_name = 'Equation Native'
    stream_name.encode('utf-16-le') + b'\x00\x00'

    header = bytearray(SECTOR)
    header[0:8] = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
    struct.pack_into('<HH', header, 0x18, 3, 0x3E)
    struct.pack_into('<H', header, 0x1C, 0xFFFE)
    struct.pack_into('<H', header, 0x1E, 9)
    struct.pack_into('<H', header, 0x20, 6)
    struct.pack_into('<I', header, 0x2C, 1)  # one FAT sector
    struct.pack_into('<I', header, 0x30, 1)  # first directory sector
    struct.pack_into('<I', header, 0x38, 4096)  # mini stream cutoff
    struct.pack_into('<I', header, 0x3C, 3)  # first mini FAT sector
    struct.pack_into('<I', header, 0x40, 1)  # mini FAT count
    struct.pack_into('<I', header, 0x44, _FREESECT)  # first DIFAT sector
    struct.pack_into('<I', header, 0x48, 0)  # number of DIFAT sectors
    struct.pack_into('<I', header, 0x4C + 4 * 0, 2)  # DIFAT[0] -> FAT sector 2

    # Layout (sector index):
    #   0: root's mini stream payload sector
    #   1: directory sector
    #   2: FAT sector
    #   3: mini FAT sector
    mini_stream = bytearray(payload + b'\x00' * (SECTOR - len(payload)))

    # Directory entries: root storage + one stream.
    dir_sector = bytearray(SECTOR)

    def _entry(name: str, obj_type: int, start: int, size: int, offset: int) -> None:
        raw = name.encode('utf-16-le') + b'\x00\x00'
        dir_sector[offset : offset + len(raw)] = raw
        struct.pack_into('<H', dir_sector, offset + 64, len(raw))
        dir_sector[offset + 66] = obj_type
        struct.pack_into('<I', dir_sector, offset + 116, start)
        struct.pack_into('<Q', dir_sector, offset + 120, size)

    _entry('Root Entry', _ROOT, 0, SECTOR, 0)
    _entry(stream_name, _STREAM, 0, len(payload), 128)

    fat = bytearray(SECTOR)
    entries = [_FREESECT] * (SECTOR // 4)
    entries[0] = _ENDOFCHAIN  # mini stream sector
    entries[1] = _ENDOFCHAIN  # directory
    entries[2] = _FATSECT  # FAT self
    entries[3] = _ENDOFCHAIN  # mini FAT
    for i, v in enumerate(entries):
        struct.pack_into('<I', fat, 4 * i, v)

    mini_fat = bytearray(SECTOR)
    num_mini_sectors = (len(payload) + 63) // 64
    for i in range(num_mini_sectors):
        struct.pack_into('<I', mini_fat, 4 * i, i + 1 if i + 1 < num_mini_sectors else _ENDOFCHAIN)

    out = bytearray(header)
    out += mini_stream
    out += dir_sector
    out += fat
    out += mini_fat
    ole = OleCompoundFile(bytes(out))
    assert ole.get_stream('Equation Native') == payload
