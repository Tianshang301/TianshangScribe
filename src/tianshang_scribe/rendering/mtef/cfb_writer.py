"""Minimal OLE Compound File Binary (CFB) writer.

Builds a single-stream OLE compound file — the container format used by
MathType equation objects embedded in Word documents.  Payloads smaller than
the mini-stream cutoff (4096 bytes, the common case for MTEF equations) are
stored in the root storage's mini stream using the mini FAT chain; larger
payloads use a regular FAT chain.  Both layouts are read back correctly by
``ole_util.open_ole``.

Reference: MS-CFB (Microsoft Compound File Binary File Format).
"""

from __future__ import annotations

import struct

# CFB constants
_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
_FREESECT = 0xFFFFFFFF
_ENDOFCHAIN = 0xFFFFFFFE
_FATSECT = 0xFFFFFFFD
_SECTOR_SIZE = 512
_MINISECTOR_SIZE = 64
_MINI_CUTOFF = 4096

# Directory object types
_STREAM = 2
_ROOT = 5

# Directory entry field offsets
_DE_NAME_SIZE = 64
_DE_TYPE = 66
_DE_START = 116
_DE_SIZE = 120


def make_ole(stream_name: str, payload: bytes) -> bytes:
    """Wrap ``payload`` in an OLE compound file under ``stream_name``.

    Small payloads (<4096 bytes) use the mini-stream layout that mirrors real
    MathType equation objects; larger payloads use a regular FAT chain.
    """
    if not stream_name:
        raise ValueError('stream_name must not be empty')
    name_utf16 = stream_name.encode('utf-16-le') + b'\x00\x00'
    if len(name_utf16) > 62:
        raise ValueError('stream_name too long (max 31 chars)')

    if len(payload) < _MINI_CUTOFF:
        return _make_ole_mini(stream_name, name_utf16, payload)
    return _make_ole_regular(stream_name, name_utf16, payload)


def _make_ole_regular(stream_name: str, name_utf16: bytes, payload: bytes) -> bytes:
    """Build a CFB with a single regular-chain stream (payload >= 4096)."""
    num_payload = (len(payload) + _SECTOR_SIZE - 1) // _SECTOR_SIZE

    header = bytearray(_SECTOR_SIZE)
    header[0:8] = _MAGIC
    struct.pack_into('<HH', header, 0x18, 3, 0x3E)
    struct.pack_into('<H', header, 0x1C, 0xFFFE)
    struct.pack_into('<H', header, 0x1E, 9)
    struct.pack_into('<H', header, 0x20, 6)
    struct.pack_into('<I', header, 0x2C, 1)
    struct.pack_into('<I', header, 0x30, num_payload)
    struct.pack_into('<I', header, 0x38, _MINI_CUTOFF)
    struct.pack_into('<I', header, 0x3C, _FREESECT)
    struct.pack_into('<I', header, 0x40, 0)
    struct.pack_into('<I', header, 0x44, _FREESECT)
    struct.pack_into('<I', header, 0x48, 0)
    struct.pack_into('<I', header, 0x4C + 4 * 0, num_payload + 1)
    for i in range(1, 109):
        struct.pack_into('<I', header, 0x4C + 4 * i, _FREESECT)

    dir_sector = num_payload
    fat_sector = num_payload + 1

    fat = bytearray(_SECTOR_SIZE)
    entries = [_FREESECT] * (_SECTOR_SIZE // 4)
    for i in range(num_payload):
        entries[i] = i + 1 if i + 1 < num_payload else _ENDOFCHAIN
    entries[dir_sector] = _ENDOFCHAIN
    entries[fat_sector] = _FATSECT
    for i, value in enumerate(entries):
        struct.pack_into('<I', fat, 4 * i, value)

    dir_sector_bytes = bytearray(_SECTOR_SIZE)
    _write_entry(dir_sector_bytes, 0, stream_name, name_utf16, _STREAM, 0, len(payload))

    out = bytearray(header)
    out += payload
    pad = (_SECTOR_SIZE - (len(payload) % _SECTOR_SIZE)) % _SECTOR_SIZE
    out += b'\x00' * pad
    out += dir_sector_bytes
    out += fat
    return bytes(out)


def _make_ole_mini(stream_name: str, name_utf16: bytes, payload: bytes) -> bytes:
    """Build a CFB whose stream lives in the root storage's mini stream.

    Layout (sector index):
        0 .. m-1      : root's mini stream payload sectors
        m             : directory sector
        m+1           : FAT sector
        m+2           : mini FAT sector
    """
    num_stream_sectors = (len(payload) + _SECTOR_SIZE - 1) // _SECTOR_SIZE
    mini_stream = bytearray(payload)
    mini_stream += b'\x00' * (num_stream_sectors * _SECTOR_SIZE - len(payload))

    dir_sector = num_stream_sectors
    fat_sector = num_stream_sectors + 1
    mini_fat_sector = num_stream_sectors + 2

    dir_bytes = bytearray(_SECTOR_SIZE)
    _write_entry(dir_bytes, 0, 'Root Entry', 'Root Entry'.encode('utf-16-le') + b'\x00\x00',
                 _ROOT, 0, len(mini_stream))
    _write_entry(dir_bytes, 1, stream_name, name_utf16, _STREAM, 0, len(payload))

    fat = bytearray(_SECTOR_SIZE)
    entries = [_FREESECT] * (_SECTOR_SIZE // 4)
    for i in range(num_stream_sectors):
        entries[i] = i + 1 if i + 1 < num_stream_sectors else _ENDOFCHAIN
    entries[dir_sector] = _ENDOFCHAIN
    entries[fat_sector] = _FATSECT
    entries[mini_fat_sector] = _ENDOFCHAIN
    for i, value in enumerate(entries):
        struct.pack_into('<I', fat, 4 * i, value)

    mini_fat = bytearray(_SECTOR_SIZE)
    num_mini_sectors = (len(payload) + _MINISECTOR_SIZE - 1) // _MINISECTOR_SIZE
    for i in range(num_mini_sectors):
        next_val = i + 1 if i + 1 < num_mini_sectors else _ENDOFCHAIN
        struct.pack_into('<I', mini_fat, 4 * i, next_val)

    header = bytearray(_SECTOR_SIZE)
    header[0:8] = _MAGIC
    struct.pack_into('<HH', header, 0x18, 3, 0x3E)
    struct.pack_into('<H', header, 0x1C, 0xFFFE)
    struct.pack_into('<H', header, 0x1E, 9)
    struct.pack_into('<H', header, 0x20, 6)
    struct.pack_into('<I', header, 0x2C, 1)  # one FAT sector
    struct.pack_into('<I', header, 0x30, dir_sector)  # first directory sector
    struct.pack_into('<I', header, 0x38, _MINI_CUTOFF)
    struct.pack_into('<I', header, 0x3C, mini_fat_sector)  # first mini FAT sector
    struct.pack_into('<I', header, 0x40, 1)  # mini FAT count
    struct.pack_into('<I', header, 0x44, _FREESECT)
    struct.pack_into('<I', header, 0x48, 0)
    struct.pack_into('<I', header, 0x4C + 4 * 0, fat_sector)  # DIFAT[0] -> FAT sector
    for i in range(1, 109):
        struct.pack_into('<I', header, 0x4C + 4 * i, _FREESECT)

    out = bytearray(header)
    out += mini_stream
    out += dir_bytes
    out += fat
    out += mini_fat
    return bytes(out)


def _write_entry(
    sector: bytearray,
    index: int,
    name: str,
    name_utf16: bytes,
    obj_type: int,
    start: int,
    size: int,
) -> None:
    """Write a 128-byte directory entry at ``index`` (0 or 1) within ``sector``."""
    offset = index * 128
    sector[offset : offset + len(name_utf16)] = name_utf16
    struct.pack_into('<H', sector, offset + _DE_NAME_SIZE, len(name_utf16))
    sector[offset + _DE_TYPE] = obj_type
    struct.pack_into('<I', sector, offset + _DE_START, start)
    struct.pack_into('<Q', sector, offset + _DE_SIZE, size)
