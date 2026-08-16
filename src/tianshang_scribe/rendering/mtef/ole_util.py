"""Minimal OLE Compound File Binary (CFB) reader.

Extracts named streams from an OLE compound file — the container format used
by MathType equation objects embedded in Word documents
(``word/embeddings/oleObject*.bin``). Only the subset needed to locate the
``Equation Native`` stream is implemented; writes are out of scope.

Reference: MS-CFB (Microsoft Compound File Binary File Format), and the
Apache-2.0 ``mtef-go``/``MTEF-py-FIX`` projects.
"""

from __future__ import annotations

import struct
from typing import Any

# --- CFB magic & sector size constants ---------------------------------------
_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
_FREESECT = 0xFFFFFFFF
_ENDOFCHAIN = 0xFFFFFFFE
_FATSECT = 0xFFFFFFFD
_DIFSECT = 0xFFFFFFFC
_MAXREGSECT = 0xFFFFFFFA

# Directory object types
_STGTY_INVALID = 0
_STGTY_STORAGE = 1
_STGTY_STREAM = 2
_STGTY_ROOT = 5


class OleError(ValueError):
    """Raised when a buffer is not a valid OLE compound file."""


class OleStream:
    """A stream inside an OLE compound file, readable as bytes."""

    __slots__ = ('data', 'name')

    def __init__(self, name: str, data: bytes) -> None:
        """Store one named stream's raw bytes."""
        self.name = name
        self.data = data


class OleCompoundFile:
    """Parsed OLE compound file exposing its named streams."""

    def __init__(self, data: bytes) -> None:
        """Parse ``data`` as an OLE compound file."""
        self.data = data
        self.streams: dict[str, OleStream] = {}
        self._parse(data)

    # -- public API -----------------------------------------------------------

    def get_stream(self, name: str) -> bytes:
        """Return the raw bytes of the stream named ``name``."""
        stream = self.streams.get(name)
        if stream is None:
            raise KeyError(f'OLE stream not found: {name}')
        return stream.data

    def list_streams(self) -> list[str]:
        """Return the names of all parsed streams."""
        return sorted(self.streams)

    @classmethod
    def open(cls, data: bytes) -> OleCompoundFile:
        """Parse ``data`` as an OLE compound file."""
        return cls(data)

    # -- low-level parsing ----------------------------------------------------

    def _parse(self, data: bytes) -> None:
        if len(data) < 512 or data[:8] != _MAGIC:
            raise OleError('not an OLE compound file (bad magic)')

        major = struct.unpack_from('<H', data, 24)[0]
        if major == 3:
            self.sector_size = 512
        elif major == 4:
            self.sector_size = 4096
        else:
            raise OleError(f'unsupported CFB major version: {major}')

        sector_shift = struct.unpack_from('<H', data, 0x1E)[0]
        if sector_shift != 0 and sector_shift != self.sector_size.bit_length() - 1:
            raise OleError(f'bad sector shift: {sector_shift}')

        # CFB header field offsets:
        #   0x18 major, 0x1A minor, 0x1C byte order, 0x1E sector shift,
        #   0x20 mini sector shift, 0x2C num FAT, 0x30 first directory,
        #   0x38 mini cutoff, 0x3C first mini FAT, 0x40 num mini FAT,
        #   0x44 first DIFAT, 0x48 num DIFAT, 0x4C DIFAT entries (109).
        num_fat_sectors = struct.unpack_from('<I', data, 0x2C)[0]
        first_dir_sector = struct.unpack_from('<I', data, 0x30)[0]
        mini_stream_cutoff = struct.unpack_from('<I', data, 0x38)[0]
        first_mini_fat = struct.unpack_from('<I', data, 0x3C)[0]
        first_difat = struct.unpack_from('<I', data, 0x44)[0]
        num_difat = struct.unpack_from('<I', data, 0x48)[0]

        # DIFAT: first 109 entries live in the header; extra DIFAT sectors
        # continue the array.
        difat_entries: list[int] = []
        for i in range(109):
            difat_entries.append(struct.unpack_from('<I', data, 0x4C + 4 * i)[0])

        next_difat = first_difat
        for _ in range(num_difat):
            if next_difat in (_FREESECT, _ENDOFCHAIN):
                break
            sector = self._read_sector(data, next_difat)
            for i in range(self.sector_size // 4 - 1):
                difat_entries.append(struct.unpack_from('<I', sector, 4 * i)[0])
            next_difat = struct.unpack_from('<I', sector, self.sector_size - 4)[0]

        fat_sectors = [d for d in difat_entries if d != _FREESECT]
        if len(fat_sectors) < num_fat_sectors:
            # Fall back: read FAT sector chain from DIFAT if header count is off.
            pass

        # Build the FAT table.
        fat_entries: list[int] = []
        for sec in fat_sectors:
            if sec in (_FREESECT, _ENDOFCHAIN) or sec >= _MAXREGSECT:
                continue
            sector = self._read_sector(data, sec)
            for i in range(self.sector_size // 4):
                fat_entries.append(struct.unpack_from('<I', sector, 4 * i)[0])

        self.sector_size = self.sector_size
        self.fat = fat_entries
        self.mini_stream_cutoff = mini_stream_cutoff

        # Directory entries live in a chain of sectors starting at
        # first_dir_sector.
        dir_chain = self._chain_sectors(data, first_dir_sector, fat_entries)
        dir_bytes = b''.join(self._read_sector(data, s) for s in dir_chain)

        entries = self._parse_directory(dir_bytes)

        root = next((e for e in entries if e['type'] == _STGTY_ROOT), None)

        # The mini stream (root stream) holds all streams smaller than the
        # mini-stream cutoff. It only exists when a root storage is present.
        if root is not None:
            mini_stream = self._read_chain(data, root['start_sector'], root['size'], fat_entries)
        else:
            mini_stream = b''
        self.mini_stream = mini_stream

        # Build mini FAT.
        mini_fat_entries: list[int] = []
        for s in self._chain_sectors(data, first_mini_fat, fat_entries):
            sector = self._read_sector(data, s)
            for i in range(self.sector_size // 4):
                mini_fat_entries.append(struct.unpack_from('<I', sector, 4 * i)[0])
        self.mini_fat = mini_fat_entries
        self.mini_sector_size = 64

        # Register streams.
        for e in entries:
            if e['type'] == _STGTY_STREAM:
                name = e['name']
                if e['size'] < mini_stream_cutoff:
                    data_bytes = self._read_mini_chain(e['start_sector'], e['size'])
                else:
                    data_bytes = self._read_chain(data, e['start_sector'], e['size'], fat_entries)
                self.streams[name] = OleStream(name, data_bytes)

    def _read_sector(self, data: bytes, index: int) -> bytes:
        start = (index + 1) * self.sector_size
        return data[start : start + self.sector_size]

    def _chain_sectors(self, data: bytes, start: int, fat: list[int]) -> list[int]:
        chain: list[int] = []
        seen: set[int] = set()
        n = start
        while n != _ENDOFCHAIN and n != _FREESECT:
            if n in seen or n >= len(fat):
                break
            seen.add(n)
            chain.append(n)
            n = fat[n]
        return chain

    def _read_chain(self, data: bytes, start: int, size: int, fat: list[int]) -> bytes:
        out = bytearray()
        for s in self._chain_sectors(data, start, fat):
            out += self._read_sector(data, s)
            if len(out) >= size:
                break
        return bytes(out[:size])

    def _read_mini_chain(self, start: int, size: int) -> bytes:
        out = bytearray()
        seen: set[int] = set()
        n = start
        while n != _ENDOFCHAIN and n != _FREESECT and n not in seen:
            seen.add(n)
            offset = n * self.mini_sector_size
            out += self.mini_stream[offset : offset + self.mini_sector_size]
            if len(out) >= size:
                break
            if n >= len(self.mini_fat):
                break
            n = self.mini_fat[n]
        return bytes(out[:size])

    def _parse_directory(self, dir_bytes: bytes) -> list[dict[str, Any]]:
        entries: list[dict[str, Any]] = []
        count = len(dir_bytes) // 128
        for i in range(count):
            entry = dir_bytes[i * 128 : i * 128 + 128]
            if len(entry) < 128:
                break
            name_len = struct.unpack_from('<H', entry, 64)[0]
            name_raw = entry[: name_len - 2] if name_len >= 2 else b''
            try:
                name = name_raw.decode('utf-16-le', errors='replace')
            except Exception:  # never fail on a malformed name
                name = ''
            obj_type = entry[66]
            start_sector = struct.unpack_from('<I', entry, 116)[0]
            size = struct.unpack_from('<Q', entry, 120)[0]
            if obj_type == _STGTY_INVALID:
                continue
            entries.append(
                {
                    'name': name,
                    'type': obj_type,
                    'start_sector': start_sector,
                    'size': size,
                }
            )
        return entries


def open_ole(data: bytes) -> OleCompoundFile:
    """Parse an OLE compound file and return the object."""
    return OleCompoundFile(data)


def extract_native_stream(data: bytes) -> bytes:
    """Return the ``Equation Native`` stream bytes from a MathType OLE object.

    Raises ``KeyError`` if the OLE file has no ``Equation Native`` stream.
    """
    ole = OleCompoundFile(data)
    return ole.get_stream('Equation Native')
