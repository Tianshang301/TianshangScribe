"""Unit tests for the OLE compound file writer (cfb_writer.make_ole).

Verifies that payloads of all sizes round-trip through the companion reader
(``ole_util``), covering both the mini-stream and regular-chain layouts.
"""

from __future__ import annotations

import pytest

from tianshang_scribe.rendering.mtef.cfb_writer import make_ole
from tianshang_scribe.rendering.mtef.ole_util import open_ole


@pytest.mark.parametrize(
    'size',
    [1, 63, 64, 65, 100, 511, 512, 513, 1000, 2048, 4000, 4095, 4096, 5000, 10000],
)
def test_roundtrip_various_sizes(size: int) -> None:
    payload = bytes(range(256)) * (size // 256 + 1)
    payload = payload[:size]
    ole = make_ole('Equation Native', payload)
    back = open_ole(ole).get_stream('Equation Native')
    assert back == payload


def test_extract_native_stream() -> None:
    from tianshang_scribe.rendering.mtef.ole_util import extract_native_stream

    payload = b'equation-data' * 20
    ole = make_ole('Equation Native', payload)
    assert extract_native_stream(ole) == payload


def test_bad_magic_not_generated() -> None:
    ole = make_ole('Equation Native', b'x')
    assert ole[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'


def test_empty_name_raises() -> None:
    with pytest.raises(ValueError):
        make_ole('', b'x')


def test_long_name_raises() -> None:
    with pytest.raises(ValueError):
        make_ole('X' * 40, b'x')


def test_missing_stream_raises() -> None:
    ole = make_ole('Equation Native', b'data')
    with pytest.raises(KeyError):
        open_ole(ole).get_stream('NoSuchStream')
