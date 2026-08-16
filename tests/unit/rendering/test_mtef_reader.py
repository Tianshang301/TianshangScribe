"""Unit tests for the MTEF binary reader (MathType -> LaTeX)."""

from __future__ import annotations

import struct

import pytest

from tianshang_scribe.rendering.mtef.mtef_reader import MTEFParseError, mtef_to_latex

_FN_VARIABLE = 3
_FN_NUMBER = 8


def _char(c: str | int) -> bytes:
    """A CHAR record for an ASCII char (variable typeface, no options)."""
    code = ord(c) if isinstance(c, str) else c
    typeface = _FN_NUMBER if chr(code) in '0123456789' else _FN_VARIABLE
    # record(2), options=0, typeface, mtcode uint16
    return bytes([2, 0, typeface]) + struct.pack('<H', code)


def _char_mt(code: int, typeface: int = _FN_VARIABLE) -> bytes:
    """A CHAR record with explicit MathType char code and typeface."""
    return bytes([2, 0, typeface]) + struct.pack('<H', code)


def _line() -> bytes:
    """A non-null LINE record with no options."""
    return bytes([1, 0])


def _end() -> bytes:
    return bytes([0])


def _tmpl(selector: int, variation: int = 0) -> bytes:
    """A TMPL record (options=0, selector, variation, template options=0)."""
    return bytes([3, 0, selector, variation, 0])


def _header() -> bytes:
    app = b'MathType\x00'
    return bytes([5, 1, 1, 0, 0]) + app + bytes([0])


def _wrap(records: bytes) -> bytes:
    return _header() + records


def test_header_and_plain_text() -> None:
    latex = mtef_to_latex(_wrap(_char('x') + _char('=') + _char('y') + _end()))
    assert 'x' in latex and 'y' in latex


def test_fraction() -> None:
    data = _wrap(_tmpl(11) + _char('a') + _char('b') + _end() + _end())
    latex = mtef_to_latex(data)
    assert '\\frac{a}{b}' in latex.replace(' ', '')


def test_sqrt() -> None:
    # tmROOT (10), variation 0 = square root: root slot, then content
    data = _wrap(_tmpl(10) + _char('x') + _end() + _end())
    latex = mtef_to_latex(data)
    assert '\\sqrt{x}' in latex.replace(' ', '')


def test_sum_with_limits() -> None:
    # tmSUM (16): slots are main, lower, upper (order in MTEF is main/lower/upper)
    data = _wrap(_tmpl(16) + _char('i') + _char('1') + _char('n') + _char('x') + _end() + _end())
    latex = mtef_to_latex(data)
    assert 'sum' in latex


def test_superscript() -> None:
    # tmSUP (28): slots are base, superscript
    data = _wrap(_tmpl(28) + _char('x') + _char('2') + _end() + _end())
    latex = mtef_to_latex(data)
    assert 'x^{2}' in latex.replace(' ', '')


def test_parentheses() -> None:
    # tmPAREN (1): slots are main, left, right (per MTEF-py: main=0,left=1,right=2)
    data = _wrap(_tmpl(1) + _char('a') + _char('(') + _char(')') + _end() + _end())
    latex = mtef_to_latex(data)
    assert '(' in latex and ')' in latex


def test_greek_letter() -> None:
    # typeface 4 = LCGREEK, mtcode 0x03B1 = alpha
    data = _wrap(_char_mt(0x03B1, 4) + _end())
    latex = mtef_to_latex(data)
    assert '\\alpha' in latex


def test_math_symbol_int() -> None:
    # typeface 6 = SYMBOL, mtcode 0x222B = integral
    data = _wrap(_char_mt(0x222B, 6) + _end())
    latex = mtef_to_latex(data)
    assert '\\int' in latex


def test_empty_payload_raises() -> None:
    with pytest.raises(MTEFParseError):
        mtef_to_latex(b'\x05\x01\x01\x00\x00MathType\x00\x00')


def test_truncated_payload_raises() -> None:
    with pytest.raises(MTEFParseError):
        mtef_to_latex(_header() + bytes([3, 0, 11]) + _char('a') + _end())
