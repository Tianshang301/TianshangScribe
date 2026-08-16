"""LaTeX <-> MTEF char mapping tables for the MTEF writer.

The reader (``mtef_reader.py``) maps MTEF char codes to LaTeX; this module
provides the reverse direction used by ``mtef_writer.py``.  Ordinary ASCII
characters map to their Unicode code point with a variable/number typeface,
Greek letters use the lower/upper Greek typefaces, and math symbols use the
Symbol (and MT Extra) typeface tables.

The symbol table extends the reader's subset with additional common commands
so that writing most LaTeX formulas does not fall back to plain text.
"""

from __future__ import annotations

from .mtef_reader import (
    _GREEK_LC as _READER_GREEK_LC,
)
from .mtef_reader import (
    _GREEK_UC as _READER_GREEK_UC,
)
from .mtef_reader import (
    _SYMBOL_CHARS as _READER_SYMBOLS,
)

# --- MTEF char typefaces ----------------------------------------------------
TF_SYMBOL = 6
TF_MTEXTRA = 11
TF_FUNCTION = 2
TF_VARIABLE = 3
TF_LCGREEK = 4
TF_UCGREEK = 5
TF_NUMBER = 8

# --- Extra symbols beyond the reader's subset -------------------------------
# key: Unicode mtcode, value: LaTeX command. The Symbol typeface table in
# MathType uses code points that coincide with Unicode for common operators.
_EXTRA_SYMBOLS: dict[int, str] = {
    0x2197: '\\nearrow',
    0x2198: '\\searrow',
    0x2199: '\\swarrow',
    0x2196: '\\nwarrow',
    0x2192: '\\to',
    0x2190: '\\gets',
    0x27F6: '\\longrightarrow',
    0x27F5: '\\longleftarrow',
    0x27F9: '\\implies',
    0x27F8: '\\impliedby',
    0x27FA: '\\Longleftrightarrow',
    0x21D4: '\\iff',
    0x226E: '\\nless',
    0x226F: '\\ngtr',
    0x2272: '\\lesssim',
    0x2273: '\\gtrsim',
    0x2AAF: '\\preceq',
    0x2AB0: '\\succeq',
    0x003C: '\\lt',
    0x003E: '\\gt',
    0x2217: '\\ast',
    0x2605: '\\bigstar',
    0x2020: '\\dagger',
    0x2021: '\\ddagger',
    0x2A3F: '\\amalg',
    0x2240: '\\wr',
    0x2299: '\\odot',
    0x2291: '\\sqsubseteq',
    0x2292: '\\sqsupseteq',
    0x2205: '\\varnothing',
    0x2227: '\\land',
    0x2228: '\\lor',
    0x00AC: '\\lnot',
    0x22A4: '\\top',
    0x22A5: '\\bot',
    0x22A2: '\\vdash',
    0x22A8: '\\models',
    0x2224: '\\nmid',
    0x222F: '\\oiint',
    0x2232: '\\ointclockwise',
    0x210F: '\\hbar',
    0x00F0: '\\eth',
    0x0131: '\\imath',
    0x0237: '\\jmath',
    0x2113: '\\ell',
    0x2118: '\\wp',
    0x211C: '\\Re',
    0x2111: '\\Im',
    0x2135: '\\aleph',
    0x2136: '\\beth',
    0x2221: '\\measuredangle',
    0x2222: '\\sphericalangle',
    0x25A1: '\\square',
    0x25A0: '\\blacksquare',
    0x25CA: '\\lozenge',
    0x2127: '\\mho',
}

# LaTeX commands that resolve to ordinary printable ASCII and must NOT be
# mapped to the Symbol typeface (they render as plain variable characters).
_ASCII_ALIASES = {'*', '-', '/', '\\lt', '\\gt'}

_LATEX_SYMBOL: dict[str, int] = {}
for _code, _latex in _READER_SYMBOLS.items():
    if _latex in _ASCII_ALIASES:
        continue
    if _latex.isascii() and not _latex.startswith('\\'):
        continue
    _LATEX_SYMBOL[_latex] = _code
for _code, _latex in _EXTRA_SYMBOLS.items():
    _LATEX_SYMBOL.setdefault(_latex, _code)

_LATEX_LCGREEK = {latex: code for code, latex in _READER_GREEK_LC.items()}
_LATEX_UCGREEK = {latex: code for code, latex in _READER_GREEK_UC.items()}


def _build_lookup() -> dict[str, tuple[int, int]]:
    """Build the ``latex -> (mtcode, typeface)`` lookup table."""
    lookup: dict[str, tuple[int, int]] = {}
    # Ordinary printable ASCII: variable (letters) / number (digits).
    for code in range(0x20, 0x7F):
        ch = chr(code)
        lookup[ch] = (code, TF_NUMBER) if ch.isdigit() else (code, TF_VARIABLE)
    # Greek letters.
    for latex, code in _LATEX_LCGREEK.items():
        lookup[latex] = (code, TF_LCGREEK)
    for latex, code in _LATEX_UCGREEK.items():
        lookup[latex] = (code, TF_UCGREEK)
    # Math symbols (Symbol typeface).
    for latex, code in _LATEX_SYMBOL.items():
        lookup[latex] = (code, TF_SYMBOL)
    return lookup


TYPEFACE_LOOKUP: dict[str, tuple[int, int]] = _build_lookup()


def latex_to_char(latex: str) -> tuple[int, int] | None:
    """Map a LaTeX command or character to ``(mtcode, typeface)``.

    Returns ``None`` when the symbol is not in the table (the writer falls
    back to rendering it as plain text).
    """
    return TYPEFACE_LOOKUP.get(latex)
