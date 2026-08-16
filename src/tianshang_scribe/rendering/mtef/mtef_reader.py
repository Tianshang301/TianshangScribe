"""MTEF (MathType Equation Format) binary reader -> LaTeX.

Parses the MTEF payload stored in a MathType OLE object's ``Equation Native``
stream and emits LaTeX, which the existing ``latex_to_omml`` engine can then
convert to native OMML.

The implementation is a Python port of the Apache-2.0 ``mtef-go`` /
``MTEF-py-FIX`` projects, covering the common MTEF record set: lines, chars,
templates (fractions, radicals, n-ary ops, fences, sub/superscripts, arrows),
piles (multi-line stacks) and matrices. Unrecognised records are skipped
tolerantly instead of aborting.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from typing import Any

# --- Record type constants (MTEF v5) ----------------------------------------
_END = 0
_LINE = 1
_CHAR = 2
_TMPL = 3
_PILE = 4
_MATRIX = 5
_EMBELL = 6
_RULER = 7
_FONT_STYLE_DEF = 8
_SIZE = 9
_FULL = 10
_SUB = 11
_SUB2 = 12
_SYM = 13
_SUBSYM = 14
_COLOR = 15
_COLOR_DEF = 16
_FONT_DEF = 17
_EQN_PREFS = 18
_ENCODING_DEF = 19
_FUTURE = 100

# Option flags
_OPT_NUDGE = 0x08
_OPT_CHAR8 = 0x04
_OPT_CHAR16 = 0x10
_OPT_NOMTCODE = 0x20
_OPT_LINE_NULL = 0x01
_OPT_LINE_RULER = 0x02
_OPT_LINE_LSPACE = 0x04

# Template selectors
_TM_ANGLE = 0
_TM_PAREN = 1
_TM_BRACE = 2
_TM_BRACK = 3
_TM_BAR = 4
_TM_DBAR = 5
_TM_FLOOR = 6
_TM_CEILING = 7
_TM_INTERVAL = 9
_TM_ROOT = 10
_TM_FRACT = 11
_TM_UBAR = 12
_TM_OBAR = 13
_TM_ARROW = 14
_TM_INTEG = 15
_TM_SUM = 16
_TM_PROD = 17
_TM_COPROD = 18
_TM_UNION = 19
_TM_INTER = 20
_TM_INTOP = 21
_TM_SUMOP = 22
_TM_LIM = 23
_TM_HBRACE = 24
_TM_HBRACK = 25
_TM_LDIV = 26
_TM_SUB = 27
_TM_SUP = 28
_TM_SUBSUP = 29
_TM_DIRAC = 30
_TM_VEC = 31
_TM_TILDE = 32
_TM_HAT = 33
_TM_ARC = 34

# Embellishment types -> LaTeX
_EMBELL_MAP: dict[int, str] = {
    2: '\\dot',
    3: '\\ddot',
    4: '\\dddot',
    5: "'",
    6: "''",
    18: "'''",
    8: '\\tilde',
    9: '\\hat',
    10: '\\neg',
    11: '\\overrightarrow',
    12: '\\overleftarrow',
    16: '\\overline',
    17: '\\bar',
}

# Char typefaces that select the math-mode character table.
_FN_MTEXTRA = 11
_FN_SYMBOL = 6
_FN_FUNCTION = 2
_FN_VARIABLE = 3
_FN_NUMBER = 8
_FN_LCGREEK = 4
_FN_UCGREEK = 5


@dataclass
class _Node:
    tag: int
    value: Any = None
    children: list[_Node] = field(default_factory=list)


class MTEFParseError(ValueError):
    """Raised when the MTEF payload cannot be parsed."""


def _mtcode_to_latex(mtcode: int, typeface: int) -> str:
    """Map a MathType char code to its LaTeX form (best-effort)."""
    # Typeface 6 (Symbol) / 11 (MT Extra) use a math symbol table keyed by
    # the Unicode-ish mtcode; plain typefaces (text, variable, number, greek)
    # map directly.
    if typeface in (_FN_SYMBOL, _FN_MTEXTRA):
        return _SYMBOL_CHARS.get(mtcode, f'{{\\text{{[U+{mtcode:04x}]}}}}')
    # Greek letters (LCGREEK/UCGREEK) use standard LaTeX names.
    if typeface == _FN_LCGREEK:
        return _GREEK_LC.get(mtcode, chr(mtcode))
    if typeface == _FN_UCGREEK:
        return _GREEK_UC.get(mtcode, chr(mtcode))
    # Fallback: any printable codepoint maps to itself.
    if 0x20 <= mtcode <= 0x7E or mtcode >= 0xA0:
        return chr(mtcode)
    return f'{{\\text{{[U+{mtcode:04x}]}}}}'


_GREEK_LC: dict[int, str] = {
    0x03B1: '\\alpha',
    0x03B2: '\\beta',
    0x03B3: '\\gamma',
    0x03B4: '\\delta',
    0x03B5: '\\varepsilon',
    0x03B6: '\\zeta',
    0x03B7: '\\eta',
    0x03B8: '\\theta',
    0x03B9: '\\iota',
    0x03BA: '\\kappa',
    0x03BB: '\\lambda',
    0x03BC: '\\mu',
    0x03BD: '\\nu',
    0x03BE: '\\xi',
    0x03C0: '\\pi',
    0x03C1: '\\rho',
    0x03C2: '\\varsigma',
    0x03C3: '\\sigma',
    0x03C4: '\\tau',
    0x03C5: '\\upsilon',
    0x03C6: '\\varphi',
    0x03C7: '\\chi',
    0x03C8: '\\psi',
    0x03C9: '\\omega',
    0x03D5: '\\phi',
    0x03D6: '\\varpi',
    0x03D1: '\\vartheta',
}

_GREEK_UC: dict[int, str] = {
    0x0391: '\\Alpha',
    0x0392: '\\Beta',
    0x0393: '\\Gamma',
    0x0394: '\\Delta',
    0x0395: '\\Epsilon',
    0x0396: '\\Zeta',
    0x0397: '\\Eta',
    0x0398: '\\Theta',
    0x0399: '\\Iota',
    0x039A: '\\Kappa',
    0x039B: '\\Lambda',
    0x039C: '\\Mu',
    0x039D: '\\Nu',
    0x039E: '\\Xi',
    0x03A0: '\\Pi',
    0x03A1: '\\Rho',
    0x03A3: '\\Sigma',
    0x03A4: '\\Tau',
    0x03A5: '\\Upsilon',
    0x03A6: '\\Phi',
    0x03A7: '\\Chi',
    0x03A8: '\\Psi',
    0x03A9: '\\Omega',
}

# Common symbol-table characters (subset of MTEF's full map).
_SYMBOL_CHARS: dict[int, str] = {
    0x2211: '\\sum',
    0x220F: '\\prod',
    0x2210: '\\coprod',
    0x222B: '\\int',
    0x222C: '\\iint',
    0x222D: '\\iiint',
    0x222E: '\\oint',
    0x221E: '\\infty',
    0x00B1: '\\pm',
    0x2213: '\\mp',
    0x00D7: '\\times',
    0x00F7: '\\div',
    0x22C5: '\\cdot',
    0x2217: '*',
    0x22C6: '\\star',
    0x2218: '\\circ',
    0x2022: '\\bullet',
    0x2261: '\\equiv',
    0x2260: '\\neq',
    0x2248: '\\approx',
    0x223C: '\\sim',
    0x2243: '\\simeq',
    0x2245: '\\cong',
    0x221D: '\\propto',
    0x2264: '\\leq',
    0x2265: '\\geq',
    0x226A: '\\ll',
    0x226B: '\\gg',
    0x227A: '\\prec',
    0x227B: '\\succ',
    0x2282: '\\subset',
    0x2283: '\\supset',
    0x2286: '\\subseteq',
    0x2287: '\\supseteq',
    0x2208: '\\in',
    0x2209: '\\notin',
    0x220B: '\\ni',
    0x2200: '\\forall',
    0x2203: '\\exists',
    0x2204: '\\nexists',
    0x2205: '\\emptyset',
    0x2202: '\\partial',
    0x2207: '\\nabla',
    0x2192: '\\rightarrow',
    0x2190: '\\leftarrow',
    0x2194: '\\leftrightarrow',
    0x21D2: '\\Rightarrow',
    0x21D0: '\\Leftarrow',
    0x21D4: '\\Leftrightarrow',
    0x21A6: '\\mapsto',
    0x2191: '\\uparrow',
    0x2193: '\\downarrow',
    0x2220: '\\angle',
    0x25B3: '\\triangle',
    0x27C2: '\\perp',
    0x2225: '\\parallel',
    0x00AC: '\\neg',
    0x2227: '\\wedge',
    0x2228: '\\vee',
    0x2229: '\\cap',
    0x222A: '\\cup',
    0x2295: '\\oplus',
    0x2296: '\\ominus',
    0x2297: '\\otimes',
    0x2298: '\\oslash',
    0x22C0: '\\bigwedge',
    0x22C1: '\\bigvee',
    0x22C2: '\\bigcap',
    0x22C3: '\\bigcup',
    0x22EF: '\\cdots',
    0x22EE: '\\vdots',
    0x22F1: '\\ddots',
    0x2026: '\\ldots',
    0x2032: '\\prime',
    0x2033: "\\prime''",
    0x2308: '\\lceil',
    0x2309: '\\rceil',
    0x230A: '\\lfloor',
    0x230B: '\\rfloor',
    0x2329: '\\langle',
    0x232A: '\\rangle',
    0x2212: '-',
    0x2215: '/',
    0x2044: '/',
}


class _Reader:
    """Byte reader over the MTEF payload."""

    def __init__(self, data: bytes) -> None:
        self.data = data
        self.pos = 0

    def read(self, n: int) -> bytes:
        if self.pos + n > len(self.data):
            raise MTEFParseError('MTEF payload truncated')
        chunk = self.data[self.pos : self.pos + n]
        self.pos += n
        return chunk

    def u8(self) -> int:
        return self.read(1)[0]

    def u16(self) -> int:
        return int(struct.unpack('<H', self.read(2))[0])

    def i16(self) -> int:
        return int(struct.unpack('<h', self.read(2))[0])

    def null_string(self) -> str:
        buf = bytearray()
        while True:
            b = self.read(1)[0]
            if b == 0:
                break
            buf.append(b)
        return buf.decode('ascii', errors='replace')


def _skip_future(reader: _Reader) -> None:
    """Skip a FUTURE record (>=100): a length byte then that many bytes."""
    length = reader.u8()
    reader.read(length)


def _read_nudge(reader: _Reader) -> None:
    b1 = reader.u16()
    b2 = reader.u16()
    if b1 == 128 or b2 == 128:
        reader.i16()
        reader.i16()


def _read_header(reader: _Reader) -> None:
    reader.read(5)  # mMtefVer, mPlatform, mProduct, mVersion, mVersionSub
    reader.null_string()  # application name
    reader.read(1)  # mInline


def _read_line(reader: _Reader, node: _Node) -> None:
    options = reader.u8()
    if options & _OPT_NUDGE:
        _read_nudge(reader)
    if options & _OPT_LINE_LSPACE:
        reader.read(1)
    if options & _OPT_LINE_RULER:
        n_stops = reader.u8()
        for _ in range(n_stops):
            reader.read(1)
            reader.read(2)
    node.value = {'null': bool(options & _OPT_LINE_NULL)}


def _read_char(reader: _Reader, node: _Node) -> None:
    options = reader.u8()
    if options & _OPT_NUDGE:
        _read_nudge(reader)
    typeface = reader.u8()
    mtcode = 0
    if not (options & _OPT_NOMTCODE):
        mtcode = reader.u16()
    if options & _OPT_CHAR8:
        reader.u8()
    elif options & _OPT_CHAR16:
        reader.u16()
    node.value = {'mtcode': mtcode, 'typeface': typeface}


def _read_tmpl(reader: _Reader, node: _Node) -> None:
    options = reader.u8()
    if options & _OPT_NUDGE:
        _read_nudge(reader)
    selector = reader.u8()
    b1 = reader.u8()
    variation = b1 & 127 | reader.u8() << 8 if b1 & 128 else b1
    reader.u8()  # template options
    node.value = {'selector': selector, 'variation': variation}


def _read_pile(reader: _Reader, node: _Node) -> None:
    options = reader.u8()
    if options & _OPT_NUDGE:
        _read_nudge(reader)
    halign = reader.u8()
    valign = reader.u8()
    node.value = {'halign': halign, 'valign': valign}


def _read_matrix(reader: _Reader, node: _Node) -> None:
    options = reader.u8()
    if options & _OPT_NUDGE:
        _read_nudge(reader)
    valign = reader.u8()
    h_just = reader.u8()
    reader.u8()
    rows = reader.u8()
    cols = reader.u8()
    node.value = {'rows': rows, 'cols': cols, 'h_just': h_just, 'valign': valign}


def _read_embell(reader: _Reader, node: _Node) -> None:
    options = reader.u8()
    if options & _OPT_NUDGE:
        _read_nudge(reader)
    node.value = {'embell_type': reader.u8()}


def parse_mtef(data: bytes) -> _Node:
    """Parse raw MTEF bytes into an AST and return the root node."""
    reader = _Reader(data)
    _read_header(reader)

    nodes: list[_Node] = []
    while reader.pos < len(data):
        record = reader.u8()
        if record >= _FUTURE:
            _skip_future(reader)
            continue
        if record == _END:
            nodes.append(_Node(_END))
        elif record == _LINE:
            node = _Node(_LINE)
            _read_line(reader, node)
            nodes.append(node)
        elif record == _CHAR:
            node = _Node(_CHAR)
            _read_char(reader, node)
            nodes.append(node)
        elif record == _TMPL:
            node = _Node(_TMPL)
            _read_tmpl(reader, node)
            nodes.append(node)
        elif record == _PILE:
            node = _Node(_PILE)
            _read_pile(reader, node)
            nodes.append(node)
        elif record == _MATRIX:
            node = _Node(_MATRIX)
            _read_matrix(reader, node)
            nodes.append(node)
            nodes.append(_Node(_LINE, {'null': True}))
            nodes.append(_Node(_LINE, {'null': True}))
        elif record == _EMBELL:
            node = _Node(_EMBELL)
            _read_embell(reader, node)
            nodes.append(node)
        elif record == _SIZE:
            reader.read(2)
        elif record in (_FONT_STYLE_DEF, _FONT_DEF):
            reader.u8()
            reader.null_string()
        elif record == _COLOR:
            reader.read(1)
        elif record == _COLOR_DEF:
            options = reader.u8()
            if options & 0x01:  # CMYK
                reader.read(8)
            else:
                reader.read(6)
            if options & 0x04:  # named colour
                reader.null_string()
        elif record == _EQN_PREFS:
            _skip_eqn_prefs(reader)
        elif record == _ENCODING_DEF:
            reader.null_string()
        elif record in (_SUB, _SUB2, _SYM, _SUBSYM, _RULER, _FULL):
            nodes.append(_Node(record))
        else:
            # Unknown record: skip one byte and continue tolerantly.
            reader.read(1)

    return _build_ast(nodes)


def _skip_eqn_prefs(reader: _Reader) -> None:
    reader.u8()  # options
    for _ in range(3):  # sizes, spaces, styles
        size = reader.u8()
        if size > 0:
            reader.read(size)


def _build_ast(nodes: list[_Node]) -> _Node:
    """Assemble the flat record list into a tree using a stack."""
    root = _Node(0xFF)
    stack: list[_Node] = [root]

    for node in nodes:
        if node.tag == _END:
            if len(stack) > 1:
                stack.pop()
        elif node.tag in (_LINE, _TMPL, _PILE, _MATRIX, _EMBELL):
            if not node.value or not node.value.get('null'):
                stack[-1].children.append(node)
                stack.append(node)
        elif node.tag in (_SUB, _SUB2, _SYM, _SUBSYM, _RULER, _FULL):
            stack[-1].children.append(node)
        else:
            stack[-1].children.append(node)

    return root


def mtef_to_latex(data: bytes) -> str:
    """Convert an MTEF binary payload to a LaTeX string.

    Raises ``MTEFParseError`` if the payload cannot be read.
    """
    root = parse_mtef(data)
    parts: list[str] = []
    for child in root.children:
        parts.append(_render(child))
    latex = ''.join(parts).strip()
    if not latex:
        raise MTEFParseError('MTEF payload contains no renderable content')
    return latex


def _render(node: _Node) -> str:
    if node.tag == _LINE:
        return ''.join(_render(c) for c in node.children)
    if node.tag == _CHAR:
        return _mtcode_to_latex(node.value['mtcode'], node.value['typeface'])
    if node.tag == _EMBELL:
        embell = node.value['embell_type']
        if embell in _EMBELL_MAP:
            return _EMBELL_MAP[embell] + '{}'
        return ''
    if node.tag == _TMPL:
        return _render_tmpl(node)
    if node.tag == _PILE:
        return _render_pile(node)
    if node.tag == _MATRIX:
        return _render_matrix(node)
    return ''.join(_render(c) for c in node.children)


def _render_tmpl(node: _Node) -> str:
    value = node.value
    selector = value['selector']
    variation = value['variation']

    def child(idx: int) -> _Node:
        return node.children[idx] if idx < len(node.children) else _Node(_LINE)

    def render(idx: int) -> str:
        return _render(child(idx))

    if selector == _TM_FRACT:
        num = render(0)
        den = render(1)
        return f'\\frac{{{num}}}{{{den}}}'
    if selector == _TM_ROOT:
        if variation == 0:
            return f'\\sqrt{{{render(0)}}}'
        return f'\\sqrt[{render(1)}]{{{render(0)}}}'
    if selector in (
        _TM_SUM,
        _TM_PROD,
        _TM_COPROD,
        _TM_UNION,
        _TM_INTER,
        _TM_INTEG,
        _TM_INTOP,
        _TM_SUMOP,
    ):
        return _render_bigop(selector, variation, node)
    if selector in (_TM_SUB, _TM_SUP, _TM_SUBSUP):
        return _render_scripts(selector, node)
    if selector in (
        _TM_ANGLE,
        _TM_PAREN,
        _TM_BRACE,
        _TM_BRACK,
        _TM_BAR,
        _TM_DBAR,
        _TM_FLOOR,
        _TM_CEILING,
        _TM_INTERVAL,
    ):
        return _render_fences(selector, node)
    if selector == _TM_UBAR:
        return f'\\underline{{{render(0)}}}'
    if selector == _TM_OBAR:
        return f'\\overline{{{render(0)}}}'
    if selector in (_TM_VEC, _TM_HAT, _TM_TILDE, _TM_ARC):
        return _render_hat(selector, node)
    if selector == _TM_LIM:
        main = render(0)
        lower = render(1) if len(node.children) > 1 else ''
        if lower:
            return f'\\lim_{{{lower}}} {main}'
        return f'\\lim {main}'
    if selector == _TM_ARROW:
        return _render_arrow(variation, node)
    return ''.join(render(i) for i in range(len(node.children)))


def _render_bigop(selector: int, variation: int, node: _Node) -> str:
    op_names = {
        _TM_SUM: '\\sum',
        _TM_PROD: '\\prod',
        _TM_COPROD: '\\coprod',
        _TM_UNION: '\\bigcup',
        _TM_INTER: '\\bigcap',
        _TM_INTEG: {1: '\\int', 2: '\\iint', 3: '\\iiint'}.get(variation, '\\int'),
        _TM_INTOP: '\\int',
        _TM_SUMOP: '\\sum',
    }
    if selector == _TM_INTEG and isinstance(op_names[_TM_INTEG], dict):
        op = op_names[_TM_INTEG]
    else:
        op = op_names.get(selector, '\\sum')

    main = _render(node.children[0]) if node.children else ''
    lower = _render(node.children[1]) if len(node.children) > 1 else ''
    upper = _render(node.children[2]) if len(node.children) > 2 else ''
    sup_str = f'^{{{upper}}}' if upper else ''
    sub_str = f'_{{{lower}}}' if lower else ''
    return f'{op}{sub_str}{sup_str} {main}'


def _render_scripts(selector: int, node: _Node) -> str:
    base = _render(node.children[0]) if node.children else ''
    if selector == _TM_SUB:
        sub = _render(node.children[1]) if len(node.children) > 1 else ''
        return f'{base}_{{{sub}}}'
    if selector == _TM_SUP:
        sup = _render(node.children[1]) if len(node.children) > 1 else ''
        return f'{base}^{{{sup}}}'
    sub = _render(node.children[1]) if len(node.children) > 1 else ''
    sup = _render(node.children[2]) if len(node.children) > 2 else ''
    sub_str = f'_{{{sub}}}' if sub else ''
    sup_str = f'^{{{sup}}}' if sup else ''
    return f'{base}{sup_str}{sub_str}'


def _render_fences(selector: int, node: _Node) -> str:
    fence_map = {
        _TM_ANGLE: ('\\langle', '\\rangle'),
        _TM_PAREN: ('(', ')'),
        _TM_BRACE: ('\\{', '\\}'),
        _TM_BRACK: ('[', ']'),
        _TM_BAR: ('|', '|'),
        _TM_DBAR: ('\\|', '\\|'),
        _TM_FLOOR: ('\\lfloor', '\\rfloor'),
        _TM_CEILING: ('\\lceil', '\\rceil'),
    }
    left, right = fence_map.get(selector, ('(', ')'))
    main = _render(node.children[0]) if node.children else ''
    return f'\\left{left} {main} \\right{right}'


def _render_hat(selector: int, node: _Node) -> str:
    cmd_map = {_TM_VEC: 'vec', _TM_HAT: 'hat', _TM_TILDE: 'tilde', _TM_ARC: 'wideparen'}
    cmd = cmd_map.get(selector, 'hat')
    main = _render(node.children[0]) if node.children else ''
    return f'\\{cmd}{{{main}}}'


def _render_arrow(variation: int, node: _Node) -> str:
    top = _render(node.children[0]) if node.children else ''
    bottom = _render(node.children[1]) if len(node.children) > 1 else ''
    if variation & 0x0001:
        arrow = '\\xrightarrow' if variation & 0x0020 else '\\xleftarrow'
    elif variation & 0x0010:
        arrow = '\\xleftarrow'
    else:
        arrow = '\\xrightarrow'
    if bottom and top:
        return f'{arrow}[{bottom}]{{{top}}}'
    if bottom:
        return f'{arrow}[{bottom}]{{}}'
    return f'{arrow}{{{top}}}'


def _render_pile(node: _Node) -> str:
    lines = [_render(c) for c in node.children]
    if len(lines) <= 1:
        return ''.join(lines)
    return ' \\begin{{gathered}} {} \\end{{gathered}} '.format(' \\\\ '.join(lines))


def _render_matrix(node: _Node) -> str:
    value = node.value
    rows = value['rows']
    cols = value['cols']
    cells = [_render(c) for c in node.children]
    # Skip the leading placeholder line nodes added during parsing.
    if len(cells) >= 2 and cells[0] == '' and cells[1] == '':
        cells = cells[2:]
    if not rows or not cols:
        return ''.join(cells)
    lines: list[str] = []
    for r in range(rows):
        row_cells = cells[r * cols : (r + 1) * cols]
        lines.append(' & '.join(row_cells))
    return ' \\begin{{matrix}} {} \\end{{matrix}} '.format(' \\\\ '.join(lines))
