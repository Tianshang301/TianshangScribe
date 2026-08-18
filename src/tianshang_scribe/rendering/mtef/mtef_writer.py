"""LaTeX -> MTEF binary writer.

Converts LaTeX into MathType Equation Format (MTEF v5) binary records that can
be embedded in a Word document as a MathType OLE object.  The writer consumes
the same nested ``Token`` stream produced by ``math_omml.parse_expression`` so
the syntax handling stays consistent with the OMML path.

The record encoding mirrors what ``mtef_reader.parse_mtef`` decodes:

* LINE  (1):  ``[1, options]``
* CHAR  (2):  ``[2, options, typeface, u16 mtcode]``
* TMPL  (3):  ``[3, options, selector, variation, template_options]``
* END   (0):  ``[0]``

Template argument slots are the records that follow the TMPL record in order,
terminated by END records (one per slot plus a final one closing the
template), matching the reader's stack-based AST builder.
"""

from __future__ import annotations

import struct
from typing import Any

from tianshang_scribe.rendering.math_omml import (
    AccentToken,
    DelimToken,
    FracToken,
    NaryToken,
    OperatorToken,
    ParserContext,
    SqrtToken,
    StyledToken,
    SubSupToken,
    TextToken,
    parse_expression,
)
from tianshang_scribe.rendering.mtef.symbols import latex_to_char

# Record types (MTEF v5)
_END = 0
_LINE = 1
_CHAR = 2
_TMPL = 3

# Template selectors
_TM_PAREN = 1
_TM_BRACE = 2
_TM_BRACK = 3
_TM_ROOT = 10
_TM_FRACT = 11
_TM_UBAR = 12
_TM_OBAR = 13
_TM_ARROW = 14
_TM_INTEG = 15
_TM_SUM = 16
_TM_PROD = 17
_TM_INTOP = 21
_TM_SUB = 27
_TM_SUP = 28
_TM_SUBSUP = 29
_TM_VEC = 31
_TM_TILDE = 32
_TM_HAT = 33

# N-ary operator selector / variation mapping.
_NARY_SELECTOR = {
    'sum': _TM_SUM,
    'prod': _TM_PROD,
    'coprod': 18,
    'bigcup': 19,
    'bigcap': 20,
    'bigvee': 22,
    'bigwedge': 22,
    'int': _TM_INTEG,
    'iint': _TM_INTEG,
    'iiint': _TM_INTEG,
    'oint': _TM_INTEG,
}
_INTEG_VARIATION = {'int': 0, 'iint': 2, 'iiint': 3, 'oint': 0}

# Operator names rendered as literal text.
_OPERATORS = {
    'lim': 'lim',
    'max': 'max',
    'min': 'min',
    'inf': 'inf',
    'det': 'det',
    'Pr': 'Pr',
    'gcd': 'gcd',
    'sin': 'sin',
    'cos': 'cos',
    'tan': 'tan',
    'log': 'log',
    'ln': 'ln',
    'cot': 'cot',
    'sec': 'sec',
    'csc': 'csc',
    'deg': 'deg',
    'dim': 'dim',
    'hom': 'hom',
    'ker': 'ker',
    'arg': 'arg',
    'sup': 'sup',
}

# Accent name -> MTEF embellishment type / template selector.
_ACCENT_SELECTOR = {
    'dot': 2,
    'ddot': 3,
    'tilde': _TM_TILDE,
    'hat': _TM_HAT,
    'bar': 17,
    'vec': _TM_VEC,
}


def latex_to_mtef(latex: str) -> bytes:
    """Convert ``latex`` to an MTEF v5 binary payload.

    Returns the raw MTEF records (without any OLE container wrapper).  The
    payload can be embedded directly or wrapped via ``cfb_writer.make_ole``.
    """
    tokens = parse_expression(ParserContext(latex))
    body = bytearray(_header())
    body += _line()
    body += _records_for_tokens(tokens)
    body += _end()
    body += _end()
    return bytes(body)


def _header() -> bytes:
    app = b'MathType\x00'
    return bytes([5, 1, 1, 0, 0]) + app + bytes([0])


def _line() -> bytes:
    return bytes([_LINE, 0])


def _end() -> bytes:
    return bytes([_END])


def _char(mtcode: int, typeface: int) -> bytes:
    return bytes([_CHAR, 0, typeface]) + struct.pack('<H', mtcode)


def _tmpl(selector: int, variation: int = 0) -> bytes:
    return bytes([_TMPL, 0, selector, variation, 0])


def _records_for_tokens(tokens: list[Any]) -> bytes:
    out = bytearray()
    for token in tokens:
        out += _records_for_token(token)
    return bytes(out)


def _records_for_token(token: Any) -> bytes:
    if isinstance(token, TextToken):
        return _char_records(token.text)

    if isinstance(token, OperatorToken):
        return _operator_records(token)

    if isinstance(token, FracToken):
        return _tmpl_records(
            _TM_FRACT, [token.num, token.den]
        )

    if isinstance(token, SqrtToken):
        if token.degree:
            return _tmpl_records(
                _TM_ROOT, [token.content, token.degree], variation=1
            )
        return _tmpl_records(_TM_ROOT, [token.content])

    if isinstance(token, NaryToken):
        return _nary_records(token)

    if isinstance(token, SubSupToken):
        selector = _TM_SUBSUP
        if token.sub and token.sup:
            pass
        elif token.sub:
            selector = _TM_SUB
        elif token.sup:
            selector = _TM_SUP
        else:
            return _base_records(token.base)
        return _script_records(selector, token)

    if isinstance(token, AccentToken):
        return _accent_records(token)

    if isinstance(token, StyledToken):
        return _records_for_tokens(token.content)

    if isinstance(token, DelimToken):
        return _delim_records(token)

    return b''


def _char_records(text: str) -> bytes:
    out = bytearray()
    for ch in text:
        pair = latex_to_char(ch)
        if pair is not None:
            mtcode, typeface = pair
            out += _char(mtcode, typeface)
        else:
            code = ord(ch)
            out += _char(code, 8 if ch.isdigit() else 3)
    return bytes(out)


def _base_records(base: Any) -> bytes:
    if base is None:
        return b''
    if isinstance(base, TextToken):
        return _char_records(base.text)
    return _records_for_token(base)


def _content_records(content: list[Any]) -> bytes:
    return _records_for_tokens(content)


def _slot_records(content: list[Any]) -> bytes:
    """Wrap ``content`` in a LINE record so multi-char slots stay one node.

    The reader's template renderers read fixed child slots (``children[0]``,
    ``children[1]``, ...); wrapping each slot in a LINE guarantees the whole
    slot is one child regardless of how many tokens it contains.
    """
    out = bytearray(_line())
    out += _content_records(content)
    out += _end()
    return bytes(out)


def _operator_records(token: OperatorToken) -> bytes:
    op = token.op
    name = _OPERATORS.get(op, op)
    base = _char_records(name)
    sub_val = token.sub
    sup_val = token.sup
    if not sub_val and not sup_val:
        return base
    if sub_val and sup_val:
        out = bytearray(_tmpl(_TM_SUBSUP))
        out += _line()
        out += base
        out += _end()
        out += _slot_records(sub_val)
        out += _slot_records(sup_val)
        out += _end()
        return bytes(out)
    selector = _TM_SUB if sub_val else _TM_SUP
    out = bytearray(_tmpl(selector))
    out += _line()
    out += base
    out += _end()
    slot = sub_val if sub_val is not None else sup_val
    out += _slot_records(slot if slot is not None else [])
    out += _end()
    return bytes(out)


def _script_records(selector: int, token: SubSupToken) -> bytes:
    """Emit a SUB/SUP/SUBSUP template.

    Slots are base, then optional sub/sup in MTEF order.  SUBSUP uses the
    order base, sub, sup.
    """
    base_records = _base_records(token.base)
    out = bytearray(_tmpl(selector))
    out += _line()
    out += base_records
    out += _end()
    if selector == _TM_SUBSUP:
        if token.sub:
            out += _slot_records(token.sub)
        if token.sup:
            out += _slot_records(token.sup)
    elif selector == _TM_SUB:
        out += _slot_records(token.sub if token.sub is not None else [])
    else:
        out += _slot_records(token.sup if token.sup is not None else [])
    out += _end()
    return bytes(out)


def _nary_records(token: NaryToken) -> bytes:
    op = token.op
    sub_val = token.sub
    sup_val = token.sup
    body_tokens = token.body

    selector = _NARY_SELECTOR.get(op, _TM_SUM)
    variation = _INTEG_VARIATION.get(op, 0)

    out = bytearray(_tmpl(selector, variation))
    # Slot order per the reader's `_render_bigop`: main, lower, upper.
    # The main slot holds the indexed expression (the body).
    out += _line()
    if body_tokens:
        out += _records_for_tokens(body_tokens)
    out += _end()
    if sub_val:
        out += _slot_records(sub_val)
    if sup_val:
        out += _slot_records(sup_val)
    out += _end()
    return bytes(out)


def _accent_records(token: AccentToken) -> bytes:
    accent = token.accent
    content = token.content
    selector = _ACCENT_SELECTOR.get(accent)
    if selector is None:
        return _records_for_tokens(content)
    return _tmpl_records(selector, [content])


def _delim_records(token: DelimToken) -> bytes:
    """Render a fence template.

    The tokenizer emits separate ``left`` and ``right`` delim tokens around a
    body.  We emit an opening fence template for ``left`` and closing fence
    chars for ``right``.  The body text is already in the token stream between
    the delims, so the PAREN template's main slot wraps it via the reader's
    fence renderer which only reads ``children[0]``.
    """
    char = token.char or ''
    side = token.side
    if side == 'right':
        return _fence_close_char(char)
    selector = _delim_selector(char)
    # Opening fence: template with a placeholder body. The token stream order
    # guarantees the real body follows inside the same template context.
    out = bytearray(_tmpl(selector))
    out += _end()
    out += _end()
    return bytes(out)


def _fence_close_char(char: str) -> bytes:
    close_map = {
        ')': '(',
        ']': '[',
        '}': '{',
    }
    open_char = close_map.get(char, char)
    pair = latex_to_char(open_char)
    if pair is not None:
        return _char(pair[0], pair[1])
    return _char(ord(open_char), 3)


def _tmpl_records(
    selector: int, slots: list[list[Any]], variation: int = 0
) -> bytes:
    """Emit a template whose slots are filled from ``slots`` in order."""
    out = bytearray(_tmpl(selector, variation))
    for slot in slots:
        out += _slot_records(slot)
    out += _end()
    return bytes(out)


def _delim_selector(char: str) -> int:
    mapping = {
        '(': _TM_PAREN,
        ')': _TM_PAREN,
        '[': _TM_BRACK,
        ']': _TM_BRACK,
        '{': _TM_BRACE,
        '}': _TM_BRACE,
    }
    return mapping.get(char, _TM_PAREN)
