"""LaTeX to OMML (Office Math Markup Language) converter.

Rewritten as a recursive descent parser with:
- ``ParserContext`` encapsulating latex/pos/end (directive 3)
- ``Token`` frozen dataclasses replacing dict tokens (directive 5)
- ``CMD_HANDLERS`` dispatch table (directive 2)
- ``parse_expression``/``parse_term``/``parse_factor``/``parse_atom`` grammar (directive 4)
- Zero-copy ``_extract_one_arg`` returning ``(start, end, next)`` (directive 1)
- Sub/sup built in ``parse_factor``; ``_group_sup_sub`` removed (directive 7)
- Precompiled regexes at module level (directive 6)
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

# ── Lookup tables (unchanged) ───────────────────────────────────────────

GREEK_MAP: dict[str, str] = {
    'alpha': '\u03b1', 'beta': '\u03b2', 'gamma': '\u03b3',
    'delta': '\u03b4', 'epsilon': '\u03b5', 'varepsilon': '\u03b5',
    'zeta': '\u03b6', 'eta': '\u03b7', 'theta': '\u03b8',
    'vartheta': '\u03d1', 'iota': '\u03b9', 'kappa': '\u03ba',
    'lambda': '\u03bb', 'mu': '\u03bc', 'nu': '\u03bd',
    'xi': '\u03be', 'pi': '\u03c0', 'varpi': '\u03d6',
    'rho': '\u03c1', 'varrho': '\u03f1', 'sigma': '\u03c3',
    'varsigma': '\u03c2', 'tau': '\u03c4', 'upsilon': '\u03c5',
    'phi': '\u03c6', 'varphi': '\u03d5', 'chi': '\u03c7',
    'psi': '\u03c8', 'omega': '\u03c9',
    'Gamma': '\u0393', 'Delta': '\u0394', 'Theta': '\u0398',
    'Lambda': '\u039b', 'Xi': '\u039e', 'Pi': '\u03a0',
    'Sigma': '\u03a3', 'Upsilon': '\u03a5', 'Phi': '\u03a6',
    'Psi': '\u03a8', 'Omega': '\u03a9',
}

SYMBOL_MAP: dict[str, str] = {
    'infty': '\u221e', 'pm': '\u00b1', 'mp': '\u2213',
    'times': '\u00d7', 'div': '\u00f7', 'cdot': '\u00b7',
    'ast': '\u2217', 'star': '\u22c6', 'circ': '\u2218',
    'bullet': '\u2022', 'equiv': '\u2261', 'neq': '\u2260',
    'approx': '\u2248', 'sim': '\u223c', 'simeq': '\u2243',
    'cong': '\u2245', 'propto': '\u221d', 'leq': '\u2264',
    'geq': '\u2265', 'll': '\u226a', 'gg': '\u226b',
    'prec': '\u227a', 'succ': '\u227b', 'preceq': '\u2aaf',
    'succeq': '\u2ab0', 'subset': '\u2282', 'supset': '\u2283',
    'subseteq': '\u2286', 'supseteq': '\u2287',
    'in': '\u2208', 'notin': '\u2209', 'ni': '\u220b',
    'forall': '\u2200', 'exists': '\u2203', 'nexists': '\u2204',
    'emptyset': '\u2205', 'varnothing': '\u2205',
    'partial': '\u2202', 'nabla': '\u2207',
    'to': '\u2192', 'rightarrow': '\u2192', 'Rightarrow': '\u21d2',
    'leftarrow': '\u2190', 'Leftarrow': '\u21d0',
    'leftrightarrow': '\u2194', 'Leftrightarrow': '\u21d4',
    'mapsto': '\u21a6', 'uparrow': '\u2191', 'downarrow': '\u2193',
    'angle': '\u2220', 'triangle': '\u25b3', 'perp': '\u27c2',
    'parallel': '\u2225', 'lnot': '\u00ac', 'neg': '\u00ac',
    'land': '\u2227', 'lor': '\u2228', 'cap': '\u2229',
    'cup': '\u222a', 'oplus': '\u2295', 'ominus': '\u2296',
    'otimes': '\u2297', 'oslash': '\u2298', 'odot': '\u2299',
}

ACCENT_MAP: dict[str, str] = {
    'hat': '\u0302', 'widehat': '\u0302', 'bar': '\u0304',
    'overline': '\u0305', 'tilde': '\u0303', 'widetilde': '\u0303',
    'dot': '\u0307', 'ddot': '\u0308', 'vec': '\u20d7',
    'check': '\u030c', 'acute': '\u0301', 'grave': '\u0300',
    'breve': '\u0306',
}

# ── Precompiled regex (directive 6) ─────────────────────────────────────

_CMD_RE = re.compile(r'\\([a-zA-Z]+)')
_MATHTYPE_TEXT_RE = re.compile(
    r'\\(text|mathrm|mathsf|mathtt|mathit)\{([^{}]*)\}'
)

# ── Token dataclasses (directive 5) ─────────────────────────────────────

Token = Any  # forward alias for type annotations


@dataclass(frozen=True)
class TextToken:
    """A literal text run, e.g. ``x`` or a Greek/symbol glyph."""

    text: str
    norm: bool = False


@dataclass(frozen=True)
class FracToken:
    """A fraction with nested numerator/denominator token lists."""

    num: list[Token] = field(default_factory=list)
    den: list[Token] = field(default_factory=list)


@dataclass(frozen=True)
class SqrtToken:
    """A root with nested content and optional degree token lists."""

    content: list[Token] = field(default_factory=list)
    degree: list[Token] | None = None


@dataclass(frozen=True)
class NaryToken:
    """A large operator (sum/int/prod) with optional limits and body."""

    op: str = 'sum'
    sub: list[Token] | None = None
    sup: list[Token] | None = None
    body: list[Token] = field(default_factory=list)


@dataclass(frozen=True)
class OperatorToken:
    """A named operator (sin/log/lim/...) with optional sub/sup limits."""

    op: str = ''
    sub: list[Token] | None = None
    sup: list[Token] | None = None


@dataclass(frozen=True)
class SubSupToken:
    """A base with optional sub/sup script token lists."""

    base: Token | None = None
    sub: list[Token] | None = None
    sup: list[Token] | None = None


@dataclass(frozen=True)
class AccentToken:
    """An accent (hat/bar/tilde/...) applied to nested content."""

    accent: str = 'hat'
    content: list[Token] = field(default_factory=list)


@dataclass(frozen=True)
class StyledToken:
    """A style-wrapped (mathrm/mathbf/...) nested content."""

    style: str = 'normal'
    content: list[Token] = field(default_factory=list)


@dataclass(frozen=True)
class DelimToken:
    r"""A ``\left``/``\right`` fence marker."""

    char: str | None = None
    side: str = 'left'


# ── ParserContext (directive 3) ─────────────────────────────────────────


@dataclass
class ParserContext:
    """Mutable parser state shared by the recursive-descent functions."""

    latex: str = ''
    pos: int = 0
    end: int | None = None
    single_char_scripts: bool = False

    def __post_init__(self) -> None:
        """Default *end* to the length of *latex* when not supplied."""
        if self.end is None:
            self.end = len(self.latex)

    def at_end(self) -> bool:
        """Return True when the cursor has reached the parse boundary."""
        return self.pos >= self.end  # type: ignore[operator]

    def peek(self) -> str:
        """Return the current char, or a NUL sentinel at end of input."""
        if self.pos < self.end:  # type: ignore[operator]
            return self.latex[self.pos]
        return '\0'


# ── Zero-copy extraction helpers (directive 1) ──────────────────────────


def _extract_one_arg(
    s: str, start: int
) -> tuple[int, int, int] | None:
    """Return ``(content_start, content_end, next_pos)`` for a ``{…}`` group."""
    if start >= len(s) or s[start] != '{':
        return None

    depth = 0
    for i in range(start, len(s)):
        if s[i] == '{':
            depth += 1
        elif s[i] == '}':
            depth -= 1
            if depth == 0:
                return (start + 1, i, i + 1)

    # unbalanced → consume to end
    return (start + 1, len(s), len(s))


def _extract_two_args(
    s: str, start: int
) -> tuple[tuple[int, int, int], tuple[int, int, int]] | None:
    first = _extract_one_arg(s, start)
    if first is None:
        return None
    second = _extract_one_arg(s, first[2])
    if second is None:
        return None
    return first, second


def _extract_sqrt_args(
    s: str, start: int
) -> tuple[tuple[int, int], tuple[int, int, int]] | None:
    """Return ``((deg_start, deg_end), (content_start, content_end, next))``."""
    if start >= len(s):
        return None

    if s[start] == '[':
        try:
            bracket_end = s.index(']', start) + 1
        except ValueError:
            return None
        content = _extract_one_arg(s, bracket_end)
        if content:
            return ((start + 1, bracket_end - 1), content)
        return None

    content = _extract_one_arg(s, start)
    if content:
        # empty degree sentinel
        return ((start, start), content)
    return None


def _extract_delim(
    s: str, start: int
) -> tuple[str | None, int] | None:
    if start >= len(s):
        return None

    delim_map: dict[str, str | None] = {
        '(': '\u0028', ')': '\u0029',
        '[': '\u005b', ']': '\u005d',
        '{': '\u007b', '}': '\u007d',
        '|': '\u007c', '\\|': '\u007c',
        '.': None,
        '\\{': '\u007b', '\\}': '\u007d',
        '\\langle': '\u2329', '\\rangle': '\u232a',
        '\\lceil': '\u2308', '\\rceil': '\u2309',
        '\\lfloor': '\u230a', '\\rfloor': '\u230b',
    }

    ch = s[start]
    if ch == '\\':
        m = _CMD_RE.match(s, start)
        if m:
            key = '\\' + m.group(1)
            if key in delim_map:
                return (delim_map[key], m.end())
        two_char = s[start : start + 2]
        if two_char in ('\\{', '\\}', '\\|'):
            return (delim_map[two_char], start + 2)
    elif ch in delim_map:
        return (delim_map[ch], start + 1)

    return None


# ── Script argument parser ──────────────────────────────────────────────


def _parse_script_arg(
    latex: str, pos: int, end: int, single_char: bool = False
) -> tuple[bool, list[Token], int]:
    """Parse ``^{...}`` / ``_a`` at *pos* (pointing at ``^``/``_``).

    Returns ``(exists, content_tokens, next_pos)``.
    """
    pos += 1
    if pos >= end:
        return (False, [], pos)

    if latex[pos] == '{':
        span = _extract_one_arg(latex, pos)
        if span is not None:
            cs, ce, nx = span
            return (True, parse_expression(ParserContext(latex, cs, ce)), nx)
        return (False, [], pos)

    if single_char:
        if latex[pos].isalnum():
            ch = latex[pos]
            return (True, [TextToken(ch, ch.isdigit())], pos + 1)
        return (False, [], pos)

    # run mode
    end2 = pos
    while end2 < end and latex[end2].isalnum():
        end2 += 1
    tokens = parse_expression(ParserContext(latex, pos, end2))
    return (True, tokens, end2)


# ── Limits extractor ────────────────────────────────────────────────────


def _extract_limits(
    latex: str, start: int
) -> tuple[list[Token] | None, list[Token] | None, int] | None:
    sub: list[Token] | None = None
    sup: list[Token] | None = None
    pos = start

    if pos < len(latex) and latex[pos] == '_':
        exists, content, pos = _parse_script_arg(latex, pos, len(latex))
        if exists:
            sub = content

    if pos < len(latex) and latex[pos] == '^':
        exists, content, pos = _parse_script_arg(latex, pos, len(latex))
        if exists:
            sup = content

    if sub or sup:
        return sub, sup, pos
    return None


# ── Precompiled maps for renderers ──────────────────────────────────────

_STY_VAL_MAP: dict[str, str] = {
    'normal': 'p', 'roman': 'p', 'bold': 'b', 'italic': 'i',
    'bold-italic': 'bi', 'script': 'scr', 'fraktur': 'fr',
    'double-struck': 'ds', 'sans-serif': 'ss', 'monospace': 'tt',
}

_NARY_OP_MAP: dict[str, str] = {
    'sum': '\u2211', 'prod': '\u220f', 'int': '\u222b',
    'iint': '\u222c', 'iiint': '\u222d', 'oint': '\u222e',
    'coprod': '\u2210', 'bigcup': '\u22c3', 'bigcap': '\u22c2',
    'bigvee': '\u22c1', 'bigwedge': '\u22c0',
}

_ACCENT_OMML_MAP: dict[str, str] = {
    'hat': '\u0302', 'widehat': '\u0302', 'bar': '\u0304',
    'tilde': '\u0303', 'widetilde': '\u0303',
    'dot': '\u0307', 'ddot': '\u0308', 'vec': '\u20d7',
    'check': '\u030c', 'acute': '\u0301', 'grave': '\u0300',
    'breve': '\u0306',
}

_STYLE_CMD_MAP: dict[str, str] = {
    'mathrm': 'normal', 'mathbf': 'bold', 'mathit': 'italic',
    'mathsf': 'sans-serif', 'mathtt': 'monospace',
    'mathbb': 'double-struck', 'text': 'normal', 'mathcal': 'script',
}

_NARY_CMDS = frozenset({
    'sum', 'prod', 'int', 'iint', 'iiint', 'oint', 'coprod',
    'bigcup', 'bigcap', 'bigvee', 'bigwedge',
})

_OPERATOR_LIMIT_CMDS = frozenset({
    'lim', 'max', 'min', 'sup', 'inf', 'det', 'Pr', 'gcd',
    'cot', 'sec', 'csc', 'deg', 'dim', 'hom', 'ker', 'arg',
})

_SIMPLE_OP_CMDS = frozenset({'sin', 'cos', 'tan', 'log', 'ln'})

_ACCENT_CMDS = frozenset(ACCENT_MAP)

_STYLED_CMDS = frozenset(_STYLE_CMD_MAP)


# ── Command handlers ────────────────────────────────────────────────────

_CMD_BREAK_LIST = frozenset({
    'frac', 'sqrt', 'sum', 'int', 'prod', 'lim', 'sin', 'cos',
    'tan', 'log', 'ln', 'det', 'max', 'min', 'sup', 'inf', 'gcd',
    'Pr', 'cot', 'sec', 'csc', 'vec', 'hat', 'bar', 'tilde', 'dot',
    'ddot', 'mathbb', 'mathbf', 'mathit', 'mathrm', 'mathcal', 'left',
    'right', 'iint', 'iiint', 'coprod', 'bigcup', 'bigcap', 'bigvee',
    'bigwedge',
})


def _handle_frac(
    ctx: ParserContext, _cmd: str, cmd_end: int
) -> list[Token] | None:
    args = _extract_two_args(ctx.latex, cmd_end)
    if args:
        (ns, ne, _nx), (ds, de, dx) = args
        num = parse_expression(ParserContext(ctx.latex, ns, ne))
        den = parse_expression(ParserContext(ctx.latex, ds, de))
        ctx.pos = dx
        return [FracToken(num=num, den=den)]
    ctx.pos = cmd_end
    return None


def _handle_sqrt(
    ctx: ParserContext, _cmd: str, cmd_end: int
) -> list[Token] | None:
    args = _extract_sqrt_args(ctx.latex, cmd_end)
    if args:
        (ds, de), (cs, ce, nx) = args
        degree_tokens: list[Token] | None = None
        if ds < de:
            degree_tokens = parse_expression(
                ParserContext(ctx.latex, ds, de)
            )
        content = parse_expression(ParserContext(ctx.latex, cs, ce))
        ctx.pos = nx
        return [SqrtToken(content=content, degree=degree_tokens)]
    ctx.pos = cmd_end
    return None


def _handle_delim(
    ctx: ParserContext, cmd: str, cmd_end: int
) -> list[Token] | None:
    paren = _extract_delim(ctx.latex, cmd_end)
    if paren:
        char_val, nx = paren
        ctx.pos = nx
        return [DelimToken(char=char_val, side=cmd)]
    ctx.pos = cmd_end
    return None


def _handle_nary(
    ctx: ParserContext, cmd: str, cmd_end: int
) -> list[Token] | None:
    limits = _extract_limits(ctx.latex, cmd_end)
    body_start = limits[2] if limits else cmd_end
    body, body_end = _collect_body_tokens(ctx.latex, body_start)
    ctx.pos = body_end
    return [
        NaryToken(
            op=cmd,
            sub=limits[0] if limits else None,
            sup=limits[1] if limits else None,
            body=body,
        )
    ]


def _handle_operator_limits(
    ctx: ParserContext, cmd: str, cmd_end: int
) -> list[Token] | None:
    limits = _extract_limits(ctx.latex, cmd_end)
    ctx.pos = limits[2] if limits else cmd_end
    return [
        OperatorToken(
            op=cmd,
            sub=limits[0] if limits else None,
            sup=limits[1] if limits else None,
        )
    ]


def _handle_simple_operator(
    ctx: ParserContext, cmd: str, cmd_end: int
) -> list[Token]:
    ctx.pos = cmd_end
    return [OperatorToken(op=cmd)]


def _handle_accent(
    ctx: ParserContext, cmd: str, cmd_end: int
) -> list[Token] | None:
    if cmd_end < ctx.end and ctx.latex[cmd_end] == '{':  # type: ignore[operator]
        span = _extract_one_arg(ctx.latex, cmd_end)
        if span is not None:
            cs, ce, nx = span
            content = parse_expression(ParserContext(ctx.latex, cs, ce))
            ctx.pos = nx
            return [AccentToken(accent=cmd, content=content)]
    # no brace group → consume the command and drop it (matches old behavior)
    ctx.pos = cmd_end
    return []


def _handle_mbsp(
    _ctx: ParserContext, _cmd: str, cmd_end: int
) -> list[Token]:
    _ctx.pos = cmd_end
    tokens: list[Token] = [TextToken(' ', False)]
    if _ctx.pos < _ctx.end and _ctx.latex[_ctx.pos] == '{':  # type: ignore[operator]
        span = _extract_one_arg(_ctx.latex, _ctx.pos)
        if span is not None:
            _ctx.pos = span[2]
    return tokens


def _handle_greek(
    ctx: ParserContext, cmd: str, cmd_end: int
) -> list[Token]:
    ctx.pos = cmd_end
    return [TextToken(GREEK_MAP[cmd], cmd[0].isupper())]


def _handle_symbol(
    ctx: ParserContext, cmd: str, cmd_end: int
) -> list[Token]:
    ctx.pos = cmd_end
    return [TextToken(SYMBOL_MAP[cmd])]


def _handle_styled(
    ctx: ParserContext, cmd: str, cmd_end: int
) -> list[Token] | None:
    span = _extract_one_arg(ctx.latex, cmd_end)
    if span is not None:
        cs, ce, nx = span
        content = parse_expression(ParserContext(ctx.latex, cs, ce))
        ctx.pos = nx
        return [
            StyledToken(style=_STYLE_CMD_MAP.get(cmd, 'normal'), content=content)
        ]
    ctx.pos = cmd_end
    return None


# ── Dispatch table (directive 2) ────────────────────────────────────────

CMD_HANDLERS: dict[str, Any] = {}

for _c in _NARY_CMDS:
    CMD_HANDLERS[_c] = _handle_nary
for _c in _OPERATOR_LIMIT_CMDS:
    CMD_HANDLERS[_c] = _handle_operator_limits
for _c in _SIMPLE_OP_CMDS:
    CMD_HANDLERS[_c] = _handle_simple_operator
CMD_HANDLERS['frac'] = _handle_frac
CMD_HANDLERS['sqrt'] = _handle_sqrt
CMD_HANDLERS['left'] = _handle_delim
CMD_HANDLERS['right'] = _handle_delim
CMD_HANDLERS['mbsp'] = _handle_mbsp
for _c in _ACCENT_CMDS:
    CMD_HANDLERS[_c] = _handle_accent
for _c in _STYLE_CMD_MAP:
    CMD_HANDLERS[_c] = _handle_styled
for _c, _ch in GREEK_MAP.items():
    CMD_HANDLERS[_c] = _handle_greek
for _c, _ch in SYMBOL_MAP.items():
    CMD_HANDLERS[_c] = _handle_symbol


# ── Text-run collector (shared by parse_atom & body) ────────────────────

_TOP_LEVEL_EXCLUDE = frozenset(
    ('\\', '^', '_', ' ', '(', ')', '[', ']', '{', '}')
)
_BODY_EXCLUDE = frozenset(('\\', '^', '_', ' '))


def _collect_text_run(
    ctx: ParserContext, exclude: frozenset[str]
) -> list[Token]:
    tokens: list[Token] = []
    while ctx.pos < ctx.end and ctx.latex[ctx.pos] not in exclude:  # type: ignore[operator]
        is_digit = ctx.latex[ctx.pos].isdigit()
        start = ctx.pos
        while (
            ctx.pos < ctx.end  # type: ignore[operator]
            and ctx.latex[ctx.pos] not in exclude
            and ctx.latex[ctx.pos].isdigit() == is_digit
        ):
            ctx.pos += 1
        text = ctx.latex[start : ctx.pos]
        if text:
            tokens.append(TextToken(text, is_digit))
    return tokens


# ── Recursive descent parser (directive 4) ──────────────────────────────


def parse_atom(ctx: ParserContext) -> list[Token]:
    """Parse a single atom: text run, command, or brace/paren literal.

    Returns ``[]`` for spaces (callers must skip transparently).
    """
    while True:
        if ctx.at_end():
            return []

        ch = ctx.latex[ctx.pos]

        # space → not an atom; caller skips
        if ch == ' ':
            return []

        # backslash command
        if ch == '\\':
            m = _CMD_RE.match(ctx.latex, ctx.pos)
            if m:
                cmd = m.group(1)
                cmd_end = m.end()
                handler = CMD_HANDLERS.get(cmd)
                if handler is not None:
                    result = handler(ctx, cmd, cmd_end)
                    if result is not None:
                        return result  # type: ignore[no-any-return]
                    # result is [] → command consumed & dropped; keep scanning
                    if result == []:
                        continue
                # unknown command → text
                ctx.pos = cmd_end
                return [TextToken('\\' + cmd)]
            # bare backslash
            ctx.pos += 1
            return [TextToken('\\')]

        # literal brace/paren
        if ch in ('{', '}', '(', ')', '[', ']'):
            ctx.pos += 1
            return [TextToken(ch)]

        # text run
        return _collect_text_run(ctx, _TOP_LEVEL_EXCLUDE)


def _attach_script(
    tokens: list[Token],
    kind: str,
    content: list[Token],
) -> None:
    """Attach a sub/sup script to the last token in *tokens*."""
    if tokens and isinstance(tokens[-1], SubSupToken):
        last = tokens[-1]
        if kind == 'sub':
            tokens[-1] = SubSupToken(base=last.base, sub=content, sup=last.sup)
        else:
            tokens[-1] = SubSupToken(base=last.base, sub=last.sub, sup=content)
    elif tokens:
        base = tokens.pop()
        if kind == 'sub':
            tokens.append(SubSupToken(base=base, sub=content))
        else:
            tokens.append(SubSupToken(base=base, sup=content))
    else:
        if kind == 'sub':
            tokens.append(SubSupToken(base=None, sub=content))
        else:
            tokens.append(SubSupToken(base=None, sup=content))


def parse_factor(ctx: ParserContext) -> list[Token]:
    """Parse an atom followed by optional ``^``/``_`` scripts."""
    base_tokens = parse_atom(ctx)

    # skip spaces for script lookahead
    while ctx.pos < ctx.end and ctx.latex[ctx.pos] == ' ':  # type: ignore[operator]
        ctx.pos += 1

    saw_script = False
    while ctx.pos < ctx.end and ctx.latex[ctx.pos] in ('^', '_'):  # type: ignore[operator]
        kind = 'sup' if ctx.latex[ctx.pos] == '^' else 'sub'
        ctx.pos += 1
        if ctx.pos >= ctx.end:  # type: ignore[operator]
            break
        if ctx.latex[ctx.pos] == '{':
            span = _extract_one_arg(ctx.latex, ctx.pos)
            if span is not None:
                cs, ce, nx = span
                content = parse_expression(ParserContext(ctx.latex, cs, ce))
                ctx.pos = nx
                _attach_script(base_tokens, kind, content)
                saw_script = True
            else:
                break
        elif ctx.single_char_scripts:
            if ctx.latex[ctx.pos].isalnum():
                ch = ctx.latex[ctx.pos]
                ctx.pos += 1
                _attach_script(base_tokens, kind, [TextToken(ch, ch.isdigit())])
                saw_script = True
            else:
                break
        else:
            end2 = ctx.pos
            while end2 < ctx.end and ctx.latex[end2].isalnum():  # type: ignore[operator]
                end2 += 1
            content = parse_expression(ParserContext(ctx.latex, ctx.pos, end2))
            ctx.pos = end2
            _attach_script(base_tokens, kind, content)
            saw_script = True

        # skip spaces between consecutive scripts
        while ctx.pos < ctx.end and ctx.latex[ctx.pos] == ' ':  # type: ignore[operator]
            ctx.pos += 1

    if not saw_script and not base_tokens:
        return []
    return base_tokens


def parse_term(ctx: ParserContext) -> list[Token]:
    """Parse one or more factors (implicit sequence)."""
    result: list[Token] = []
    while not ctx.at_end():
        # skip spaces transparently
        while ctx.pos < ctx.end and ctx.latex[ctx.pos] == ' ':  # type: ignore[operator]
            ctx.pos += 1
        before = ctx.pos
        factors = parse_factor(ctx)
        if not factors:
            if ctx.pos == before:
                break  # no progress at all
            continue  # consumed input but produced no tokens; keep scanning
        result.extend(factors)
    return result


def parse_expression(ctx: ParserContext) -> list[Token]:
    """Parse a sequence of terms (top-level grammar rule)."""
    result: list[Token] = []
    while not ctx.at_end():
        # skip spaces transparently
        while ctx.pos < ctx.end and ctx.latex[ctx.pos] == ' ':  # type: ignore[operator]
            ctx.pos += 1
        before = ctx.pos
        terms = parse_term(ctx)
        if not terms:
            if ctx.pos == before:
                break  # no progress at all
            continue  # consumed input but produced no tokens; keep scanning
        result.extend(terms)
    return result


# ── Nary body collector ─────────────────────────────────────────────────


def _collect_body_tokens(
    latex: str, start: int
) -> tuple[list[Token], int]:
    r"""Collect tokens for a nary body until the next ``\command`` or end.

    Matches the old ``_collect_body_tokens`` semantics exactly:
    - Text runs split at digit boundaries, exclude ``\^_ `` (not parens/braces)
    - ``^``/``_`` use single-char scripts (not run)
    - Stops at any ``\command``
    """
    tokens: list[Token] = []
    pos = start
    length = len(latex)

    while pos < length:
        ch = latex[pos]
        if ch == '\\':
            break
        elif ch in ('^', '_'):
            kind = 'sup' if ch == '^' else 'sub'
            pos += 1
            if pos < length and latex[pos] == '{':
                span = _extract_one_arg(latex, pos)
                if span is not None:
                    cs, ce, nx = span
                    content = parse_expression(
                        ParserContext(latex, cs, ce)
                    )
                    pos = nx
                    _attach_script(tokens, kind, content)
                else:
                    break
            elif pos < length and latex[pos].isalnum():
                ch2 = latex[pos]
                pos += 1
                _attach_script(
                    tokens, kind, [TextToken(ch2, ch2.isdigit())]
                )
            continue
        elif ch == ' ':
            pos += 1
            continue
        else:
            while pos < length and latex[pos] not in _BODY_EXCLUDE:
                is_digit = latex[pos].isdigit()
                run_start = pos
                while (
                    pos < length
                    and latex[pos] not in _BODY_EXCLUDE
                    and latex[pos].isdigit() == is_digit
                ):
                    pos += 1
                text = latex[run_start:pos]
                if text:
                    tokens.append(TextToken(text, is_digit))

    return tokens, pos


# ── MathType preprocess (unchanged) ────────────────────────────────────


def _mathtype_preprocess(latex: str) -> str:
    r"""Normalize MathType-dialect LaTeX before the standard parser runs."""
    latex = latex.replace('~', '\\mbsp{}')

    def _protect_text_spaces(match: re.Match[str]) -> str:
        cmd = match.group(1)
        body = match.group(2)
        protected = body.replace(' ', '\\mbsp{}')
        return f'\\{cmd}{{{protected}}}'

    latex = _MATHTYPE_TEXT_RE.sub(_protect_text_spaces, latex)
    return latex


# ── Public API ──────────────────────────────────────────────────────────


def latex_to_omml(latex: str, style: str = 'office') -> Any:
    r"""Convert a LaTeX formula string to an OMML ``m:oMath``/``m:oMathPara`` element."""
    latex = latex.strip()
    if not latex:
        return None

    if style == 'mathtype':
        latex = _mathtype_preprocess(latex)

    is_display = latex.startswith(r'\[') or latex.startswith('$$')
    if is_display:
        latex = re.sub(
            r'^(\\\[|\\begin\{equation\*\}|\\begin\{equation\}|\$\$)',
            '',
            latex,
        )
        latex = re.sub(
            r'(\\\]|\\end\{equation\*\}|\\end\{equation\}|\$\$)$',
            '',
            latex,
        )
    else:
        latex = re.sub(r'^\$', '', latex)
        latex = re.sub(r'\$$', '', latex)

    latex = latex.strip()
    if not latex:
        return None

    tokens = parse_expression(ParserContext(latex))
    elements = _render_tokens(tokens)

    if not elements:
        return None

    if is_display:
        para = OxmlElement('m:oMathPara')
        omath = OxmlElement('m:oMath')
        for el in elements:
            omath.append(el)
        para.append(omath)
        return para
    else:
        omath = OxmlElement('m:oMath')
        for el in elements:
            omath.append(el)
        return omath


# ── OMML rendering helpers ──────────────────────────────────────────────


def _make_run(text: str, math_style: str = 'auto', norm: bool = False) -> Any:
    m_run = OxmlElement('m:r')
    rpr = None
    if math_style and math_style != 'auto':
        rpr = _make_run_props(math_style)
    if norm:
        if rpr is None:
            rpr = OxmlElement('m:rPr')
        nor_el = OxmlElement('m:nor')
        nor_el.set(qn('m:val'), '1')
        rpr.insert(0, nor_el)
    if rpr is not None:
        m_run.append(rpr)
    mt = OxmlElement('m:t')
    mt.set(qn('xml:space'), 'preserve')
    mt.text = text
    m_run.append(mt)
    return m_run


def _make_run_props(math_style: str) -> Any:
    rpr = OxmlElement('m:rPr')
    sty = OxmlElement('m:sty')
    sty.set(qn('m:val'), _STY_VAL_MAP.get(math_style, 'p'))
    rpr.append(sty)
    return rpr


def _rpr_equivalent(a: Any, b: Any) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return bool(etree.tostring(a) == etree.tostring(b))


def _inject_style_recursive(el: Any, style: str, norm: bool) -> None:
    if el.tag == qn('m:r'):
        rpr = OxmlElement('m:rPr')
        if norm:
            nor_el = OxmlElement('m:nor')
            nor_el.set(qn('m:val'), '1')
            rpr.append(nor_el)
        sty = OxmlElement('m:sty')
        sty.set(qn('m:val'), _STY_VAL_MAP.get(style, 'p'))
        rpr.append(sty)
        existing_rpr = el.find(qn('m:rPr'))
        if existing_rpr is not None:
            from copy import deepcopy

            for attr in existing_rpr:
                if attr.tag in (qn('m:nor'), qn('m:sty')):
                    continue
                rpr.append(deepcopy(attr))
            el.remove(existing_rpr)
        el.insert(0, rpr)
    else:
        for child in el:
            _inject_style_recursive(child, style, norm)


# ── Sub-expression rendering ────────────────────────────────────────────


def _render_tokens(tokens: list[Token]) -> list[Any]:
    """Render a list of tokens into OMML elements, merging adjacent runs."""
    result_tokens: list[Any] = []
    for item in tokens:
        elem = _token_to_omml(item)
        if elem is not None:
            if isinstance(elem, list):
                result_tokens.extend(elem)
            else:
                result_tokens.append(elem)
    if not result_tokens:
        return []

    elements: list[Any] = []
    pending_texts: list[Any] = []
    pending_rpr = None

    for el in result_tokens:
        tag = el.tag
        if tag == qn('m:r'):
            mt_children = [c for c in el if c.tag == qn('m:t')]
            rpr_child = next(
                (c for c in el if c.tag == qn('m:rPr')), None
            )
            if pending_texts and not _rpr_equivalent(pending_rpr, rpr_child):
                mr = OxmlElement('m:r')
                if pending_rpr is not None:
                    from copy import deepcopy

                    mr.append(deepcopy(pending_rpr))
                for t in pending_texts:
                    mr.append(t)
                elements.append(mr)
                pending_texts = []
                pending_rpr = None
            pending_texts.extend(mt_children)
            if pending_rpr is None and rpr_child is not None:
                pending_rpr = rpr_child
        elif tag == qn('m:t'):
            pending_texts.append(el)
        else:
            if pending_texts:
                mr = OxmlElement('m:r')
                if pending_rpr is not None:
                    from copy import deepcopy

                    mr.append(deepcopy(pending_rpr))
                for t in pending_texts:
                    mr.append(t)
                elements.append(mr)
                pending_texts = []
                pending_rpr = None
            elements.append(el)

    if pending_texts:
        mr = OxmlElement('m:r')
        if pending_rpr is not None:
            from copy import deepcopy

            mr.append(deepcopy(pending_rpr))
        for t in pending_texts:
            mr.append(t)
        elements.append(mr)

    return elements


def _build_sub_omath(content: list[Token]) -> list[Any]:
    """Render token list for a slot; empty → empty run."""
    if not content:
        return [_make_run('')]
    return _render_tokens(content)


def _build_styled_elements(
    content: list[Token], style: str
) -> list[Any]:
    norm = style != 'italic'
    elements = _render_tokens(content)
    for el in elements:
        _inject_style_recursive(el, style, norm)
    return elements


# ── Per-token OMML builder ──────────────────────────────────────────────


def _token_to_omml(token: Token) -> Any:
    if isinstance(token, TextToken):
        return _make_run(token.text, norm=token.norm)

    if isinstance(token, OperatorToken):
        op_elem = OxmlElement(
            'm:sSubSup'
            if (token.sub and token.sup)
            else ('m:sSub' if token.sub else 'm:sSup')
            if (token.sub or token.sup)
            else 'm:r'
        )
        if token.sub or token.sup:
            e = OxmlElement('m:e')
            e.append(_make_run(token.op, 'normal', norm=True))
            op_elem.append(e)
            if token.sub:
                sub_e = OxmlElement('m:sub')
                for el in _build_sub_omath(token.sub):
                    sub_e.append(el)
                op_elem.append(sub_e)
            if token.sup:
                sup_e = OxmlElement('m:sup')
                for el in _build_sub_omath(token.sup):
                    sup_e.append(el)
                op_elem.append(sup_e)
            return op_elem
        return _make_run(token.op, 'normal', norm=True)

    if isinstance(token, FracToken):
        f = OxmlElement('m:f')
        num = OxmlElement('m:num')
        for el in _build_sub_omath(token.num):
            num.append(el)
        f.append(num)
        den = OxmlElement('m:den')
        for el in _build_sub_omath(token.den):
            den.append(el)
        f.append(den)
        return f

    if isinstance(token, SqrtToken):
        rad = OxmlElement('m:rad')
        if token.degree:
            deg = OxmlElement('m:deg')
            for el in _build_sub_omath(token.degree):
                deg.append(el)
            rad.append(deg)
        else:
            rad_pr = OxmlElement('m:radPr')
            deg_hide = OxmlElement('m:degHide')
            deg_hide.set(qn('m:val'), '1')
            rad_pr.append(deg_hide)
            rad.insert(0, rad_pr)
        e = OxmlElement('m:e')
        for el in _build_sub_omath(token.content):
            e.append(el)
        rad.append(e)
        return rad

    if isinstance(token, NaryToken):
        nary = OxmlElement('m:nary')
        nary_pr = OxmlElement('m:naryPr')
        chr_el = OxmlElement('m:chr')
        chr_el.set(qn('m:val'), _NARY_OP_MAP.get(token.op, '\u2211'))
        nary_pr.append(chr_el)
        nary.append(nary_pr)

        if token.sub:
            sub = OxmlElement('m:sub')
            for el in _build_sub_omath(token.sub):
                sub.append(el)
            nary.append(sub)

        if token.sup:
            sup = OxmlElement('m:sup')
            for el in _build_sub_omath(token.sup):
                sup.append(el)
            nary.append(sup)

        if token.body:
            e = OxmlElement('m:e')
            for el in _render_tokens(token.body):
                e.append(el)
            nary.append(e)

        return nary

    if isinstance(token, SubSupToken):
        # replicate old behavior: base renders as text if TextToken, else empty
        has_sub = token.sub is not None
        has_sup = token.sup is not None
        if has_sub and has_sup:
            ss = OxmlElement('m:sSubSup')
        elif has_sub:
            ss = OxmlElement('m:sSub')
        else:
            ss = OxmlElement('m:sSup')

        e = OxmlElement('m:e')
        if isinstance(token.base, TextToken):
            for el in _build_sub_omath([token.base]):
                e.append(el)
        else:
            e.append(_make_run(''))
        ss.append(e)

        if has_sub:
            sub = OxmlElement('m:sub')
            for el in _build_sub_omath(token.sub or []):
                sub.append(el)
            ss.append(sub)
        if has_sup:
            sup = OxmlElement('m:sup')
            for el in _build_sub_omath(token.sup or []):
                sup.append(el)
            ss.append(sup)
        return ss

    if isinstance(token, AccentToken):
        acc = OxmlElement('m:acc')
        acc_pr = OxmlElement('m:accPr')
        chr_el = OxmlElement('m:chr')
        chr_el.set(
            qn('m:val'),
            _ACCENT_OMML_MAP.get(token.accent, '\u0302'),
        )
        acc_pr.append(chr_el)
        acc.append(acc_pr)
        e = OxmlElement('m:e')
        for el in _build_sub_omath(token.content):
            e.append(el)
        acc.append(e)
        return acc

    if isinstance(token, StyledToken):
        return _build_styled_elements(token.content, token.style)

    if isinstance(token, DelimToken):
        return _make_run('')

    return _make_run('')


def _stack_to_element(tokens: list[Token]) -> Any:
    """Merge all tokens into a single ``m:r`` element."""
    elements = _render_tokens(tokens)
    if not elements:
        return _make_run('')
    if len(elements) == 1:
        return elements[0]

    m_run = OxmlElement('m:r')
    merged_rpr = None
    for el in elements:
        tag = el.tag
        if tag == qn('m:t'):
            m_run.append(el)
        elif tag == qn('m:r'):
            for child in el:
                if child.tag == qn('m:rPr'):
                    if merged_rpr is None:
                        merged_rpr = child
                elif child.tag == qn('m:t'):
                    m_run.append(child)
        else:
            return el

    if merged_rpr is not None:
        from copy import deepcopy

        m_run.insert(0, deepcopy(merged_rpr))
    return m_run
