from __future__ import annotations

import re
from typing import Any

from src.rendering.math_omml import ACCENT_MAP, GREEK_MAP, SYMBOL_MAP

_TEXT_SYMBOLS: dict[str, str] = {
    **GREEK_MAP,
    **SYMBOL_MAP,
    'lim': 'lim', 'sin': 'sin', 'cos': 'cos', 'tan': 'tan',
    'log': 'log', 'ln': 'ln', 'det': 'det',
    'max': 'max', 'min': 'min', 'sup': 'sup', 'inf': 'inf',
    'Pr': 'Pr', 'gcd': 'gcd', 'deg': 'deg', 'dim': 'dim',
    'hom': 'hom', 'ker': 'ker',
}

_TEXT_ACCENTS: dict[str, str] = ACCENT_MAP

_MATH_CMD_NAMES = frozenset({
    'frac', 'sqrt', 'sum', 'int', 'prod', 'iint', 'iiint', 'oint',
    'lim', 'sin', 'cos', 'tan', 'log', 'ln', 'det',
    'max', 'min', 'sup', 'inf', 'Pr', 'gcd', 'deg', 'dim',
    'hom', 'ker', 'cot', 'sec', 'csc', 'arg',
    'coprod', 'bigcup', 'bigcap', 'bigvee', 'bigwedge',
    'overline', 'vec', 'dot', 'ddot',
    'hat', 'bar', 'tilde', 'widehat', 'widetilde',
    'check', 'acute', 'grave', 'breve',
})


def _extract_math_expression(text: str, start: int) -> tuple[str, int]:
    i = start
    length = len(text)
    while i < length and text[i].isspace():
        i += 1
    while i < length:
        ch = text[i]
        if ch == '{':
            depth = 0
            j = i
            while j < length:
                if text[j] == '{':
                    depth += 1
                elif text[j] == '}':
                    depth -= 1
                    if depth == 0:
                        i = j + 1
                        break
                j += 1
            else:
                return text[start:i + 1], i + 1
            continue
        elif ch == '[':
            depth = 0
            j = i
            while j < length:
                if text[j] == '[':
                    depth += 1
                elif text[j] == ']':
                    depth -= 1
                    if depth == 0:
                        i = j + 1
                        break
                j += 1
            continue
        elif ch == '_' or ch == '^':
            i += 1
            if i < length and text[i] == '{':
                depth = 0
                j = i
                while j < length:
                    if text[j] == '{':
                        depth += 1
                    elif text[j] == '}':
                        depth -= 1
                        if depth == 0:
                            i = j + 1
                            break
                    j += 1
            elif i < length and text[i].isalnum():
                i += 1
            continue
        else:
            break

    while i < length:
        ch = text[i]
        if ch in ('\\', '$', '\n'):
            break
        if '\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f':
            break
        if ch in ('\uff0c', '\u3002', '\uff0e', '\uff01', '\uff1b', '\uff1a', '\uff1f'):
            break
        i += 1

    return text[start:i], i


def _wrap_math_commands(text: str) -> str:
    pattern = re.compile(
        r'\\(frac|sqrt|sum|int|prod|iint|iiint|oint|'
        r'lim|sin|cos|tan|log|ln|det|max|min|sup|inf|'
        r'Pr|gcd|deg|dim|hom|ker|cot|sec|csc|arg|'
        r'coprod|bigcup|bigcap|bigvee|bigwedge|'
        r'hat|bar|tilde|widehat|widetilde|vec|dot|ddot|'
        r'check|acute|grave|breve|overline)'
        r'(?![a-zA-Z])'
    )

    result_parts: list[str] = []
    last_end = 0

    for m in pattern.finditer(text):
        if m.start() < last_end:
            continue
        result_parts.append(text[last_end:m.start()])
        cmd_name = m.group(1)
        expr, end = _extract_math_expression(text, m.end())
        full_formula = '\\' + cmd_name + expr
        result_parts.append('$' + full_formula + '$')
        last_end = end

    result_parts.append(text[last_end:])
    result = ''.join(result_parts)
    result = _merge_adjacent_math(result)
    return result


def _merge_adjacent_math(text: str) -> str:
    result = text
    for _ in range(5):
        prev = result
        result = re.sub(
            r'\$([^$]+)\$\s*\$([^$]+)\$',
            lambda m: '${} {}$'.format(m.group(1), m.group(2)),
            result,
        )
        if result == prev:
            break
    return result


_SIMPLE_CMDS = (
    r'\\(bfseries|itshape|scshape|rmfamily|sffamily|ttfamily|underline|noindent)'
    r'\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}'
)

_PARAM_CMDS = (
    r'\\(fontfamily|fontsize|color|linespread|centering|raggedright|raggedleft|indent)'
    r'\{([^}]+)\}\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}'
)

_HEADING_CMD = r'\\(heading)\{(\d+)\}\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}'

_NEWPAGE_CMD = r'\\(newpage)'

_INCLUDEGRAPHICS_CMD = r'\\(includegraphics)\{([^}]+)\}'

_INLINE_MATH = r'\$([^$]+)\$'

_DISPLAY_MATH = r'\$\$([^$]+)\$\$'

_FONT_CONFIG_CMD = (
    r'\\(setmainfont|setCJKmainfont|setsansfont|setCJKsansfont'
    r'|setmonofont|setCJKmonofont)\{([^}]+)\}'
)

_TOKEN_PATTERN = re.compile(
    _SIMPLE_CMDS + r'|' + _PARAM_CMDS + r'|' + _HEADING_CMD +
    r'|' + _DISPLAY_MATH + r'|' + _INLINE_MATH +
    r'|' + _NEWPAGE_CMD + r'|' + _INCLUDEGRAPHICS_CMD +
    r'|' + _FONT_CONFIG_CMD
)

STYLE_MAP: dict[str, dict[str, str | bool | int | None]] = {
    'bfseries': {'bold': True},
    'itshape': {'italic': True},
    'rmfamily': {'font_name': 'Times New Roman'},
    'sffamily': {'font_name': 'Arial'},
    'ttfamily': {'font_name': 'Courier New'},
    'underline': {'underline': True},
}


_DOLLAR_ESCAPE = '\u2404'


def _escape_dollar(text: str) -> str:
    return text.replace(r'\$', _DOLLAR_ESCAPE)


def _unescape_dollar(text: str) -> str:
    return text.replace(_DOLLAR_ESCAPE, '$')


def preprocess_text(text: str) -> str:
    text = _escape_dollar(text)

    orig_regions = _find_math_regions(text)
    if orig_regions:
        merged = _merge_regions(orig_regions)
        parts: list[str] = []
        last_end = 0
        for start, end in merged:
            parts.append(_wrap_math_commands(text[last_end:start]))
            parts.append(text[start:end])
            last_end = end
        parts.append(_wrap_math_commands(text[last_end:]))
        text = ''.join(parts)
    else:
        text = _wrap_math_commands(text)

    math_regions = _find_math_regions(text)

    if not math_regions:
        return _unescape_dollar(_apply_preprocessing(text))

    merged = _merge_regions(math_regions)

    wrapped_parts: list[str] = []
    last_end = 0
    for start, end in merged:
        wrapped_parts.append(_apply_preprocessing(text[last_end:start]))
        wrapped_parts.append(text[start:end])
        last_end = end
    wrapped_parts.append(_apply_preprocessing(text[last_end:]))

    return _unescape_dollar(''.join(wrapped_parts))


def _find_math_regions(text: str) -> list[tuple[int, int]]:
    regions: list[tuple[int, int]] = []
    for m in re.finditer(r'(?<!\\)\$\$?([^$]+)(?<!\\)\$\$?', text):
        regions.append((m.start(), m.end()))
    for m in re.finditer(r'\\\[(.*?)\\\]', text, re.DOTALL):
        regions.append((m.start(), m.end()))
    regions.sort()
    return regions


def _merge_regions(regions: list[tuple[int, int]]) -> list[tuple[int, int]]:
    merged: list[tuple[int, int]] = []
    for start, end in regions:
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def _apply_preprocessing(text: str) -> str:
    result = text
    result = _replace_symbols(result)
    result = _preprocess_accents(result)
    result = _preprocess_subscripts(result)
    return result


def _replace_symbols(text: str) -> str:
    result = text
    sorted_symbols = sorted(_TEXT_SYMBOLS.items(), key=lambda x: -len(x[0]))
    for latex_cmd, unicode_char in sorted_symbols:
        result = re.sub(
            r'\\' + re.escape(latex_cmd) + r'(?![a-zA-Z])',
            unicode_char,
            result,
        )
    return result


def _preprocess_accents(text: str) -> str:
    result = text
    for accent_name, accent_char in _TEXT_ACCENTS.items():
        result = re.sub(
            r'\\' + re.escape(accent_name) + r'\{([^}]+)\}',
            lambda m: m.group(1) + accent_char,
            result,
        )
    return result


def _preprocess_subscripts(text: str) -> str:
    sub_map = {str(i): chr(0x2080 + i) for i in range(10)}
    sup_map = {
        str(i): chr(0x2070 + i) for i in range(10)
        if i not in (2, 3)
    }
    sup_map['2'] = '\u00b2'
    sup_map['3'] = '\u00b3'
    sup_map['1'] = '\u00b9'

    result = re.sub(
        r'_(?:\{([^}]+)\}|([a-zA-Z0-9]))',
        lambda m: ''.join(sub_map.get(c, c) for c in (m.group(1) or m.group(2) or '')),
        text,
    )
    result = re.sub(
        r'\^(?:\{([^}]+)\}|([a-zA-Z0-9]))',
        lambda m: ''.join(sup_map.get(c, c) for c in (m.group(1) or m.group(2) or '')),
        result,
    )
    return result


def _extract_cmd_info(groups: tuple[str | None, ...]) -> dict[str, Any] | None:
    if groups[0]:
        return {'cmd': groups[0], 'content': groups[1] or ''}
    if groups[2]:
        return {'cmd': groups[2], 'arg': groups[3] or '', 'content': groups[4] or ''}
    if groups[5]:
        return {'cmd': groups[5], 'level': int(groups[6] or '1'), 'content': groups[7] or ''}
    if groups[8]:
        return {'cmd': 'math', 'display': True, 'latex': groups[8] or ''}
    if groups[9]:
        return {'cmd': 'math', 'display': False, 'latex': groups[9] or ''}
    if groups[10]:
        return {'cmd': groups[10]}
    if groups[11]:
        return {'cmd': groups[11], 'image': groups[12] or ''}
    if groups[13]:
        return {'cmd': 'set_font', 'role': groups[13], 'font': groups[14] or ''}
    return None


def parse_latex_style(text: str) -> str:
    result = text

    while True:
        match = _TOKEN_PATTERN.search(result)
        if not match:
            break

        info = _extract_cmd_info(match.groups())
        if info is None:
            break

        cmd = info['cmd']

        inner = info.get('content', '')
        if inner:
            plain = _strip_styles(inner)
        else:
            plain = ''

        if cmd == 'newpage':
            result = result[:match.start()] + '\n\n---\n\n' + result[match.end():]
        elif cmd == 'includegraphics':
            image_path = info.get('image', '')
            result = result[:match.start()] + f'[Image: {image_path}]' + result[match.end():]
        elif cmd == 'math':
            latex = info.get('latex', '')
            result = result[:match.start()] + f'[Math: {latex}]' + result[match.end():]
        else:
            result = result[:match.start()] + plain + result[match.end():]

    return result


def _strip_styles(text: str) -> str:
    return parse_latex_style(text) if _TOKEN_PATTERN.search(text) else text


def parse_structured(text: str) -> list[dict[str, Any]]:
    text = preprocess_text(text)
    tokens: list[dict[str, Any]] = []
    pos = 0

    while pos < len(text):
        match = _TOKEN_PATTERN.search(text, pos)
        if not match:
            tokens.append({'type': 'text', 'content': text[pos:]})
            break

        if match.start() > pos:
            tokens.append({'type': 'text', 'content': text[pos:match.start()]})

        info = _extract_cmd_info(match.groups())
        if info is None:
            tokens.append({'type': 'text', 'content': text[pos:match.end()]})
            pos = match.end()
            continue

        cmd = info['cmd']
        token: dict[str, Any] = {'type': 'command', 'command': cmd}

        if cmd == 'math':
            token['display'] = info.get('display', False)
            token['latex'] = info.get('latex', '')
        elif cmd in STYLE_MAP:
            token['style'] = STYLE_MAP[cmd]
            token['content'] = info.get('content', '')
        elif cmd == 'fontfamily':
            token['font_name'] = info.get('arg', '')
            token['content'] = info.get('content', '')
        elif cmd == 'fontsize':
            token['font_size'] = int(info.get('arg', '12')) if info.get('arg') else 12
            token['content'] = info.get('content', '')
        elif cmd == 'color':
            token['color'] = info.get('arg', '')
            token['content'] = info.get('content', '')
        elif cmd == 'heading':
            token['level'] = info.get('level', 1)
            token['content'] = info.get('content', '')
        elif cmd == 'includegraphics':
            token['image'] = info.get('image', '')
        elif cmd == 'newpage':
            token['content'] = ''
        elif cmd == 'set_font':
            token['role'] = info.get('role', '')
            token['font'] = info.get('font', '')
            token['content'] = ''
        else:
            token['content'] = info.get('content', '')

        tokens.append(token)
        pos = match.end()

    return tokens
