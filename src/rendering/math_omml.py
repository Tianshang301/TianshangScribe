from __future__ import annotations

import re
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

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
    'ast': '\u2217', 'star': '\u22c6',
    'circ': '\u2218', 'bullet': '\u2022',
    'equiv': '\u2261', 'neq': '\u2260', 'approx': '\u2248',
    'sim': '\u223c', 'simeq': '\u2243', 'cong': '\u2245',
    'propto': '\u221d',
    'leq': '\u2264', 'geq': '\u2265', 'll': '\u226a', 'gg': '\u226b',
    'prec': '\u227a', 'succ': '\u227b', 'preceq': '\u2aaf', 'succeq': '\u2ab0',
    'subset': '\u2282', 'supset': '\u2283', 'subseteq': '\u2286', 'supseteq': '\u2287',
    'in': '\u2208', 'notin': '\u2209', 'ni': '\u220b',
    'forall': '\u2200', 'exists': '\u2203', 'nexists': '\u2204',
    'emptyset': '\u2205', 'varnothing': '\u2205',
    'partial': '\u2202', 'nabla': '\u2207',
    'to': '\u2192', 'rightarrow': '\u2192', 'Rightarrow': '\u21d2',
    'leftarrow': '\u2190', 'Leftarrow': '\u21d0',
    'leftrightarrow': '\u2194', 'Leftrightarrow': '\u21d4',
    'mapsto': '\u21a6',
    'uparrow': '\u2191', 'downarrow': '\u2193',
    'angle': '\u2220', 'triangle': '\u25b3',
    'perp': '\u27c2', 'parallel': '\u2225',
    'lnot': '\u00ac', 'neg': '\u00ac',
    'land': '\u2227', 'lor': '\u2228',
    'cap': '\u2229', 'cup': '\u222a',
    'oplus': '\u2295', 'ominus': '\u2296', 'otimes': '\u2297',
    'oslash': '\u2298', 'odot': '\u2299',
}

ACCENT_MAP: dict[str, str] = {
    'hat': '\u0302', 'widehat': '\u0302',
    'bar': '\u0304', 'overline': '\u0305',
    'tilde': '\u0303', 'widetilde': '\u0303',
    'dot': '\u0307', 'ddot': '\u0308',
    'vec': '\u20d7',
    'check': '\u030c',
    'acute': '\u0301',
    'grave': '\u0300',
    'breve': '\u0306',
}


def latex_to_omml(latex: str) -> Any:
    latex = latex.strip()
    if not latex:
        return None

    is_display = latex.startswith(r'\[') or latex.startswith('$$')
    if is_display:
        latex = re.sub(r'^(\\\[|\\begin\{equation\*\}|\\begin\{equation\}|\$\$)', '', latex)
        latex = re.sub(r'(\\\]|\\end\{equation\*\}|\\end\{equation\}|\$\$)$', '', latex)
    else:
        latex = re.sub(r'^\$', '', latex)
        latex = re.sub(r'\$$', '', latex)

    latex = latex.strip()
    if not latex:
        return None

    tokens = _tokenize(latex)
    elements = _build_omath_elements(tokens)

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
    sty_val_map = {
        'normal': 'p',
        'roman': 'p',
        'bold': 'b',
        'italic': 'i',
        'bold-italic': 'bi',
        'script': 'scr',
        'fraktur': 'fr',
        'double-struck': 'ds',
        'sans-serif': 'ss',
        'monospace': 'tt',
    }
    val = sty_val_map.get(math_style, 'p')
    sty.set(qn('m:val'), val)
    rpr.append(sty)
    return rpr


def _inject_style_recursive(el: Any, style: str, norm: bool) -> None:
    sty_val_map = {
        'normal': 'p', 'roman': 'p', 'bold': 'b', 'italic': 'i',
        'bold-italic': 'bi', 'script': 'scr', 'fraktur': 'fr',
        'double-struck': 'ds', 'sans-serif': 'ss', 'monospace': 'tt',
    }
    if el.tag == qn('m:r'):
        rpr = OxmlElement('m:rPr')
        if norm:
            nor_el = OxmlElement('m:nor')
            nor_el.set(qn('m:val'), '1')
            rpr.append(nor_el)
        sty = OxmlElement('m:sty')
        sty.set(qn('m:val'), sty_val_map.get(style, 'p'))
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


def _build_styled_elements(content: str, style: str) -> list[Any]:
    norm = style != 'italic'
    sub_tokens = _tokenize(content)
    elements = _stack_to_elements(sub_tokens)
    for el in elements:
        _inject_style_recursive(el, style, norm)
    return elements


def _collect_body_tokens(latex: str, start: int) -> tuple[list[dict[str, Any]], int]:
    tokens: list[dict[str, Any]] = []
    i = start
    length = len(latex)

    while i < length:
        ch = latex[i]
        if ch == '\\':
            match = re.match(r'\\([a-zA-Z]+)', latex[i:])
            if match:
                cmd = match.group(1)
                if cmd in ('frac', 'sqrt', 'sum', 'int', 'prod', 'lim',
                           'sin', 'cos', 'tan', 'log', 'ln', 'det', 'max',
                           'min', 'sup', 'inf', 'gcd', 'Pr', 'cot', 'sec',
                           'csc', 'vec', 'hat', 'bar', 'tilde', 'dot', 'ddot',
                           'mathbb', 'mathbf', 'mathit', 'mathrm', 'mathcal',
                           'left', 'right', 'int', 'iint', 'iiint', 'coprod',
                           'bigcup', 'bigcap', 'bigvee', 'bigwedge'):
                    break
            break
        elif ch == '^' or ch == '_':
            op_type = 'sup' if ch == '^' else 'sub'
            i += 1
            if i < length and latex[i] == '{':
                arg = _extract_one_arg(latex, i)
                if arg:
                    tokens.append({'type': op_type, 'content': arg[0]})
                    i = arg[1]
                else:
                    break
            elif i < length and latex[i].isalnum():
                tokens.append({'type': op_type, 'content': latex[i]})
                i += 1
            continue
        elif ch == ' ':
            i += 1
            continue
        else:
            exclude = ('\\', '^', '_', ' ')
            while i < length and latex[i] not in exclude:
                is_digit = latex[i].isdigit()
                start = i
                while i < length and latex[i] not in exclude and latex[i].isdigit() == is_digit:
                    i += 1
                text = latex[start:i]
                if text:
                    tokens.append({'type': 'text', 'text': text, 'norm': is_digit})
            continue

    return tokens, i


def _tokenize(latex: str) -> list[dict[str, Any]]:
    tokens: list[dict[str, Any]] = []
    i = 0
    length = len(latex)

    while i < length:
        ch = latex[i]

        if ch == '\\':
            match = re.match(
                r'\\([a-zA-Z]+)',
                latex[i:]
            )
            if match:
                cmd = match.group(1)
                cmd_end = match.end() + i

                if cmd == 'frac':
                    args = _extract_two_args(latex, cmd_end)
                    if args:
                        tokens.append({'type': 'frac', 'num': args[0], 'den': args[1]})
                        i = args[2]
                        continue

                elif cmd == 'sqrt':
                    args = _extract_sqrt_args(latex, cmd_end)
                    if args:
                        tokens.append({'type': 'sqrt', 'degree': args[0], 'content': args[1]})
                        i = args[2]
                        continue

                elif cmd in ('left', 'right'):
                    paren = _extract_delim(latex, cmd_end)
                    if paren:
                        tokens.append({'type': 'delim', 'char': paren[0], 'side': cmd})
                        i = paren[1]
                        continue

                elif cmd in ('sum', 'prod', 'int', 'iint', 'iiint', 'oint',
                             'coprod', 'bigcup', 'bigcap', 'bigvee', 'bigwedge'):
                    limits = _extract_limits(latex, cmd_end)
                    body_start = limits[2] if limits else cmd_end
                    body_tokens, body_end = _collect_body_tokens(latex, body_start)
                    tokens.append({
                        'type': 'nary',
                        'op': cmd,
                        'sub': limits[0] if limits else None,
                        'sup': limits[1] if limits else None,
                        'body': body_tokens,
                    })
                    i = body_end
                    continue

                elif cmd in ('lim', 'max', 'min', 'sup', 'inf', 'det', 'Pr', 'gcd',
                             'cot', 'sec', 'csc', 'deg', 'dim', 'hom', 'ker', 'arg'):
                    limits = _extract_limits(latex, cmd_end)
                    tokens.append({
                        'type': 'operator',
                        'op': cmd,
                        'sub': limits[0] if limits else None,
                        'sup': limits[1] if limits else None,
                    })
                    i = limits[2] if limits else cmd_end
                    continue

                elif cmd in ('sin', 'cos', 'tan', 'log', 'ln'):
                    tokens.append({'type': 'operator', 'op': cmd})
                    i = cmd_end
                    continue

                elif cmd in ('hat', 'widehat', 'bar', 'tilde', 'widetilde',
                             'dot', 'ddot', 'vec', 'check', 'acute', 'grave', 'breve'):
                    if cmd_end < length and latex[cmd_end] == '{':
                        args = _extract_one_arg(latex, cmd_end)
                        if args:
                            tokens.append({'type': 'accent', 'accent': cmd, 'content': args[0]})
                            i = args[1]
                            continue
                    i = cmd_end
                    continue

                elif cmd in GREEK_MAP:
                    tokens.append({
                        'type': 'text', 'text': GREEK_MAP[cmd],
                        'norm': cmd[0].isupper(),
                    })
                    i = cmd_end
                    continue

                elif cmd in SYMBOL_MAP:
                    tokens.append({'type': 'text', 'text': SYMBOL_MAP[cmd]})
                    i = cmd_end
                    continue

                elif cmd in ('text', 'mathrm', 'mathbf', 'mathsf', 'mathtt',
                             'mathit', 'mathcal', 'mathbb'):
                    style_map_cmd = {
                        'mathrm': 'normal', 'mathbf': 'bold', 'mathit': 'italic',
                        'mathsf': 'sans-serif', 'mathtt': 'monospace',
                        'mathbb': 'double-struck',
                        'text': 'normal', 'mathcal': 'script',
                    }
                    args = _extract_one_arg(latex, cmd_end)
                    if args:
                        tokens.append({
                            'type': 'styled',
                            'style': style_map_cmd.get(cmd, 'normal'),
                            'content': args[0],
                        })
                        i = args[1]
                        continue

                else:
                    tokens.append({'type': 'text', 'text': '\\' + cmd})
                    i = cmd_end
                    continue
            else:
                tokens.append({'type': 'text', 'text': ch})
                i += 1
                continue

        elif ch in ('^', '_'):
            op_type = 'sup' if ch == '^' else 'sub'
            if i + 1 < length:
                if latex[i + 1] == '{':
                    args = _extract_one_arg(latex, i + 1)
                    if args:
                        tokens.append({'type': op_type, 'content': args[0]})
                        i = args[1]
                        continue
                else:
                    end = i + 1
                    while end < length and latex[end].isalnum():
                        end += 1
                    tokens.append({'type': op_type, 'content': latex[i + 1:end]})
                    i = end
                    continue
            i += 1
            continue

        elif ch == '(' or ch == ')' or ch == '[' or ch == ']' or ch == '{' or ch == '}':
            tokens.append({'type': 'text', 'text': ch})
            i += 1
            continue

        elif ch == ' ':
            i += 1
            continue

        else:
            exclude = ('\\', '^', '_', ' ', '(', ')', '[', ']', '{', '}')
            while i < length and latex[i] not in exclude:
                is_digit = latex[i].isdigit()
                start = i
                while i < length and latex[i] not in exclude and latex[i].isdigit() == is_digit:
                    i += 1
                text = latex[start:i]
                if text:
                    tokens.append({'type': 'text', 'text': text, 'norm': is_digit})
            continue

    return tokens


def _extract_one_arg(s: str, start: int) -> tuple[str, int] | None:
    if start >= len(s) or s[start] != '{':
        return None

    depth = 0
    for i in range(start, len(s)):
        if s[i] == '{':
            depth += 1
        elif s[i] == '}':
            depth -= 1
            if depth == 0:
                return s[start + 1:i], i + 1

    return s[start + 1:], len(s)


def _extract_two_args(s: str, start: int) -> tuple[str, str, int] | None:
    first = _extract_one_arg(s, start)
    if first is None:
        return None
    second = _extract_one_arg(s, first[1])
    if second is None:
        return None
    return first[0], second[0], second[1]


def _extract_sqrt_args(s: str, start: int) -> tuple[str, str, int] | None:
    if start >= len(s):
        return None

    if s[start] == '[':
        bracket_end = s.index(']', start) + 1
        degree = s[start + 1:bracket_end - 1]
        content = _extract_one_arg(s, bracket_end)
        if content:
            return degree, content[0], content[1]
        return None

    content = _extract_one_arg(s, start)
    if content:
        return '', content[0], content[1]
    return None


def _extract_delim(s: str, start: int) -> tuple[str, int] | None:
    if start >= len(s):
        return None

    delim_map = {
        '(': '\u0028', ')': '\u0029',
        '[': '\u005b', ']': '\u005d',
        '{': '\u007b', '}': '\u007d',
        '|': '\u007c',
        '.': None,
        '\\{': '\u007b', '\\}': '\u007d',
        '\\langle': '\u2329', '\\rangle': '\u232a',
        '\\lceil': '\u2308', '\\rceil': '\u2309',
        '\\lfloor': '\u230a', '\\rfloor': '\u230b',
    }

    ch = s[start]
    if ch == '\\':
        match = re.match(r'\\([a-zA-Z]+)', s[start:])
        if match:
            key = '\\' + match.group(1)
            if key in delim_map:
                return delim_map[key], match.end()
    elif ch in delim_map:
        return delim_map[ch], start + 1

    return None


def _extract_limits(s: str, start: int) -> tuple[str, str, int] | None:
    sub_val = None
    sup_val = None
    pos = start

    if pos < len(s) and s[pos] == '_':
        if pos + 1 < len(s) and s[pos + 1] == '{':
            sub = _extract_one_arg(s, pos + 1)
            if sub:
                sub_val = sub[0]
                pos = sub[1]
        else:
            end = pos + 1
            while end < len(s) and s[end].isalnum():
                end += 1
            sub_val = s[pos + 1:end]
            pos = end

    if pos < len(s) and s[pos] == '^':
        if pos + 1 < len(s) and s[pos + 1] == '{':
            sup = _extract_one_arg(s, pos + 1)
            if sup:
                sup_val = sup[0]
                pos = sup[1]
        else:
            end = pos + 1
            while end < len(s) and s[end].isalnum():
                end += 1
            sup_val = s[pos + 1:end]
            pos = end

    if sub_val or sup_val:
        return sub_val, sup_val, pos
    return None


def _rpr_equivalent(a: Any, b: Any) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return etree.tostring(a) == etree.tostring(b)


def _build_omath_elements(tokens: list[dict[str, Any]]) -> list[Any]:
    grouped = _group_sup_sub(tokens)
    result_tokens: list[Any] = []

    for item in grouped:
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
            rpr_child = next((c for c in el if c.tag == qn('m:rPr')), None)
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


def _stack_to_elements(tokens: list[dict[str, Any]]) -> list[Any]:
    grouped = _group_sup_sub(tokens)
    elements: list[Any] = []
    pending: list[Any] = []
    pending_rpr = None

    for item in grouped:
        elem = _token_to_omml(item)
        if elem is None:
            continue
        flat = elem if isinstance(elem, list) else [elem]
        for el in flat:
            tag = el.tag
            if tag == qn('m:r'):
                mt = [c for c in el if c.tag == qn('m:t')]
                rpr = next((c for c in el if c.tag == qn('m:rPr')), None)
                if pending and not _rpr_equivalent(pending_rpr, rpr):
                    mr = OxmlElement('m:r')
                    if pending_rpr is not None:
                        from copy import deepcopy
                        mr.append(deepcopy(pending_rpr))
                    for t in pending:
                        mr.append(t)
                    elements.append(mr)
                    pending = []
                    pending_rpr = None
                pending.extend(mt)
                if pending_rpr is None and rpr is not None:
                    pending_rpr = rpr
            elif tag == qn('m:t'):
                pending.append(el)
            else:
                if pending:
                    mr = OxmlElement('m:r')
                    if pending_rpr is not None:
                        from copy import deepcopy
                        mr.append(deepcopy(pending_rpr))
                    for t in pending:
                        mr.append(t)
                    elements.append(mr)
                    pending = []
                    pending_rpr = None
                elements.append(el)

    if pending:
        mr = OxmlElement('m:r')
        if pending_rpr is not None:
            from copy import deepcopy
            mr.append(deepcopy(pending_rpr))
        for t in pending:
            mr.append(t)
        elements.append(mr)

    return elements


def _stack_to_element(tokens: list[dict[str, Any]]) -> Any:
    grouped = _group_sup_sub(tokens)
    result_tokens: list[dict[str, Any]] = []

    for item in grouped:
        elem = _token_to_omml(item)
        if elem is not None:
            result_tokens.append({'type': 'el', 'el': elem})

    if not result_tokens:
        return _make_run('')

    if len(result_tokens) == 1:
        return result_tokens[0]['el']

    m_run = OxmlElement('m:r')
    merged_rpr = None
    for item in result_tokens:
        el = item['el']
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


def _group_sup_sub(tokens: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    i = 0
    while i < len(tokens):
        token = tokens[i]
        if i + 1 < len(tokens) and tokens[i + 1]['type'] in ('sub', 'sup'):
            base = token
            sub_token = None
            sup_token = None
            j = i + 1
            while j < len(tokens) and tokens[j]['type'] in ('sub', 'sup'):
                if tokens[j]['type'] == 'sub':
                    sub_token = tokens[j]
                else:
                    sup_token = tokens[j]
                j += 1

            if sub_token and sup_token:
                result.append({
                    'type': 'subsup',
                    'base': base,
                    'sub': sub_token['content'],
                    'sup': sup_token['content'],
                })
            elif sub_token:
                result.append({'type': 'sub', 'base': base, 'content': sub_token['content']})
            elif sup_token:
                result.append({'type': 'sup', 'base': base, 'content': sup_token['content']})
            else:
                result.append(base)

            i = j
        else:
            result.append(token)
            i += 1
    return result


def _token_to_omml(token: dict[str, Any]) -> Any:
    t = token.get('type', 'text')

    if t == 'text':
        return _make_run(token.get('text', ''), norm=token.get('norm', False))

    elif t == 'operator':
        op_name = token.get('op', '')
        sub_val = token.get('sub')
        sup_val = token.get('sup')

        if sub_val or sup_val:
            op_elem = OxmlElement('m:sSubSup' if (sub_val and sup_val)
                                  else ('m:sSub' if sub_val else 'm:sSup'))
            e = OxmlElement('m:e')
            op_run = _make_run(op_name, 'normal', norm=True)
            e.append(op_run)
            op_elem.append(e)
            if sub_val:
                sub_e = OxmlElement('m:sub')
                for el in _build_sub_omath(sub_val):
                    sub_e.append(el)
                op_elem.append(sub_e)
            if sup_val:
                sup_e = OxmlElement('m:sup')
                for el in _build_sub_omath(sup_val):
                    sup_e.append(el)
                op_elem.append(sup_e)
            return op_elem
        else:
            return _make_run(op_name, 'normal', norm=True)

    elif t == 'frac':
        f = OxmlElement('m:f')
        num = OxmlElement('m:num')
        for el in _build_sub_omath(token.get('num', '')):
            num.append(el)
        f.append(num)
        den = OxmlElement('m:den')
        for el in _build_sub_omath(token.get('den', '')):
            den.append(el)
        f.append(den)
        return f

    elif t == 'sqrt':
        rad = OxmlElement('m:rad')
        deg_val = token.get('degree', '')
        if deg_val:
            deg = OxmlElement('m:deg')
            for el in _build_sub_omath(deg_val):
                deg.append(el)
            rad.append(deg)
        else:
            rad_pr = OxmlElement('m:radPr')
            deg_hide = OxmlElement('m:degHide')
            deg_hide.set(qn('m:val'), '1')
            rad_pr.append(deg_hide)
            rad.insert(0, rad_pr)
        e = OxmlElement('m:e')
        for el in _build_sub_omath(token.get('content', '')):
            e.append(el)
        rad.append(e)
        return rad

    elif t == 'nary':
        nary = OxmlElement('m:nary')
        op_name = token.get('op', 'sum')
        op_map = {
            'sum': '\u2211', 'prod': '\u220f', 'int': '\u222b',
            'iint': '\u222c', 'iiint': '\u222d', 'oint': '\u222e',
            'coprod': '\u2210', 'bigcup': '\u22c3', 'bigcap': '\u22c2',
            'bigvee': '\u22c1', 'bigwedge': '\u22c0',
        }

        nary_pr = OxmlElement('m:naryPr')
        chr_el = OxmlElement('m:chr')
        chr_el.set(qn('m:val'), op_map.get(op_name, '\u2211'))
        nary_pr.append(chr_el)
        nary.append(nary_pr)

        sub_val = token.get('sub')
        if sub_val:
            sub = OxmlElement('m:sub')
            for el in _build_sub_omath(sub_val):

                sub.append(el)
            nary.append(sub)

        sup_val = token.get('sup')
        if sup_val:
            sup = OxmlElement('m:sup')
            for el in _build_sub_omath(sup_val):

                sup.append(el)
            nary.append(sup)

        body = token.get('body')
        if body:
            e = OxmlElement('m:e')
            for el in _stack_to_elements(body):
                e.append(el)
            nary.append(e)

        return nary

    elif t == 'sub':
        ssub = OxmlElement('m:sSub')
        e = OxmlElement('m:e')
        for el in _build_sub_omath(token.get('base', {}).get('text', '')):
            e.append(el)
        ssub.append(e)
        sub = OxmlElement('m:sub')
        for el in _build_sub_omath(token.get('content', '')):
            sub.append(el)
        ssub.append(sub)
        return ssub

    elif t == 'sup':
        ssup = OxmlElement('m:sSup')
        e = OxmlElement('m:e')
        for el in _build_sub_omath(token.get('base', {}).get('text', '')):
            e.append(el)
        ssup.append(e)
        sup = OxmlElement('m:sup')
        for el in _build_sub_omath(token.get('content', '')):
            sup.append(el)
        ssup.append(sup)
        return ssup

    elif t == 'subsup':
        ss = OxmlElement('m:sSubSup')
        e = OxmlElement('m:e')
        for el in _build_sub_omath(token.get('base', {}).get('text', '')):
            e.append(el)
        ss.append(e)
        sub_val = token.get('sub', '')
        sub = OxmlElement('m:sub')
        for el in _build_sub_omath(sub_val):

            sub.append(el)
        ss.append(sub)
        sup_val = token.get('sup', '')
        sup = OxmlElement('m:sup')
        for el in _build_sub_omath(sup_val):

            sup.append(el)
        ss.append(sup)
        return ss

    elif t == 'accent':
        acc = OxmlElement('m:acc')
        acc_pr = OxmlElement('m:accPr')
        chr_el = OxmlElement('m:chr')
        accent_map_omml = {
            'hat': '\u0302', 'widehat': '\u0302', 'bar': '\u0304',
            'tilde': '\u0303', 'widetilde': '\u0303',
            'dot': '\u0307', 'ddot': '\u0308', 'vec': '\u20d7',
            'check': '\u030c', 'acute': '\u0301', 'grave': '\u0300', 'breve': '\u0306',
        }
        chr_el.set(qn('m:val'), accent_map_omml.get(token.get('accent', 'hat'), '\u0302'))
        acc_pr.append(chr_el)
        acc.append(acc_pr)
        e = OxmlElement('m:e')
        for el in _build_sub_omath(token.get('content', '')):
            e.append(el)
        acc.append(e)
        return acc

    elif t == 'styled':
        elements = _build_styled_elements(
            token.get('content', ''),
            token.get('style', 'normal'),
        )
        return elements if len(elements) == 1 else elements

    return _make_run('')


def _build_sub_omath(content: str | dict) -> list[Any]:
    if isinstance(content, str):
        if not content:
            return [_make_run('')]
        sub_tokens = _tokenize(content)
        return _stack_to_elements(sub_tokens)
    return [_make_run(str(content))]
