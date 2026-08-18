from __future__ import annotations

import json
from pathlib import Path

from docx.oxml.ns import qn
from lxml import etree

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
    _build_sub_omath,
    _collect_body_tokens,
    _extract_delim,
    _extract_limits,
    _extract_one_arg,
    _extract_sqrt_args,
    _extract_two_args,
    _make_run,
    _make_run_props,
    _rpr_equivalent,
    _stack_to_element,
    _token_to_omml,
    latex_to_omml,
    parse_expression,
)

GOLDEN_FILE = Path(__file__).resolve().parents[2] / 'testdata' / 'omml_golden.json'


def parse(s: str) -> list:
    return parse_expression(ParserContext(s))


class TestMathTokenize:
    def test_plain_text(self) -> None:
        tokens = parse('x')
        assert len(tokens) == 1
        assert tokens[0] == TextToken(text='x', norm=False)

    def test_superscript(self) -> None:
        tokens = parse('x^2')
        assert len(tokens) == 1
        assert isinstance(tokens[0], SubSupToken)
        assert tokens[0].base == TextToken(text='x', norm=False)
        assert tokens[0].sup == [TextToken(text='2', norm=True)]

    def test_superscript_braces(self) -> None:
        tokens = parse('x^{n+1}')
        assert isinstance(tokens[0], SubSupToken)
        assert tokens[0].sup == [
            TextToken(text='n+', norm=False),
            TextToken(text='1', norm=True),
        ]

    def test_subscript(self) -> None:
        tokens = parse('x_i')
        assert isinstance(tokens[0], SubSupToken)
        assert tokens[0].sub == [TextToken(text='i', norm=False)]

    def test_subscript_braces(self) -> None:
        tokens = parse('x_{i+1}')
        assert tokens[0].sub == [
            TextToken(text='i+', norm=False),
            TextToken(text='1', norm=True),
        ]

    def test_sub_and_sup(self) -> None:
        tokens = parse('x_i^2')
        assert isinstance(tokens[0], SubSupToken)
        assert tokens[0].sub == [TextToken(text='i', norm=False)]
        assert tokens[0].sup == [TextToken(text='2', norm=True)]

    def test_fraction(self) -> None:
        tokens = parse(r'\frac{a}{b}')
        assert isinstance(tokens[0], FracToken)
        assert tokens[0].num == [TextToken(text='a', norm=False)]
        assert tokens[0].den == [TextToken(text='b', norm=False)]

    def test_fraction_complex(self) -> None:
        tokens = parse(r'\frac{x+y}{z}')
        assert tokens[0].num == [
            TextToken(text='x+y', norm=False),
        ]
        assert tokens[0].den == [TextToken(text='z', norm=False)]

    def test_sqrt(self) -> None:
        tokens = parse(r'\sqrt{x}')
        assert isinstance(tokens[0], SqrtToken)
        assert tokens[0].degree is None
        assert tokens[0].content == [TextToken(text='x', norm=False)]

    def test_sqrt_nth(self) -> None:
        tokens = parse(r'\sqrt[3]{x}')
        assert tokens[0].degree == [TextToken(text='3', norm=True)]
        assert tokens[0].content == [TextToken(text='x', norm=False)]

    def test_greek_letters(self) -> None:
        tokens = parse(r'\alpha \beta \Gamma')
        assert tokens[0] == TextToken(text='\u03b1', norm=False)
        assert tokens[1] == TextToken(text='\u03b2', norm=False)
        assert tokens[2] == TextToken(text='\u0393', norm=True)

    def test_operator_sum(self) -> None:
        tokens = parse(r'\sum_{i=0}^{n}')
        assert isinstance(tokens[0], NaryToken)
        assert tokens[0].op == 'sum'
        assert tokens[0].sub == [
            TextToken(text='i=', norm=False),
            TextToken(text='0', norm=True),
        ]
        assert tokens[0].sup == [TextToken(text='n', norm=False)]

    def test_operator_int(self) -> None:
        tokens = parse(r'\int_{0}^{\infty}')
        assert tokens[0].op == 'int'
        assert tokens[0].sub == [TextToken(text='0', norm=True)]
        assert tokens[0].sup == [TextToken(text='\u221e', norm=False)]

    def test_accent(self) -> None:
        tokens = parse(r'\hat{x}')
        assert isinstance(tokens[0], AccentToken)
        assert tokens[0].accent == 'hat'
        assert tokens[0].content == [TextToken(text='x', norm=False)]

    def test_accent_bar(self) -> None:
        tokens = parse(r'\bar{y}')
        assert tokens[0].accent == 'bar'

    def test_symbols(self) -> None:
        tokens = parse(r'\infty \pm \times \div')
        assert tokens[0] == TextToken(text='\u221e', norm=False)
        assert tokens[1] == TextToken(text='\u00b1', norm=False)
        assert tokens[2] == TextToken(text='\u00d7', norm=False)
        assert tokens[3] == TextToken(text='\u00f7', norm=False)

    def test_digits_have_norm_true(self) -> None:
        tokens = parse(r'\frac{42}{2a}')
        frac = tokens[0]
        assert frac.num == [TextToken(text='42', norm=True)]
        assert frac.den == [
            TextToken(text='2', norm=True),
            TextToken(text='a', norm=False),
        ]

    def test_mixed_text_splits_at_digit_boundary(self) -> None:
        tokens = parse(r'\frac{4ac}{2}')
        frac = tokens[0]
        assert frac.num == [
            TextToken(text='4', norm=True),
            TextToken(text='ac', norm=False),
        ]
        assert frac.den == [TextToken(text='2', norm=True)]

    def test_plain_letters_no_norm(self) -> None:
        tokens = parse('abc')
        assert tokens == [TextToken(text='abc', norm=False)]

    def test_greek_uppercase_has_norm(self) -> None:
        tokens = parse(r'\Gamma \Delta \Theta \alpha \beta')
        assert tokens[0].norm is True
        assert tokens[1].norm is True
        assert tokens[2].norm is True
        assert tokens[3].norm is False
        assert tokens[4].norm is False

    def test_styled_mathrm_not_lost(self) -> None:
        tokens = parse(r'\mathrm{abc}')
        assert len(tokens) == 1
        assert isinstance(tokens[0], StyledToken)
        assert tokens[0].style == 'normal'
        assert tokens[0].content == [TextToken(text='abc', norm=False)]

    def test_styled_mathbf_not_lost(self) -> None:
        tokens = parse(r'\mathbf{def}')
        assert tokens[0].style == 'bold'
        assert tokens[0].content == [TextToken(text='def', norm=False)]

    def test_collect_body_tokens_splits_digits(self) -> None:
        tokens, _ = _collect_body_tokens('4ab', 0)
        assert tokens[0] == TextToken(text='4', norm=True)
        assert tokens[1] == TextToken(text='ab', norm=False)


class TestLatexToOMML:
    def test_simple_expression_returns_element(self) -> None:
        result = latex_to_omml('x = y')
        assert result is not None
        assert 'Math' in str(result.tag) or 'oMath' in result.tag.split('}')[-1]

    def test_fraction_omml(self) -> None:
        result = latex_to_omml(r'\frac{a}{b}')
        assert result is not None

    def test_superscript_omml(self) -> None:
        result = latex_to_omml('x^2')
        assert result is not None

    def test_display_math(self) -> None:
        result = latex_to_omml(r'$$x = 1$$')
        assert result is not None

    def test_empty_returns_none(self) -> None:
        result = latex_to_omml('')
        assert result is None

    def test_nary_sum_omml(self) -> None:
        result = latex_to_omml(r'\sum_{i=0}^{n} i^2')
        assert result is not None

    def test_sqrt_omml(self) -> None:
        result = latex_to_omml(r'\sqrt{a^2 + b^2}')
        assert result is not None

    def test_accents_omml(self) -> None:
        result = latex_to_omml(r'\hat{x} + \bar{y}')
        assert result is not None

    def test_lim_omml(self) -> None:
        result = latex_to_omml(r'\lim_{x \to 0} f(x)')
        assert result is not None

    def test_sin_omml(self) -> None:
        result = latex_to_omml(r'\sin x')
        assert result is not None

    def test_log_omml(self) -> None:
        result = latex_to_omml(r'\log_{2} 8 = 3')
        assert result is not None

    def test_mathrm_produces_nor(self) -> None:
        result = latex_to_omml(r'\mathrm{abc}')
        assert result is not None
        nor_els = result.findall('.//' + qn('m:nor'))
        assert len(nor_els) >= 1
        assert nor_els[0].get(qn('m:val')) == '1'
        t_els = result.findall('.//' + qn('m:t'))
        texts = [t.text or '' for t in t_els]
        assert 'abc' in texts

    def test_mathbf_produces_bold_nor(self) -> None:
        result = latex_to_omml(r'\mathbf{def}')
        assert result is not None
        sty_els = result.findall('.//' + qn('m:sty'))
        found_bold = any(s.get(qn('m:val')) == 'b' for s in sty_els)
        assert found_bold
        nor_els = result.findall('.//' + qn('m:nor'))
        assert len(nor_els) >= 1

    def test_mathit_no_nor(self) -> None:
        result = latex_to_omml(r'\mathit{ghi}')
        assert result is not None
        nor_els = result.findall('.//' + qn('m:nor'))
        assert len(nor_els) == 0

    def test_operator_sin_has_nor(self) -> None:
        result = latex_to_omml(r'\sin x')
        assert result is not None
        nor_els = result.findall('.//' + qn('m:nor'))
        assert len(nor_els) >= 1

    def test_digits_in_frac_have_nor(self) -> None:
        result = latex_to_omml(r'\frac{4}{2}')
        assert result is not None
        nor_els = result.findall('.//' + qn('m:nor'))
        assert len(nor_els) >= 2

    def test_styled_in_frac_not_lost(self) -> None:
        result = latex_to_omml(r'\frac{\mathrm{e}}{2}')
        assert result is not None
        t_els = result.findall('.//' + qn('m:t'))
        texts = [t.text or '' for t in t_els]
        assert 'e' in texts
        assert '2' in texts

    def test_left_right_braces(self) -> None:
        result = latex_to_omml(r'\left\{ a \right\}')
        assert result is not None

    def test_left_right_named_delims(self) -> None:
        result = latex_to_omml(r'\left\langle x \right\rangle')
        assert result is not None

    def test_left_right_floors(self) -> None:
        result = latex_to_omml(r'\left\lfloor a \right\rfloor')
        assert result is not None

    def test_left_right_ceil(self) -> None:
        result = latex_to_omml(r'\left\lceil b \right\rceil')
        assert result is not None

    def test_left_right_pipe(self) -> None:
        result = latex_to_omml(r'\left| x \right|')
        assert result is not None

    def test_left_right_bars(self) -> None:
        result = latex_to_omml(r'\left\| x \right\|')
        assert result is not None

    def test_left_dot_right(self) -> None:
        result = latex_to_omml(r'\left. a \right|')
        assert result is not None

    def test_dollar_inline(self) -> None:
        result = latex_to_omml(r'$x^2$')
        assert result is not None

    def test_display_bracket(self) -> None:
        result = latex_to_omml(r'\[x = 1\]')
        assert result is not None

    def test_display_equation_env(self) -> None:
        result = latex_to_omml(r'\begin{equation}x\end{equation}')
        assert result is not None

    def test_whitespace_only_returns_none(self) -> None:
        result = latex_to_omml('   ')
        assert result is None

    def test_dollar_only_returns_none(self) -> None:
        result = latex_to_omml('$')
        assert result is None

    def test_unknown_command_becomes_text(self) -> None:
        result = latex_to_omml(r'\unknown{x}')
        assert result is not None

    def test_bare_backslash_command(self) -> None:
        result = latex_to_omml(r'\leftx')
        assert result is not None

    def test_accent_without_braces(self) -> None:
        result = latex_to_omml(r'\hat x')
        assert result is not None

    def test_accent_unknown(self) -> None:
        result = latex_to_omml(r'\hatx')
        assert result is not None

    def test_sqrt_nth_degree(self) -> None:
        result = latex_to_omml(r'\sqrt[3]{x}')
        assert result is not None
        rad = result.find('.//' + qn('m:rad'))
        assert rad is not None
        deg = rad.find(qn('m:deg'))
        assert deg is not None

    def test_sqrt_no_degree_hides(self) -> None:
        result = latex_to_omml(r'\sqrt{x}')
        assert result is not None
        hide = result.findall('.//' + qn('m:degHide'))
        assert len(hide) >= 1

    def test_operator_with_sub_only(self) -> None:
        result = latex_to_omml(r'\log_2 8')
        assert result is not None
        assert result.find('.//' + qn('m:sSub')) is not None

    def test_operator_with_sup_only(self) -> None:
        result = latex_to_omml(r'\lim^x f')
        assert result is not None

    def test_operator_plain(self) -> None:
        result = latex_to_omml(r'\sin x')
        assert result is not None

    def test_plain_operator_tokens(self) -> None:
        result = latex_to_omml(r'\max f')
        assert result is not None

    def test_limits_single_char(self) -> None:
        result = latex_to_omml(r'\sum_0^n x')
        assert result is not None

    def test_limits_no_braces(self) -> None:
        result = latex_to_omml(r'\sum_{i}^{n} x')
        assert result is not None

    def test_sub_braces_sup_single(self) -> None:
        result = latex_to_omml(r'\int_{a}^{b} f')
        assert result is not None

    def test_sub_single_sup_braces(self) -> None:
        result = latex_to_omml(r'\int_0^{\infty} f')
        assert result is not None

    def test_frac_missing_args(self) -> None:
        result = latex_to_omml(r'\frac{a}')
        assert result is not None

    def test_frac_no_args(self) -> None:
        result = latex_to_omml(r'\frac')
        assert result is not None

    def test_sqrt_no_args(self) -> None:
        result = latex_to_omml(r'\sqrt')
        assert result is not None

    def test_left_bare_fallback(self) -> None:
        result = latex_to_omml(r'\left')
        assert result is not None

    def test_right_bare_fallback(self) -> None:
        result = latex_to_omml(r'\right')
        assert result is not None

    def test_bare_styled_command(self) -> None:
        for t in (r'\mathrm', r'\mathbf', r'\mathbb', r'\mathcal'):
            result = latex_to_omml(t)
            assert result is not None

    def test_lone_backslash(self) -> None:
        result = latex_to_omml('a \\ b')
        assert result is not None

    def test_trailing_sup(self) -> None:
        result = latex_to_omml('x^')
        assert result is not None

    def test_trailing_sub(self) -> None:
        result = latex_to_omml('x_')
        assert result is not None

    def test_script_styles(self) -> None:
        result = latex_to_omml(r'\mathcal{L} + \mathbb{R}')
        assert result is not None

    def test_text_command(self) -> None:
        result = latex_to_omml(r'\text{hello world}')
        assert result is not None
        texts = [t.text or '' for t in result.findall('.//' + qn('m:t'))]
        assert 'hello' in texts
        assert 'world' in texts

    def test_mathsf_mattset(self) -> None:
        result = latex_to_omml(r'\mathsf{ab} + \mathtt{cd}')
        assert result is not None

    def test_fraktur_unknown_style(self) -> None:
        result = latex_to_omml(r'\mathbf{xyz}')
        assert result is not None
        sty = result.findall('.//' + qn('m:sty'))
        assert any(s.get(qn('m:val')) == 'b' for s in sty)

    def test_bold_italic_style_map(self) -> None:
        rpr = _make_run_props('bold-italic')
        sty = rpr.find(qn('m:sty'))
        assert sty is not None
        assert sty.get(qn('m:val')) == 'bi'

    def test_nested_styled_recursive_rpr_merge(self) -> None:
        result = latex_to_omml(r'\mathbf{42}')
        assert result is not None

    def test_greek_maps(self) -> None:
        result = latex_to_omml(r'\alpha\beta\gamma\delta\Gamma\Delta')
        assert result is not None

    def test_symbols_more(self) -> None:
        result = latex_to_omml(r'\pm\mp\times\div\ast\star\circ\bullet')
        assert result is not None

    def test_symbols_relations(self) -> None:
        result = latex_to_omml(r'\leq\geq\ll\gg\subset\subseteq\cup\cap')
        assert result is not None

    def test_symbols_logic(self) -> None:
        result = latex_to_omml(r'\land\lor\lnot\forall\exists\in\notin')
        assert result is not None

    def test_arrows(self) -> None:
        result = latex_to_omml(
            r'\to\rightarrow\leftarrow\Leftarrow\Leftrightarrow'
        )
        assert result is not None

    def test_neg(self) -> None:
        result = latex_to_omml(r'\neg p \wedge q')
        assert result is not None

    def test_propto_sim(self) -> None:
        result = latex_to_omml(r'a \propto b \sim c \simeq d')
        assert result is not None

    def test_emptyset_partial(self) -> None:
        result = latex_to_omml(r'\emptyset \partial \nabla')
        assert result is not None

    def test_bigops(self) -> None:
        result = latex_to_omml(r'\bigcup \bigcap \bigvee \bigwedge \coprod')
        assert result is not None

    def test_oint_iint_iiint(self) -> None:
        result = latex_to_omml(r'\oint \iint \iiint')
        assert result is not None

    def test_vartheta_varrho(self) -> None:
        result = latex_to_omml(
            r'\vartheta \varrho \varphi \varpi \varsigma'
        )
        assert result is not None

    def test_det_pr(self) -> None:
        result = latex_to_omml(r'\det A + \Pr(x)')
        assert result is not None

    def test_lim_sup(self) -> None:
        result = latex_to_omml(r'\limsup')
        assert result is not None

    def test_inline_unsupported_symbol(self) -> None:
        result = latex_to_omml(r'2x + \frac{3}{4} = 5')
        assert result is not None

    def test_text_with_parens(self) -> None:
        result = latex_to_omml(r'(a + b)')
        assert result is not None

    def test_grouping_braces_plain(self) -> None:
        result = latex_to_omml(r'{x}')
        assert result is not None

    def test_all_accent_commands(self) -> None:
        result = latex_to_omml(
            r'\vec{v}\hat{w}\bar{x}\tilde{y}\dot{z}\ddot{a}\widehat{AB}'
            r'\widetilde{CD}\check{e}\acute{o}\grave{u}\breve{i}'
        )
        assert result is not None

    def test_accent_no_braces_fallback(self) -> None:
        result = latex_to_omml(r'\hatx')
        assert result is not None

    def test_bare_operator_commands(self) -> None:
        for t in (
            r'\lim', r'\max', r'\min', r'\sup', r'\inf', r'\det',
            r'\gcd', r'\Pr',
        ):
            result = latex_to_omml(t)
            assert result is not None

    def test_styled_nested(self) -> None:
        result = latex_to_omml(r'\mathbf{\mathbb{R}}')
        assert result is not None

    def test_subsup_both(self) -> None:
        result = latex_to_omml('x_i^2')
        assert result is not None
        assert result.find('.//' + qn('m:sSubSup')) is not None

    def test_subsup_in_frac(self) -> None:
        result = latex_to_omml(r'\frac{x_i^2}{y_j^3}')
        assert result is not None

    def test_script_double_struck(self) -> None:
        result = latex_to_omml(r'\mathbb{Z}')
        assert result is not None
        sty = result.findall('.//' + qn('m:sty'))
        assert any(s.get(qn('m:val')) == 'ds' for s in sty)

    def test_mathcal_script(self) -> None:
        result = latex_to_omml(r'\mathcal{F}')
        assert result is not None
        sty = result.findall('.//' + qn('m:sty'))
        assert any(s.get(qn('m:val')) == 'scr' for s in sty)

    def test_mathsf_sans(self) -> None:
        result = latex_to_omml(r'\mathsf{G}')
        assert result is not None
        sty = result.findall('.//' + qn('m:sty'))
        assert any(s.get(qn('m:val')) == 'ss' for s in sty)

    def test_mathtt_monospace(self) -> None:
        result = latex_to_omml(r'\mathtt{H}')
        assert result is not None
        sty = result.findall('.//' + qn('m:sty'))
        assert any(s.get(qn('m:val')) == 'tt' for s in sty)


class TestInternalHelpers:
    def test_build_sub_omml_tokens(self) -> None:
        els = _build_sub_omath([TextToken(text='x', norm=False)])
        assert len(els) == 1

    def test_build_sub_omml_empty(self) -> None:
        els = _build_sub_omath([])
        assert len(els) == 1
        assert els[0].tag == qn('m:r')

    def test_build_sub_omml_nested(self) -> None:
        els = _build_sub_omath([FracToken(
            num=[TextToken(text='a', norm=False)],
            den=[TextToken(text='b', norm=False)],
        )])
        assert len(els) == 1
        assert els[0].tag == qn('m:f')

    def test_stack_to_element_single(self) -> None:
        el = _stack_to_element(parse('x'))
        assert el is not None
        assert qn('m:t') in [c.tag for c in el]

    def test_stack_to_element_merged(self) -> None:
        el = _stack_to_element(parse('ab'))
        assert el.tag == qn('m:r')

    def test_stack_to_element_empty(self) -> None:
        el = _stack_to_element([])
        assert el.tag == qn('m:r')

    def test_make_run_with_style(self) -> None:
        r = _make_run('x', 'bold', norm=True)
        assert r.tag == qn('m:r')
        assert r.find(qn('m:t')) is not None
        assert r.find(qn('m:t')).text == 'x'

    def test_make_run_auto_style(self) -> None:
        r = _make_run('x', 'auto', norm=False)
        assert r.find(qn('m:rPr')) is None

    def test_make_run_props_unknown(self) -> None:
        rpr = _make_run_props('weird')
        sty = rpr.find(qn('m:sty'))
        assert sty is not None
        assert sty.get(qn('m:val')) == 'p'

    def test_token_to_omml_unknown_type(self) -> None:
        el = _token_to_omml(object())
        assert el.tag == qn('m:r')

    def test_token_to_omml_delim(self) -> None:
        el = _token_to_omml(DelimToken(char='\u0028', side='left'))
        assert el.tag == qn('m:r')

    def test_extract_limits_single_chars(self) -> None:
        sub, sup, _pos = _extract_limits('_{i=0}^{n} rest', 0)
        assert sub == [
            TextToken(text='i=', norm=False),
            TextToken(text='0', norm=True),
        ]
        assert sup == [TextToken(text='n', norm=False)]

    def test_extract_limits_plain(self) -> None:
        assert _extract_limits('x', 0) is None

    def test_extract_limits_sub_single(self) -> None:
        sub, sup, _pos = _extract_limits('_i x', 0)
        assert sub == [TextToken(text='i', norm=False)]
        assert sup is None

    def test_extract_sqrt_args_incomplete(self) -> None:
        assert _extract_sqrt_args('[3]', 0) is None

    def test_extract_sqrt_args_eof(self) -> None:
        assert _extract_sqrt_args('', 0) is None

    def test_extract_delim_unknown(self) -> None:
        assert _extract_delim('\\unknown', 0) is None

    def test_extract_delim_eof(self) -> None:
        assert _extract_delim('', 0) is None

    def test_extract_one_arg_no_brace(self) -> None:
        assert _extract_one_arg('abc', 0) is None

    def test_extract_one_arg_unbalanced(self) -> None:
        r = _extract_one_arg('{abc', 0)
        assert r is not None
        assert r[0] == 1
        assert r[1] == 4
        assert r[2] == 4

    def test_extract_one_arg_balanced(self) -> None:
        r = _extract_one_arg('{abc}', 0)
        assert r == (1, 4, 5)

    def test_extract_two_args_missing_second(self) -> None:
        assert _extract_two_args('{a}rest', 0) is None

    def test_extract_two_args_both(self) -> None:
        r = _extract_two_args('{a}{b}rest', 0)
        assert r == ((1, 2, 3), (4, 5, 6))

    def test_rpr_equivalent(self) -> None:
        assert _rpr_equivalent(None, None) is True
        assert _rpr_equivalent(None, 'x') is False
        assert _rpr_equivalent('x', None) is False


class TestMathStyleDialects:
    def test_office_keeps_strict_text_space_behavior(self) -> None:
        result = latex_to_omml(r'\text{hello world}')
        assert result is not None
        texts = [t.text or '' for t in result.findall('.//' + qn('m:t'))]
        assert ''.join(texts) == 'helloworld'

    def test_mathtype_preserves_text_spaces(self) -> None:
        result = latex_to_omml(r'\text{hello world}', style='mathtype')
        assert result is not None
        texts = [t.text or '' for t in result.findall('.//' + qn('m:t'))]
        assert ''.join(texts) == 'hello world'

    def test_mathtype_tilde_is_nbsp(self) -> None:
        result = latex_to_omml('a~b', style='mathtype')
        assert result is not None
        texts = [t.text or '' for t in result.findall('.//' + qn('m:t'))]
        assert ''.join(texts) == 'a b'

    def test_office_tilde_untouched(self) -> None:
        result = latex_to_omml('a~b')
        assert result is not None
        texts = [t.text or '' for t in result.findall('.//' + qn('m:t'))]
        assert 'a~b' in ''.join(texts)

    def test_mathtype_fraction_regression(self) -> None:
        result = latex_to_omml(r'\frac{a}{b}', style='mathtype')
        assert result is not None
        assert result.find('.//' + qn('m:f')) is not None

    def test_mathtype_mathrm_spaces(self) -> None:
        result = latex_to_omml(r'\mathrm{sin x}', style='mathtype')
        assert result is not None
        texts = [t.text or '' for t in result.findall('.//' + qn('m:t'))]
        assert ''.join(texts) == 'sin x'

    def test_unknown_style_falls_back_to_office(self) -> None:
        result = latex_to_omml(r'\text{a b}', style='unknown')
        assert result is not None


class TestGoldenSnapshot:
    def test_golden_snapshot_matches(self) -> None:
        if not GOLDEN_FILE.exists():
            import pytest

            pytest.skip('golden snapshot file not present')

        with GOLDEN_FILE.open('r', encoding='utf-8') as f:
            golden = json.load(f)

        for formula, old_xml in golden.items():
            result = latex_to_omml(formula)
            if old_xml is None:
                assert result is None, f'{formula!r} should be None'
                continue
            assert result is not None, f'{formula!r} should produce output'
            new_xml = etree.tostring(result, encoding='unicode')
            assert new_xml == old_xml, f'golden mismatch for {formula!r}'

    def test_operator_token_plain(self) -> None:
        tokens = parse(r'\sin x')
        assert isinstance(tokens[0], OperatorToken)
        assert tokens[0].op == 'sin'
        assert tokens[0].sub is None
        assert tokens[0].sup is None
