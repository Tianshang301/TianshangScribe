from __future__ import annotations

from docx.oxml.ns import qn

from src.rendering.math_omml import _tokenize, latex_to_omml


class TestMathTokenize:
    def test_plain_text(self) -> None:
        tokens = _tokenize('x')
        assert len(tokens) == 1
        assert tokens[0] == {'type': 'text', 'text': 'x', 'norm': False}

    def test_superscript(self) -> None:
        tokens = _tokenize('x^2')
        assert len(tokens) == 2
        assert tokens[0]['type'] == 'text'
        assert tokens[0]['text'] == 'x'
        assert tokens[1]['type'] == 'sup'
        assert tokens[1]['content'] == '2'

    def test_superscript_braces(self) -> None:
        tokens = _tokenize('x^{n+1}')
        assert tokens[1]['type'] == 'sup'
        assert tokens[1]['content'] == 'n+1'

    def test_subscript(self) -> None:
        tokens = _tokenize('x_i')
        assert tokens[1]['type'] == 'sub'
        assert tokens[1]['content'] == 'i'

    def test_subscript_braces(self) -> None:
        tokens = _tokenize('x_{i+1}')
        assert tokens[1]['content'] == 'i+1'

    def test_sub_and_sup(self) -> None:
        tokens = _tokenize('x_i^2')
        types = [t['type'] for t in tokens]
        assert 'sub' in types
        assert 'sup' in types

    def test_fraction(self) -> None:
        tokens = _tokenize(r'\frac{a}{b}')
        assert tokens[0]['type'] == 'frac'
        assert tokens[0]['num'] == 'a'
        assert tokens[0]['den'] == 'b'

    def test_fraction_complex(self) -> None:
        tokens = _tokenize(r'\frac{x+y}{z}')
        assert tokens[0]['type'] == 'frac'
        assert tokens[0]['num'] == 'x+y'
        assert tokens[0]['den'] == 'z'

    def test_sqrt(self) -> None:
        tokens = _tokenize(r'\sqrt{x}')
        assert tokens[0]['type'] == 'sqrt'
        assert tokens[0]['degree'] == ''
        assert tokens[0]['content'] == 'x'

    def test_sqrt_nth(self) -> None:
        tokens = _tokenize(r'\sqrt[3]{x}')
        assert tokens[0]['degree'] == '3'
        assert tokens[0]['content'] == 'x'

    def test_greek_letters(self) -> None:
        tokens = _tokenize(r'\alpha \beta \Gamma')
        assert tokens[0]['text'] == '\u03b1'
        assert tokens[1]['text'] == '\u03b2'
        assert tokens[2]['text'] == '\u0393'

    def test_operator_sum(self) -> None:
        tokens = _tokenize(r'\sum_{i=0}^{n}')
        assert tokens[0]['type'] == 'nary'
        assert tokens[0]['op'] == 'sum'
        assert tokens[0]['sub'] == 'i=0'
        assert tokens[0]['sup'] == 'n'

    def test_operator_int(self) -> None:
        tokens = _tokenize(r'\int_{0}^{\infty}')
        assert tokens[0]['op'] == 'int'
        assert tokens[0]['sub'] == '0'
        assert tokens[0]['sup'] == r'\infty'

    def test_accent(self) -> None:
        tokens = _tokenize(r'\hat{x}')
        assert tokens[0]['type'] == 'accent'
        assert tokens[0]['accent'] == 'hat'
        assert tokens[0]['content'] == 'x'

    def test_accent_bar(self) -> None:
        tokens = _tokenize(r'\bar{y}')
        assert tokens[0]['accent'] == 'bar'

    def test_symbols(self) -> None:
        tokens = _tokenize(r'\infty \pm \times \div')
        assert tokens[0]['text'] == '\u221e'
        assert tokens[1]['text'] == '\u00b1'
        assert tokens[2]['text'] == '\u00d7'
        assert tokens[3]['text'] == '\u00f7'

    def test_digits_have_norm_true(self) -> None:
        tokens = _tokenize(r'\frac{42}{2a}')
        frac = tokens[0]
        assert frac['num'] == '42'
        inner = _tokenize('42')
        assert len(inner) == 1
        assert inner[0]['type'] == 'text'
        assert inner[0]['norm'] is True

    def test_mixed_text_splits_at_digit_boundary(self) -> None:
        tokens = _tokenize(r'\frac{4ac}{2}')
        frac = tokens[0]
        inner = _tokenize(frac['num'])
        assert len(inner) == 2
        assert inner[0] == {'type': 'text', 'text': '4', 'norm': True}
        assert inner[1] == {'type': 'text', 'text': 'ac', 'norm': False}

    def test_plain_letters_no_norm(self) -> None:
        tokens = _tokenize('abc')
        assert tokens[0] == {'type': 'text', 'text': 'abc', 'norm': False}

    def test_greek_uppercase_has_norm(self) -> None:
        tokens = _tokenize(r'\Gamma \Delta \Theta \alpha \beta')
        assert tokens[0]['norm'] is True
        assert tokens[1]['norm'] is True
        assert tokens[2]['norm'] is True
        assert tokens[3]['norm'] is False
        assert tokens[4]['norm'] is False

    def test_styled_mathrm_not_lost(self) -> None:
        tokens = _tokenize(r'\mathrm{abc}')
        assert len(tokens) == 1
        assert tokens[0]['type'] == 'styled'
        assert tokens[0]['style'] == 'normal'
        assert tokens[0]['content'] == 'abc'

    def test_styled_mathbf_not_lost(self) -> None:
        tokens = _tokenize(r'\mathbf{def}')
        assert tokens[0]['type'] == 'styled'
        assert tokens[0]['style'] == 'bold'
        assert tokens[0]['content'] == 'def'

    def test_collect_body_tokens_splits_digits(self) -> None:
        from src.rendering.math_omml import _collect_body_tokens

        tokens, _ = _collect_body_tokens('4ab', 0)
        assert tokens[0] == {'type': 'text', 'text': '4', 'norm': True}
        assert tokens[1] == {'type': 'text', 'text': 'ab', 'norm': False}


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
