from __future__ import annotations

from src.rendering.latex_parser import parse_latex_style, parse_structured


class TestParseLatexStyle:
    def test_plain_text_passthrough(self) -> None:
        result = parse_latex_style('Hello World')
        assert 'Hello World' in result

    def test_bfseries(self) -> None:
        result = parse_latex_style(r'\bfseries{bold text}')
        assert 'bold text' in result

    def test_itshape(self) -> None:
        result = parse_latex_style(r'\itshape{italic text}')
        assert 'italic text' in result

    def test_fontfamily_fontsize(self) -> None:
        result = parse_latex_style(r'\fontsize{18}{big text}')
        assert 'big text' in result

    def test_color(self) -> None:
        result = parse_latex_style(r'\color{FF0000}{red text}')
        assert 'red text' in result

    def test_nested_styles_does_not_crash(self) -> None:
        result = parse_latex_style(r'\bfseries{\itshape{important}}')
        assert 'important' in result

    def test_newpage(self) -> None:
        result = parse_latex_style(r'Before \newpage After')
        assert 'Before' in result
        assert 'After' in result

    def test_heading(self) -> None:
        result = parse_latex_style(r'\heading{2}{Section Title}')
        assert 'Section Title' in result

    def test_underline(self) -> None:
        result = parse_latex_style(r'\underline{underlined}')
        assert 'underlined' in result

    def test_mixed_content(self) -> None:
        result = parse_latex_style(
            r'Normal text. \bfseries{bold}, \itshape{italic}, and \color{0000FF}{blue}.'
        )
        assert 'Normal text' in result
        assert 'bold' in result
        assert 'italic' in result
        assert 'blue' in result


class TestParseStructured:
    def test_plain_text(self) -> None:
        tokens = parse_structured('Just text')
        assert len(tokens) == 1
        assert tokens[0]['type'] == 'text'
        assert tokens[0]['content'] == 'Just text'

    def test_bfseries_command(self) -> None:
        tokens = parse_structured(r'\bfseries{Hello}')
        assert len(tokens) == 1
        assert tokens[0]['type'] == 'command'
        assert tokens[0]['command'] == 'bfseries'
        assert 'style' in tokens[0]

    def test_color_command(self) -> None:
        tokens = parse_structured(r'\color{FF0000}{Red}')
        assert len(tokens) == 1
        assert tokens[0]['command'] == 'color'
        assert tokens[0]['color'] == 'FF0000'

    def test_heading_command(self) -> None:
        tokens = parse_structured(r'\heading{1}{Title}')
        assert len(tokens) == 1
        assert tokens[0]['command'] == 'heading'
        assert tokens[0]['level'] == 1
        assert tokens[0]['content'] == 'Title'

    def test_includegraphics(self) -> None:
        tokens = parse_structured(r'\includegraphics{image.png}')
        assert len(tokens) == 1
        assert tokens[0]['command'] == 'includegraphics'
        assert tokens[0]['image'] == 'image.png'

    def test_mixed_tokens(self) -> None:
        tokens = parse_structured(r'Text \bfseries{Bold} more \itshape{Italic}')
        assert len(tokens) >= 3
        assert tokens[0]['type'] == 'text'
        assert any(t.get('command') == 'bfseries' for t in tokens)
        assert any(t.get('command') == 'itshape' for t in tokens)

    def test_inline_math(self) -> None:
        tokens = parse_structured(r'Value $x = 1$ is shown')
        assert len(tokens) >= 2
        math_tokens = [t for t in tokens if t.get('command') == 'math']
        assert len(math_tokens) >= 1
        assert math_tokens[0]['latex'] == 'x = 1'
        assert math_tokens[0]['display'] is False

    def test_display_math(self) -> None:
        tokens = parse_structured(r'Formula: $$E=mc^2$$ is famous')
        math_tokens = [t for t in tokens if t.get('command') == 'math']
        assert len(math_tokens) == 1
        assert math_tokens[0]['latex'] == 'E=mc^2'
        assert math_tokens[0]['display'] is True

    def test_latex_style_parses_math(self) -> None:
        result = parse_latex_style(r'$x=y$ and text')
        assert 'Math: x=y' in result
        assert 'text' in result

    def test_frac_auto_detection(self) -> None:
        tokens = parse_structured(r'Formula \frac{a}{b} works')
        math_tokens = [t for t in tokens if t.get('command') == 'math']
        assert len(math_tokens) == 1
        assert r'\frac{a}{b}' in math_tokens[0]['latex']

    def test_lim_auto_detection(self) -> None:
        tokens = parse_structured(r'Limit \lim_{x \to 0} f(x)')
        math_tokens = [t for t in tokens if t.get('command') == 'math']
        assert len(math_tokens) == 1
        assert r'\lim_{x \to 0}' in math_tokens[0]['latex']

    def test_sqrt_auto_detection(self) -> None:
        tokens = parse_structured(r'\sqrt{a^2+b^2}')
        math_tokens = [t for t in tokens if t.get('command') == 'math']
        assert len(math_tokens) == 1
        assert r'\sqrt{a^2+b^2}' in math_tokens[0]['latex']

    def test_dollar_escape(self) -> None:
        tokens = parse_structured(r'Price is \$100')
        text_content = ''.join(t.get('content', '') for t in tokens if t['type'] == 'text')
        assert '$' in text_content
        assert '$100' in text_content

    def test_font_config_setmainfont(self) -> None:
        tokens = parse_structured(r'\setmainfont{Times New Roman}')
        font_tokens = [t for t in tokens if t.get('command') == 'set_font']
        assert len(font_tokens) == 1
        assert font_tokens[0]['role'] == 'setmainfont'
        assert font_tokens[0]['font'] == 'Times New Roman'

    def test_font_config_cjk(self) -> None:
        tokens = parse_structured(r'\setCJKmainfont{SimSun}')
        font_tokens = [t for t in tokens if t.get('command') == 'set_font']
        assert len(font_tokens) == 1
        assert font_tokens[0]['role'] == 'setCJKmainfont'
        assert font_tokens[0]['font'] == 'SimSun'
