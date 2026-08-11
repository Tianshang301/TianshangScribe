from __future__ import annotations

from src.rendering.styles import TextStyle


class TestTextStyleFromString:
    def test_empty_string(self) -> None:
        style = TextStyle.from_string('')
        assert style.font_name is None
        assert style.font_size is None

    def test_font_and_size(self) -> None:
        style = TextStyle.from_string('font=Arial,size=14')
        assert style.font_name == 'Arial'
        assert style.font_size == 14

    def test_bool_flags(self) -> None:
        style = TextStyle.from_string('bold,italic,underline')
        assert style.bold is True
        assert style.italic is True
        assert style.underline is True

    def test_color(self) -> None:
        style = TextStyle.from_string('color=FF0000')
        assert style.color == 'FF0000'

    def test_color_with_hash(self) -> None:
        style = TextStyle.from_string('color=#00FF00')
        assert style.color == '00FF00'

    def test_aliases(self) -> None:
        style = TextStyle.from_string('font-family=Arial,font-size=18pt,align=center')
        assert style.font_name == 'Arial'
        assert style.font_size == 18
        assert style.alignment == 'center'

    def test_kv_and_flags_mixed(self) -> None:
        style = TextStyle.from_string('font=Times,bold,size=16,italic')
        assert style.font_name == 'Times'
        assert style.font_size == 16
        assert style.bold is True
        assert style.italic is True


class TestTextStyleMerge:
    def test_merge_fills_none(self) -> None:
        base = TextStyle(font_name='Times', font_size=12)
        override = TextStyle(bold=True)
        result = base.merge(override)
        assert result.font_name == 'Times'
        assert result.font_size == 12
        assert result.bold is True

    def test_merge_override(self) -> None:
        base = TextStyle(font_name='Times', font_size=12, bold=False)
        override = TextStyle(font_size=18, bold=True)
        result = base.merge(override)
        assert result.font_name == 'Times'
        assert result.font_size == 18
        assert result.bold is True

    def test_merge_chain(self) -> None:
        a = TextStyle(font_name='Arial')
        b = TextStyle(font_size=14, bold=True)
        c = TextStyle(color='FF0000', italic=True)
        result = a.merge(b, c)
        assert result.font_name == 'Arial'
        assert result.font_size == 14
        assert result.bold is True
        assert result.italic is True
        assert result.color == 'FF0000'

    def test_merge_ignores_none(self) -> None:
        base = TextStyle(font_name='Times', font_size=12)
        override = TextStyle()
        result = base.merge(override)
        assert result.font_name == 'Times'
        assert result.font_size == 12


class TestTextStyleFromLatexToken:
    def test_bfseries(self) -> None:
        token = {'type': 'command', 'command': 'bfseries', 'content': 'bold'}
        style = TextStyle.from_latex_token(token)
        assert style.bold is True

    def test_fontsize(self) -> None:
        token = {'type': 'command', 'command': 'fontsize', 'font_size': 24, 'content': 'big'}
        style = TextStyle.from_latex_token(token)
        assert style.font_size == 24

    def test_color(self) -> None:
        token = {'type': 'command', 'command': 'color', 'color': 'FF0000', 'content': 'red'}
        style = TextStyle.from_latex_token(token)
        assert style.color == 'FF0000'

    def test_style_map_token(self) -> None:
        token = {
            'type': 'command',
            'command': 'bfseries',
            'style': {'bold': True, 'italic': False},
            'content': '',
        }
        style = TextStyle.from_latex_token(token)
        assert style.bold is True


class TestTextStyleDefaults:
    def test_default_word(self) -> None:
        style = TextStyle.default_word()
        assert style.font_name == 'Times New Roman'
        assert style.font_size == 12

    def test_default_excel(self) -> None:
        style = TextStyle.default_excel()
        assert style.font_name == 'Calibri'
        assert style.font_size == 11

    def test_default_ppt(self) -> None:
        style = TextStyle.default_ppt()
        assert style.font_name == 'Calibri'
        assert style.font_size == 18


class TestTextStyleToCliString:
    def test_full_style(self) -> None:
        style = TextStyle(
            font_name='Times',
            font_size=14,
            bold=True,
            color='FF0000',
            alignment='center',
        )
        s = style.to_cli_string()
        assert 'font=Times' in s
        assert 'size=14' in s
        assert 'bold' in s
        assert 'color=FF0000' in s
        assert 'align=center' in s

    def test_minimal(self) -> None:
        style = TextStyle(font_name='Arial')
        s = style.to_cli_string()
        assert s == 'font=Arial'
