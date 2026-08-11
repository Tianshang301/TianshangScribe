from __future__ import annotations

import contextlib
from dataclasses import dataclass, replace
from typing import Any


@dataclass
class TextStyle:
    font_name: str | None = None
    cjk_font_name: str | None = None
    font_size: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: str | None = None
    alignment: str | None = None
    small_caps: bool | None = None

    @classmethod
    def from_string(cls, s: str) -> TextStyle:
        instance = cls()
        if not s.strip():
            return instance

        for pair in s.split(','):
            pair = pair.strip()
            if not pair:
                continue
            if '=' in pair:
                key, _, value = pair.partition('=')
                key = key.strip().lower()
                value = value.strip()
                cls._apply_kv(instance, key, value)
            else:
                key = pair.strip().lower()
                cls._apply_flag(instance, key)

        return instance

    @staticmethod
    def _apply_kv(style: TextStyle, key: str, value: str) -> None:
        key = key.lower()
        if key in ('font', 'font_name', 'font-family'):
            style.font_name = value
        elif key in ('size', 'font_size', 'font-size'):
            with contextlib.suppress(ValueError):
                style.font_size = int(value.replace('pt', '').strip())
        elif key == 'bold':
            style.bold = value.lower() in ('true', '1', 'yes')
        elif key == 'italic':
            style.italic = value.lower() in ('true', '1', 'yes')
        elif key in ('underline', 'underlined'):
            style.underline = value.lower() in ('true', '1', 'yes')
        elif key in ('color', 'font_color', 'font-color'):
            style.color = value.strip('#')
        elif key in ('cjk-font', 'cjk_font_name', 'cjk-font-family'):
            style.cjk_font_name = value
        elif key in ('align', 'alignment'):
            style.alignment = value.lower()

    @staticmethod
    def _apply_flag(style: TextStyle, key: str) -> None:
        key = key.lower()
        if key == 'bold':
            style.bold = True
        elif key == 'italic':
            style.italic = True
        elif key in ('underline', 'underlined'):
            style.underline = True

    def merge(self, *others: TextStyle) -> TextStyle:
        result = replace(self)
        for other in others:
            if other.font_name is not None:
                result.font_name = other.font_name
            if other.cjk_font_name is not None:
                result.cjk_font_name = other.cjk_font_name
            if other.font_size is not None:
                result.font_size = other.font_size
            if other.bold is not None:
                result.bold = other.bold
            if other.italic is not None:
                result.italic = other.italic
            if other.underline is not None:
                result.underline = other.underline
            if other.color is not None:
                result.color = other.color
            if other.alignment is not None:
                result.alignment = other.alignment
            if other.small_caps is not None:
                result.small_caps = other.small_caps
        return result

    @classmethod
    def from_latex_token(cls, token: dict[str, Any]) -> TextStyle:
        instance = cls()

        cmd = token.get('command', '')

        if cmd in ('bfseries', 'bold'):
            instance.bold = True
        elif cmd == 'itshape':
            instance.italic = True
        elif cmd == 'scshape':
            instance.small_caps = True
        elif cmd == 'rmfamily':
            instance.font_name = 'Times New Roman'
        elif cmd == 'sffamily':
            instance.font_name = 'Arial'
        elif cmd == 'ttfamily':
            instance.font_name = 'Courier New'
        elif cmd == 'underline':
            instance.underline = True
        elif cmd == 'fontfamily':
            instance.font_name = token.get('font_name')
        elif cmd == 'fontsize':
            instance.font_size = token.get('font_size')
        elif cmd == 'color':
            instance.color = token.get('color', '').lstrip('#')

        if 'style' in token:
            style_dict = token['style']
            if 'bold' in style_dict:
                instance.bold = style_dict['bold']
            if 'italic' in style_dict:
                instance.italic = style_dict['italic']
            if 'font_name' in style_dict:
                instance.font_name = style_dict['font_name']

        return instance

    @classmethod
    def default_word(cls) -> TextStyle:
        return cls(font_name='Times New Roman', cjk_font_name='SimSun', font_size=12)

    @classmethod
    def default_excel(cls) -> TextStyle:
        return cls(font_name='Calibri', font_size=11)

    @classmethod
    def default_ppt(cls) -> TextStyle:
        return cls(font_name='Calibri', font_size=18)

    def to_cli_string(self) -> str:
        parts: list[str] = []
        if self.font_name:
            parts.append(f'font={self.font_name}')
        if self.cjk_font_name:
            parts.append(f'cjk-font={self.cjk_font_name}')
        if self.font_size:
            parts.append(f'size={self.font_size}')
        if self.bold:
            parts.append('bold')
        if self.italic:
            parts.append('italic')
        if self.underline:
            parts.append('underline')
        if self.color:
            parts.append(f'color={self.color}')
        if self.alignment:
            parts.append(f'align={self.alignment}')
        return ','.join(parts)
