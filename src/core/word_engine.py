from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from src.core.document import DocumentABC
from src.rendering.styles import TextStyle

ALIGNMENT_MAP: dict[str, WD_ALIGN_PARAGRAPH] = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
    'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
}


class WordEngine(DocumentABC):

    def __init__(self, path: str | Path | None = None) -> None:
        super().__init__(path)
        self._doc: Document | None = None
        self._base_style: TextStyle = TextStyle.default_word()

    @property
    def doc(self) -> Document:
        if self._doc is None:
            raise RuntimeError('No document loaded. Call create() or open() first.')
        return self._doc

    def create(self) -> None:
        self._doc = Document()
        self._path = None
        self._base_style = TextStyle.default_word()

    def open(self, path: str | Path) -> None:
        self._path = Path(path)
        if not self._path.exists():
            raise FileNotFoundError(f'File not found: {self._path}')
        self._doc = Document(str(self._path))
        self._base_style = TextStyle.default_word()

    def save(self, path: str | Path | None = None) -> None:
        if path is not None:
            self._path = Path(path)
        if self._path is None:
            raise ValueError('No output path specified.')
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(self._path))

    def get_base_style(self) -> TextStyle:
        return self._base_style

    def add_text(
        self,
        text: str,
        bold: bool = False,
        italic: bool = False,
        font_name: str | None = None,
        font_size: int | None = None,
        color: str | None = None,
        alignment: str | None = None,
        text_style: TextStyle | None = None,
    ) -> Any:
        inline = TextStyle(
            bold=bold or None,
            italic=italic or None,
            font_name=font_name,
            font_size=font_size,
            color=color,
            alignment=alignment,
        )
        final = self._base_style
        if text_style is not None:
            final = final.merge(text_style)
        final = final.merge(inline)

        paragraph = self.doc.add_paragraph()
        run = paragraph.add_run(text)
        self._apply_run_style(run, final)

        if final.alignment and final.alignment in ALIGNMENT_MAP:
            paragraph.alignment = ALIGNMENT_MAP[final.alignment]

        return paragraph

    def add_styled_content(
        self,
        tokens: list[dict[str, Any]],
        base_style: TextStyle | None = None,
    ) -> Any:
        current_style = base_style if base_style else self._base_style

        for token in tokens:
            if token.get('command') == 'set_font':
                _apply_font_config(self, token)

        content_tokens = [
            t for t in tokens
            if not (t.get('command') == 'set_font' and not t.get('content'))
        ]
        has_visible = any(
            t.get('type') == 'text'
            or (t.get('command') and t.get('command') not in ('set_font',))
            for t in content_tokens
        )

        if not content_tokens or not has_visible:
            return None

        paragraph = self.doc.add_paragraph()

        for token in tokens:
            token_type = token.get('type', 'text')
            content = token.get('content', '')

            if not content and token_type == 'text':
                continue

            if token_type == 'text':
                run = paragraph.add_run(str(content))
                self._apply_run_style(run, current_style)

            elif token_type == 'command':
                cmd = token.get('command', '')
                if cmd == 'set_font':
                    _apply_font_config(self, token)
                    continue

                if cmd == 'newpage':
                    if paragraph.text.strip():
                        paragraph = self.doc.add_paragraph()
                    self.doc.add_page_break()
                    paragraph = self.doc.add_paragraph()
                    continue

                if cmd == 'includegraphics':
                    img_path = token.get('image', '')
                    if img_path:
                        try:
                            paragraph.add_run().add_picture(img_path)
                        except Exception:
                            run = paragraph.add_run(f'[Image: {img_path}]')
                            self._apply_run_style(run, current_style)
                    continue

                if cmd == 'heading':
                    level = token.get('level', 1)
                    h_content = token.get('content', '')
                    self.doc.add_heading(h_content, level=level)
                    paragraph = self.doc.add_paragraph()
                    continue

                if cmd == 'math':
                    latex = token.get('latex', '')
                    self._add_omath(paragraph, latex)
                    continue

                token_style = TextStyle.from_latex_token(token)
                merged = current_style.merge(token_style)

                inner_content = token.get('content', '')
                if inner_content:

                    if '\\' in inner_content:
                        from src.rendering.latex_parser import parse_structured
                        inner_tokens = parse_structured(inner_content)
                        self._render_tokens_inline(paragraph, inner_tokens, merged)
                    else:
                        run = paragraph.add_run(inner_content)
                        self._apply_run_style(run, merged)

        return paragraph

    def _render_tokens_inline(
        self,
        paragraph: Any,
        tokens: list[dict[str, Any]],
        base_style: TextStyle,
    ) -> None:
        current_style = base_style

        for token in tokens:
            token_type = token.get('type', 'text')
            content = token.get('content', '')

            if token_type == 'text':
                if content:
                    run = paragraph.add_run(str(content))
                    self._apply_run_style(run, current_style)

            elif token_type == 'command':
                cmd = token.get('command', '')
                if cmd in ('newpage', 'includegraphics', 'heading', 'set_font'):
                    continue

                if cmd == 'math':
                    latex = token.get('latex', '')
                    self._add_omath(paragraph, latex)
                    continue

                token_style = TextStyle.from_latex_token(token)
                merged = current_style.merge(token_style)

                inner_content = token.get('content', '')
                if inner_content:
                    if '\\' in inner_content:
                        from src.rendering.latex_parser import parse_structured
                        inner_tokens = parse_structured(inner_content)
                        self._render_tokens_inline(paragraph, inner_tokens, merged)
                    else:
                        run = paragraph.add_run(inner_content)
                        self._apply_run_style(run, merged)

    def add_latex_content(self, text: str) -> Any:
        from src.rendering.latex_parser import parse_structured
        tokens = parse_structured(text)
        return self.add_styled_content(tokens)

    def add_math_formula(self, latex: str) -> Any:
        paragraph = self.doc.add_paragraph()
        self._add_omath(paragraph, latex)
        return paragraph

    def _add_omath(self, paragraph: Any, latex: str) -> None:
        from src.rendering.math_omml import latex_to_omml

        omml = latex_to_omml(latex)
        if omml is not None:
            paragraph._p.append(omml)

    def _apply_run_style(self, run: Any, style: TextStyle) -> None:
        if style.font_name is not None:
            run.font.name = style.font_name

        if style.cjk_font_name is not None:
            r_pr = run._r.get_or_add_rPr()
            r_fonts = r_pr.find(qn('w:rFonts'))
            if r_fonts is None:
                r_fonts = OxmlElement('w:rFonts')
                r_pr.insert(0, r_fonts)
            r_fonts.set(qn('w:eastAsia'), style.cjk_font_name)

        if style.font_size is not None:
            run.font.size = Pt(style.font_size)
        if style.bold is not None and style.bold:
            run.bold = True
        if style.italic is not None and style.italic:
            run.italic = True
        if style.underline is not None and style.underline:
            run.underline = True
        if style.color is not None:
            try:
                run.font.color.rgb = RGBColor(
                    int(style.color[0:2], 16),
                    int(style.color[2:4], 16),
                    int(style.color[4:6], 16),
                )
            except (ValueError, IndexError):
                pass

    def replace_text(self, old: str, new: str, regex: bool = False) -> int:
        count = 0
        for paragraph in self.doc.paragraphs:
            if regex:
                if re.search(old, paragraph.text):
                    for run in paragraph.runs:
                        prev = run.text
                        run.text = re.sub(old, new, run.text)
                        if run.text != prev:
                            count += 1
            else:
                if old in paragraph.text:
                    for run in paragraph.runs:
                        if old in run.text:
                            run.text = run.text.replace(old, new)
                            count += 1
        return count

    def set_style(self, style_str: str) -> None:
        new_style = TextStyle.from_string(style_str)
        self._base_style = self._base_style.merge(new_style)

    def apply_style_to_all(self) -> None:
        for paragraph in self.doc.paragraphs:
            for run in paragraph.runs:
                self._apply_run_style(run, self._base_style)

        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            self._apply_run_style(run, self._base_style)

    def to_pdf(self, output_path: str | Path) -> None:
        self.save()
        try:
            from docx2pdf import convert
            convert(str(self._path), str(output_path))
        except ImportError:
            raise ImportError(
                'docx2pdf is required for PDF conversion. '
                'Install with: pip install tianshang-scribe[pdf]'
            )

    def add_heading(self, text: str, level: int = 1) -> Any:
        return self.doc.add_heading(text, level=level)

    def add_page_break(self) -> None:
        self.doc.add_page_break()

    def add_table(self, rows: int, cols: int) -> Any:
        return self.doc.add_table(rows=rows, cols=cols)

    def add_image(
        self, image_path: str, width: float | None = None, height: float | None = None
    ) -> Any:
        from docx.shared import Inches
        paragraph = self.doc.add_paragraph()
        run = paragraph.add_run()
        if width and height:
            return run.add_picture(image_path, width=Inches(width), height=Inches(height))
        elif width:
            return run.add_picture(image_path, width=Inches(width))
        elif height:
            return run.add_picture(image_path, height=Inches(height))
        return run.add_picture(image_path)

    def get_metadata(self) -> dict[str, str | None]:
        props = self.doc.core_properties
        return {
            'author': props.author,
            'title': props.title,
            'subject': props.subject,
            'category': props.category,
            'keywords': props.keywords,
            'comments': props.comments,
        }

    def set_metadata(self, **kwargs: str) -> None:
        props = self.doc.core_properties
        for key, value in kwargs.items():
            key_lower = key.lower()
            if key_lower == 'author':
                props.author = value
            elif key_lower == 'title':
                props.title = value
            elif key_lower == 'subject':
                props.subject = value
            elif key_lower == 'category':
                props.category = value
            elif key_lower == 'keywords':
                props.keywords = value
            elif key_lower == 'comments':
                props.comments = value

    def add_comment(self, text: str, range_start: int = 0, range_end: int = 0) -> None:
        paragraph = self.doc.paragraphs[0] if self.doc.paragraphs else self.doc.add_paragraph()
        paragraph.add_comment(text)

    def set_protection(self, password: str) -> None:
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        for section in self.doc.sections:
            protection = parse_xml(
                f'<w:documentProtection {nsdecls("w")} '
                f'w:edit="readOnly" w:enforcement="1" '
                f'w:cryptProviderType="rsaAES" '
                f'w:cryptAlgorithmClass="hash" '
                f'w:cryptAlgorithmType="typeAny" '
                f'w:cryptAlgorithmSid="14" '
                f'w:hash="placeholder"/>'
            )
            section._sectPr.append(protection)

    def unprotect(self) -> None:
        for section in self.doc.sections:
            elements = section._sectPr.findall(qn('w:documentProtection'))
            for el in elements:
                section._sectPr.remove(el)


def _apply_font_config(engine: 'WordEngine', token: dict[str, Any]) -> None:
    from src.rendering.styles import TextStyle

    role = token.get('role', '')
    font = token.get('font', '')
    if not font:
        return

    update = TextStyle()
    if 'CJK' in role:
        update.cjk_font_name = font
    else:
        update.font_name = font

    engine._base_style = engine._base_style.merge(update)
