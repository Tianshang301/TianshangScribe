"""Unit tests for the Word engine."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import ClassVar

import pytest

from tianshang_scribe.core.document import open_document
from tianshang_scribe.core.word_engine import WordEngine


class TestWordEngine:
    @pytest.fixture
    def engine(self) -> WordEngine:
        e = WordEngine()
        e.create()
        return e

    def test_create_empty(self, engine: WordEngine) -> None:
        assert len(engine.doc.paragraphs) == 0

    def test_add_text(self, engine: WordEngine) -> None:
        engine.add_text('Hello World')
        assert engine.doc.paragraphs[0].text == 'Hello World'

    def test_add_text_with_style(self, engine: WordEngine) -> None:
        engine.add_text('Bold Text', bold=True, italic=True, font_size=14)
        paragraph = engine.doc.paragraphs[0]
        run = paragraph.runs[0]
        assert run.bold is True
        assert run.italic is True
        assert run.font.size is not None

    def test_add_text_multiple(self, engine: WordEngine) -> None:
        engine.add_text('Line 1')
        engine.add_text('Line 2')
        engine.add_text('Line 3')
        assert len(engine.doc.paragraphs) == 3

    def test_replace_text_exact(self, engine: WordEngine) -> None:
        engine.add_text('Hello World')
        count = engine.replace_text('World', 'Earth')
        assert count >= 1
        assert 'Earth' in engine.doc.paragraphs[0].text

    def test_replace_text_no_match(self, engine: WordEngine) -> None:
        engine.add_text('Hello World')
        count = engine.replace_text('Mars', 'Earth')
        assert count == 0

    def test_replace_with_regex(self, engine: WordEngine) -> None:
        engine.add_text('Hello 123 World')
        count = engine.replace_text(r'\d+', 'XXX', regex=True)
        assert count >= 1

    def test_save_and_reopen(self, engine: WordEngine) -> None:
        engine.add_text('Persist Test')
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / 'test.docx'
            engine.save(path)
            assert path.exists()

            reopened = open_document(path)
            assert isinstance(reopened, WordEngine)
            paragraphs = reopened.doc.paragraphs
            assert len(paragraphs) >= 1

    def test_open_nonexistent_raises(self) -> None:
        engine = WordEngine()
        with pytest.raises(FileNotFoundError):
            engine.open('nonexistent_file.docx')

    def test_get_metadata(self, engine: WordEngine) -> None:
        metadata = engine.get_metadata()
        assert 'author' in metadata
        assert 'title' in metadata

    def test_set_metadata(self, engine: WordEngine) -> None:
        engine.set_metadata(author='Test Author', title='Test Title')
        metadata = engine.get_metadata()
        assert metadata['author'] == 'Test Author'
        assert metadata['title'] == 'Test Title'

    def test_set_style(self, engine: WordEngine) -> None:
        engine.set_style('font=Arial,size=14')
        assert engine._base_style.font_name == 'Arial'
        assert engine._base_style.font_size == 14

    def test_get_base_style(self, engine: WordEngine) -> None:
        style = engine.get_base_style()
        assert style.font_name == 'Times New Roman'
        assert style.font_size == 12

    def test_add_text_inherits_base_style(self, engine: WordEngine) -> None:
        engine.set_style('font=Courier New,size=18,bold')
        engine.add_text('Styled text')
        paragraph = engine.doc.paragraphs[0]
        run = paragraph.runs[0]
        assert run.font.name == 'Courier New'
        assert run.bold is True

    def test_add_latex_content(self, engine: WordEngine) -> None:
        engine.add_latex_content(r'\bfseries{bold} and \itshape{italic}')
        text_content = engine.doc.paragraphs[0].text
        assert 'bold' in text_content
        assert 'italic' in text_content

    def test_add_styled_content_basic(self, engine: WordEngine) -> None:
        engine.set_style('font=Arial')
        engine.add_styled_content(
            [
                {'type': 'text', 'content': 'Hello '},
                {'type': 'command', 'command': 'bfseries', 'content': 'Bold'},
            ]
        )
        paragraph = engine.doc.paragraphs[0]
        assert 'Hello' in paragraph.text
        assert 'Bold' in paragraph.text

    def test_add_toc(self, engine: WordEngine) -> None:
        engine.add_text('Intro')
        engine.add_toc()
        full_text = '\n'.join(p.text for p in engine.doc.paragraphs)
        assert 'TOC' in full_text or len(engine.doc.paragraphs) >= 2

    def test_add_section_break(self, engine: WordEngine) -> None:
        engine.add_text('Sec1')
        engine.add_section_break()
        assert len(engine.doc.sections) >= 2

    def test_set_header(self, engine: WordEngine) -> None:
        engine.set_header('My Header')
        section = engine.doc.sections[0]
        assert section.header.paragraphs[0].text == 'My Header'

    def test_set_footer(self, engine: WordEngine) -> None:
        engine.set_footer('Page X')
        section = engine.doc.sections[0]
        assert section.footer.paragraphs[0].text == 'Page X'

    def test_add_watermark(self, engine: WordEngine) -> None:
        engine.add_watermark('DRAFT')
        section = engine.doc.sections[0]
        assert section.header.paragraphs[0].text == 'DRAFT'

    def test_small_caps_via_latex(self, engine: WordEngine) -> None:
        engine.add_latex_content(r'\scshape{Small Caps Text}')
        text = engine.doc.paragraphs[0].text
        assert 'Small Caps Text' in text


class TestWordEngineEdge:
    @pytest.fixture
    def engine(self) -> WordEngine:
        e = WordEngine()
        e.create()
        return e

    def test_doc_property_unloaded(self) -> None:
        e = WordEngine()
        with pytest.raises(RuntimeError):
            _ = e.doc

    def test_save_without_path(self, engine: WordEngine) -> None:
        with pytest.raises(ValueError):
            engine.save()

    def test_add_text_with_style_object(self, engine: WordEngine) -> None:
        from tianshang_scribe.rendering.styles import TextStyle

        style = TextStyle(font_name='Arial', font_size=16, alignment='center', color='FF0000')
        engine.add_text('Styled', text_style=style)
        p = engine.doc.paragraphs[0]
        assert p.alignment == 1
        assert p.runs[0].font.name == 'Arial'

    def test_add_text_full_args(self, engine: WordEngine) -> None:
        engine.add_text(
            'X',
            bold=True,
            italic=True,
            font_name='Courier',
            font_size=10,
            color='123456',
            alignment='right',
        )
        run = engine.doc.paragraphs[0].runs[0]
        assert run.bold is True
        assert run.italic is True
        assert run.font.name == 'Courier'

    def test_add_styled_content_empty_returns_none(self, engine: WordEngine) -> None:
        assert engine.add_styled_content([]) is None
        assert (
            engine.add_styled_content([{'type': 'command', 'command': 'set_font', 'font': ''}])
            is None
        )

    def test_add_styled_content_newpage(self, engine: WordEngine) -> None:
        engine.add_styled_content(
            [
                {'type': 'text', 'content': 'before'},
                {'type': 'command', 'command': 'newpage'},
            ]
        )

    def test_add_styled_content_includegraphics_missing(
        self, engine: WordEngine, tmp_path: Path
    ) -> None:
        engine.add_styled_content(
            [
                {
                    'type': 'command',
                    'command': 'includegraphics',
                    'image': str(tmp_path / 'nope.png'),
                },
            ]
        )
        text = '\n'.join(p.text for p in engine.doc.paragraphs)
        assert '[Image:' in text

    def test_add_styled_content_heading_cmd(self, engine: WordEngine) -> None:
        engine.add_styled_content(
            [{'type': 'command', 'command': 'heading', 'level': 2, 'content': 'Sub'}]
        )
        assert any(p.text == 'Sub' for p in engine.doc.paragraphs)

    def test_add_styled_content_paragraph_level(self, engine: WordEngine) -> None:
        engine.add_styled_content(
            [
                {'type': 'command', 'command': 'centering', 'content': 'Center Me'},
            ]
        )
        assert any(p.alignment == 1 for p in engine.doc.paragraphs)

    def test_add_styled_content_alignment_variants(self, engine: WordEngine) -> None:
        for cmd in ('raggedright', 'raggedleft'):
            engine.add_styled_content([{'type': 'command', 'command': cmd, 'content': 'T'}])
        engine.add_styled_content(
            [{'type': 'command', 'command': 'linespread', 'content': '1.5', 'content2': ''}]
        )

    def test_add_styled_content_indent(self, engine: WordEngine) -> None:
        engine.add_styled_content([{'type': 'command', 'command': 'indent', 'content': 'I'}])
        engine.add_styled_content([{'type': 'command', 'command': 'noindent', 'content': 'N'}])

    def test_add_styled_content_math_cmd(self, engine: WordEngine) -> None:
        engine.add_styled_content([{'type': 'command', 'command': 'math', 'latex': 'x+y'}])

    def test_add_styled_content_nested_latex(self, engine: WordEngine) -> None:
        engine.add_styled_content(
            [
                {'type': 'command', 'command': 'bfseries', 'content': '\\itshape{deep}'},
            ]
        )
        assert 'deep' in engine.doc.paragraphs[-1].text

    def test_render_tokens_inline(self, engine: WordEngine) -> None:
        from tianshang_scribe.rendering.styles import TextStyle

        paragraph = engine.doc.add_paragraph()
        engine._render_tokens_inline(
            paragraph,
            [
                {'type': 'text', 'content': 'a'},
                {'type': 'command', 'command': 'newpage'},
                {'type': 'command', 'command': 'math', 'latex': 'z'},
                {'type': 'command', 'command': 'bfseries', 'content': 'b'},
                {'type': 'command', 'command': 'itshape', 'content': '\\underline{c}'},
            ],
            TextStyle(),
        )
        assert 'a' in paragraph.text

    def test_add_omml_none(self, engine: WordEngine, monkeypatch) -> None:
        monkeypatch.setattr(
            'tianshang_scribe.rendering.math_omml.latex_to_omml', lambda s, **k: None
        )
        engine.add_math_formula('x')

    def test_set_math_font_sets_mathfont(self, engine: WordEngine) -> None:
        engine.set_math_font('Times New Roman')
        math_font = engine.doc.settings.element.find(
            './/{http://schemas.openxmlformats.org/officeDocument/2006/math}mathFont'
        )
        assert math_font is not None
        assert (
            math_font.get('{http://schemas.openxmlformats.org/officeDocument/2006/math}val')
            == 'Times New Roman'
        )

    def test_set_math_font_roundtrip(self, engine: WordEngine) -> None:
        engine.set_math_font('Latin Modern Math')
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            tmp = Path(f.name)
        try:
            engine.save(tmp)
            reopened = WordEngine()
            reopened.open(tmp)
            math_font = reopened.doc.settings.element.find(
                './/{http://schemas.openxmlformats.org/officeDocument/2006/math}mathFont'
            )
            assert math_font is not None
            assert (
                math_font.get('{http://schemas.openxmlformats.org/officeDocument/2006/math}val')
                == 'Latin Modern Math'
            )
        finally:
            Path(tmp).unlink(missing_ok=True)

    def test_set_math_font_empty_is_noop(self, engine: WordEngine) -> None:
        engine.set_math_font('   ')
        math_font = engine.doc.settings.element.find(
            './/{http://schemas.openxmlformats.org/officeDocument/2006/math}mathFont'
        )
        assert math_font is not None
        assert (
            math_font.get('{http://schemas.openxmlformats.org/officeDocument/2006/math}val')
            == 'Cambria Math'
        )

    def test_set_math_font_overrides_default(self, engine: WordEngine) -> None:
        engine.set_math_font('Times')
        engine.set_math_font('STIX Two Math')
        math_font = engine.doc.settings.element.find(
            './/{http://schemas.openxmlformats.org/officeDocument/2006/math}mathFont'
        )
        assert math_font is not None
        assert (
            math_font.get('{http://schemas.openxmlformats.org/officeDocument/2006/math}val')
            == 'STIX Two Math'
        )

    def test_apply_cjk_font(self, engine: WordEngine) -> None:
        from tianshang_scribe.rendering.styles import TextStyle

        engine.add_text('中文', text_style=TextStyle(cjk_font_name='SimSun'))
        r_pr = (
            engine.doc.paragraphs[0]
            .runs[0]
            ._r.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
        )
        assert r_pr is not None

    def test_apply_invalid_color(self, engine: WordEngine) -> None:
        engine.add_text('bad color', color='ZZZ')
        assert 'bad color' in engine.doc.paragraphs[0].text

    def test_replace_text_regex_run(self, engine: WordEngine) -> None:
        engine.add_text('abc 123')
        count = engine.replace_text(r'\d+', '9', regex=True)
        assert count >= 1
        assert '9' in engine.doc.paragraphs[0].text

    def test_apply_style_to_all_tables(self, engine: WordEngine) -> None:
        engine.add_text('hello')
        engine.add_table_data([['a', 'b']])
        engine.set_style('font=Arial')
        engine.apply_style_to_all()
        assert engine.doc.tables[0].rows[0].cells[0].paragraphs[0].runs[0].font.name == 'Arial'

    def test_to_pdf(self, engine: WordEngine, tmp_path: Path, monkeypatch) -> None:
        engine.add_text('hi')
        out = tmp_path / 'out.pdf'
        calls = []

        def fake_save(*a, **k) -> None:
            calls.append('save')

        def fake_w2p(s, d) -> None:
            calls.append((s, d))

        monkeypatch.setattr(engine, 'save', fake_save)
        monkeypatch.setattr('tianshang_scribe.transform.pdf.word_to_pdf', fake_w2p)
        engine.to_pdf(out)
        assert 'save' in calls

    def test_add_page_break(self, engine: WordEngine) -> None:
        engine.add_page_break()

    def test_add_table(self, engine: WordEngine) -> None:
        t = engine.add_table(2, 3)
        assert len(t.rows) == 2

    def test_add_table_data_single_row(self, engine: WordEngine) -> None:
        t = engine.add_table_data([['Only']])
        assert t.rows[0].cells[0].text == 'Only'

    def test_add_table_data_empty(self, engine: WordEngine) -> None:
        with pytest.raises(ValueError):
            engine.add_table_data([])

    def test_add_image_variants(self, engine: WordEngine, tmp_path: Path) -> None:
        from PIL import Image

        img = tmp_path / 'p.png'
        Image.new('RGB', (16, 16), (0, 0, 0)).save(img)
        engine.add_image(str(img), width=1.0)
        engine.add_image(str(img), height=1.0)
        engine.add_image(str(img))

    def test_set_metadata_all_keys(self, engine: WordEngine) -> None:
        engine.set_metadata(
            author='a', title='t', subject='s', category='c', keywords='k', comments='m'
        )
        md = engine.get_metadata()
        assert md == {
            'author': 'a',
            'title': 't',
            'subject': 's',
            'category': 'c',
            'keywords': 'k',
            'comments': 'm',
        }

    def test_add_comment(self, engine: WordEngine) -> None:
        engine.add_text('first')
        engine.add_comment('note')

    def test_add_comment_empty_doc(self, engine: WordEngine) -> None:
        engine.add_comment('note')

    def test_set_protection(self, engine: WordEngine) -> None:
        engine.set_protection('pw')
        assert (
            engine.doc.sections[0]._sectPr.find(
                './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}documentProtection'
            )
            is not None
        )

    def test_unprotect(self, engine: WordEngine) -> None:
        engine.set_protection('pw')
        engine.unprotect()
        assert (
            engine.doc.sections[0]._sectPr.find(
                './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}documentProtection'
            )
            is None
        )

    def test_merge_workbooks(self, engine: WordEngine, tmp_path: Path) -> None:
        other = tmp_path / 'other.docx'
        other_engine = WordEngine()
        other_engine.create()
        other_engine.add_text('merged content')
        other_engine.save(other)
        engine.add_text('base')
        engine.merge_workbooks([str(other)])
        assert 'merged content' in engine.extract_text()

    def test_set_header_existing_paragraph(self, engine: WordEngine) -> None:
        engine.set_header('H1')
        engine.set_header('H2')
        assert engine.doc.sections[0].header.paragraphs[0].text == 'H1H2'

    def test_set_footer_existing_paragraph(self, engine: WordEngine) -> None:
        engine.set_footer('F1')
        engine.set_footer('F2')
        assert engine.doc.sections[0].footer.paragraphs[0].text == 'F1F2'

    def test_add_watermark_empty_doc(self, engine: WordEngine) -> None:
        engine.add_watermark('Confidential')
        assert engine.doc.sections[0].header.paragraphs[0].text == 'Confidential'

    def test_clear_formats(self, engine: WordEngine) -> None:
        engine.add_text('bold', bold=True, font_size=20, color='FF0000')
        engine.clear_formats()
        run = engine.doc.paragraphs[0].runs[0]
        assert run.bold is None
        assert run.font.size is None

    def test_clear_content(self, engine: WordEngine) -> None:
        engine.add_text('to clear')
        engine.clear_content()
        assert engine.doc.paragraphs[0].text == ''

    def test_extract_text_empty_cells(self, engine: WordEngine) -> None:
        engine.add_table_data([['a', 'b']])
        engine.doc.tables[0].rows[0].cells[1].paragraphs[0].add_run('')
        text = engine.extract_text()
        assert 'a | b' in text or 'a' in text

    def test_apply_font_config(self, engine: WordEngine) -> None:
        from tianshang_scribe.core.word_engine import _apply_font_config

        _apply_font_config(engine, {'role': 'CJK', 'font': 'KaiTi'})
        assert engine._base_style.cjk_font_name == 'KaiTi'
        _apply_font_config(engine, {'role': 'latin', 'font': 'Arial'})
        assert engine._base_style.font_name == 'Arial'
        _apply_font_config(engine, {'role': 'latin', 'font': ''})
        assert engine._base_style.font_name == 'Arial'

    def test_add_styled_content_empty_text_token(self, engine: WordEngine) -> None:
        engine.add_styled_content(
            [{'type': 'text', 'content': ''}, {'type': 'text', 'content': 'v'}]
        )
        assert 'v' in engine.doc.paragraphs[0].text

    def test_add_styled_content_set_font_with_content(self, engine: WordEngine) -> None:
        engine.add_styled_content(
            [
                {'type': 'command', 'command': 'set_font', 'role': 'latin', 'font': 'Courier'},
                {'type': 'text', 'content': 'fixed'},
            ]
        )
        assert engine._base_style.font_name == 'Courier'

    def test_add_styled_content_newpage_empty(self, engine: WordEngine) -> None:
        engine.add_styled_content(
            [
                {'type': 'command', 'command': 'newpage'},
                {'type': 'command', 'command': 'newpage'},
            ]
        )

    def test_add_styled_content_image_ok(self, engine: WordEngine, tmp_path: Path) -> None:
        from PIL import Image

        img = tmp_path / 'ok.png'
        Image.new('RGB', (8, 8), (0, 0, 255)).save(img)
        engine.add_styled_content(
            [{'type': 'command', 'command': 'includegraphics', 'image': str(img)}]
        )
        assert len(engine.doc.inline_shapes) >= 1

    def test_add_styled_content_indent_nested(self, engine: WordEngine) -> None:
        engine.add_styled_content(
            [
                {'type': 'command', 'command': 'indent', 'content': '\\bfseries{in}'},
            ]
        )
        assert any('in' in p.text for p in engine.doc.paragraphs)

    def test_add_heading(self, engine: WordEngine) -> None:
        h = engine.add_heading('Title', level=1)
        assert h.text == 'Title'

    def test_add_image_width_height(self, engine: WordEngine, tmp_path: Path) -> None:
        from PIL import Image

        img = tmp_path / 'wh.png'
        Image.new('RGB', (8, 8), (0, 255, 0)).save(img)
        engine.add_image(str(img), width=1.0, height=1.0)

    def test_add_image_only_width(self, engine: WordEngine, tmp_path: Path) -> None:
        from PIL import Image

        img = tmp_path / 'w.png'
        Image.new('RGB', (8, 8), (0, 0, 0)).save(img)
        engine.add_image(str(img), width=1.0)

    def test_clear_links(self, engine: WordEngine) -> None:
        engine.add_text('x')
        engine.clear_links()

    def test_image_parts_skip_unreadable(self, engine: WordEngine, monkeypatch) -> None:
        class _BadPart:
            reltype = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
            is_external = False

            @property
            def target_part(self) -> None:
                raise OSError('unreadable')

        class _FakePart:
            rels: ClassVar[dict[str, object]] = {'rId1': _BadPart()}

        class _FakeDoc:
            part = _FakePart()

        engine._doc = _FakeDoc()  # type: ignore[assignment]
        assert engine._image_parts() == []


class TestWordExtract:
    @pytest.fixture
    def engine(self, tmp_path: Path) -> WordEngine:
        e = WordEngine()
        e.create()
        e.add_text('Hello World')
        e.add_table_data([['Name', 'City'], ['Alice', 'NYC']])
        from PIL import Image

        img = tmp_path / 'pixel.png'
        Image.new('RGB', (32, 32), (200, 40, 40)).save(img)
        e.add_image(str(img))
        return e

    def test_extract_text_contains_paragraph(self, engine: WordEngine) -> None:
        assert 'Hello World' in engine.extract_text()

    def test_extract_text_contains_table_row(self, engine: WordEngine) -> None:
        assert 'Alice | NYC' in engine.extract_text()

    def test_extract_tables(self, engine: WordEngine) -> None:
        tables = engine.extract_tables()
        assert tables == [[['Name', 'City'], ['Alice', 'NYC']]]

    def test_extract_images(self, engine: WordEngine, tmp_path: Path) -> None:
        saved = engine.extract_images(str(tmp_path / 'imgs'))
        assert len(saved) == 1
        assert saved[0].exists()

    def test_extract_structure(self, engine: WordEngine) -> None:
        struct = engine.extract_structure()
        assert struct['tables'] == 1
        assert struct['images'] == 1
        assert struct['sections'] == 1
