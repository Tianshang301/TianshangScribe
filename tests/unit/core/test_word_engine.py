"""Unit tests for the Word engine."""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from src.core.document import open_document
from src.core.word_engine import WordEngine


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
