"""Unit tests for the PowerPoint engine."""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from src.core.document import open_document
from src.core.ppt_engine import PptEngine


class TestPptEngine:
    @pytest.fixture
    def engine(self) -> PptEngine:
        e = PptEngine()
        e.create()
        return e

    def test_create_has_slides(self, engine: PptEngine) -> None:
        assert len(engine.prs.slides) >= 0

    def test_add_slide(self, engine: PptEngine) -> None:
        initial = len(engine.prs.slides)
        engine.add_slide()
        assert len(engine.prs.slides) == initial + 1

    def test_save_and_reopen(self, engine: PptEngine) -> None:
        engine.add_slide()
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / 'test.pptx'
            engine.save(path)
            assert path.exists()
            reopened = open_document(path)
            assert isinstance(reopened, PptEngine)

    def test_delete_slide(self, engine: PptEngine) -> None:
        engine.add_slide()
        engine.add_slide()
        engine.add_slide()
        initial = len(engine.prs.slides)
        engine.delete_slide(0)
        assert len(engine.prs.slides) == initial - 1

    def test_move_slide(self, engine: PptEngine) -> None:
        engine.add_slide()
        engine.add_slide()
        engine.add_slide()
        engine.move_slide(0, 1)
        assert len(engine.prs.slides) == 3

    def test_add_notes(self, engine: PptEngine) -> None:
        engine.add_slide()
        engine.add_notes(0, 'Speaker note')
        notes = engine.prs.slides[0].notes_slide
        assert notes.notes_text_frame.text == 'Speaker note'

    def test_set_transition(self, engine: PptEngine) -> None:
        engine.add_slide()
        engine.set_transition('fade')
        slide = engine.prs.slides[0]
        ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        trans = slide.element.find(f'{{{ns}}}transition')
        assert trans is not None
        assert trans.find(f'{{{ns}}}fade') is not None

    def test_set_transition_invalid_raises(self, engine: PptEngine) -> None:
        engine.add_slide()
        with pytest.raises(ValueError, match='Unsupported transition'):
            engine.set_transition('nonexistent')

    def test_set_protection(self, engine: PptEngine) -> None:
        engine.set_protection('secret')
        ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        pres = engine.prs.part._element
        assert pres.find(f'{{{ns}}}modifyVerifier') is not None

    def test_unprotect(self, engine: PptEngine) -> None:
        engine.set_protection('secret')
        engine.unprotect()
        ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        pres = engine.prs.part._element
        assert pres.find(f'{{{ns}}}modifyVerifier') is None


class TestPptExtract:
    def test_extract_text(self) -> None:
        e = PptEngine()
        e.create()
        e.add_text('Slide one')
        assert '[slide 1] Slide one' in e.extract_text()

    def test_extract_structure(self) -> None:
        e = PptEngine()
        e.create()
        e.add_slide()
        struct = e.extract_structure()
        assert struct['slides'] == 1


class TestPptCompressMedia:
    def test_compress_media_reduces_size(self, tmp_path: Path) -> None:
        from PIL import Image
        from pptx import Presentation
        from pptx.util import Inches

        img = tmp_path / 'big.jpg'
        Image.new('RGB', (4000, 2000), (128, 128, 128)).save(img, 'JPEG', quality=95)
        pptx_path = tmp_path / 'deck.pptx'
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.add_picture(str(img), Inches(1), Inches(1), width=Inches(4))
        prs.save(str(pptx_path))
        original = pptx_path.stat().st_size

        e = PptEngine()
        e.open(str(pptx_path))
        saved = e.compress_media(max_dimension=1600, quality=60)
        assert saved > 0
        e.save(str(pptx_path))
        assert pptx_path.stat().st_size < original
