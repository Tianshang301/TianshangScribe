from __future__ import annotations

from pathlib import Path

import pytest

from src.core.document import DocumentType, create_document
from src.core.excel_engine import ExcelEngine
from src.core.ppt_engine import PptEngine
from src.core.word_engine import WordEngine


@pytest.fixture
def word_factory() -> type[WordEngine]:
    """Fresh WordEngine class accessor (not bound to any file)."""
    return WordEngine


@pytest.fixture
def excel_factory() -> type[ExcelEngine]:
    return ExcelEngine


@pytest.fixture
def ppt_factory() -> type[PptEngine]:
    return PptEngine


@pytest.fixture
def tmp_docx(tmp_path: Path) -> Path:
    """A freshly-created empty .docx file path."""
    engine = create_document(DocumentType.WORD)
    path = tmp_path / 'doc.docx'
    engine.save(path)
    return path


@pytest.fixture
def tmp_xlsx(tmp_path: Path) -> Path:
    """A freshly-created empty .xlsx file path."""
    engine = create_document(DocumentType.EXCEL)
    path = tmp_path / 'book.xlsx'
    engine.save(path)
    return path


@pytest.fixture
def tmp_pptx(tmp_path: Path) -> Path:
    """A freshly-created empty .pptx file path."""
    engine = create_document(DocumentType.PPT)
    path = tmp_path / 'deck.pptx'
    engine.save(path)
    return path


@pytest.fixture
def sample_template_docx(tmp_path: Path) -> Path:
    """A .docx template containing {{name}} / {{city}} and an {{#each}} loop."""
    engine = WordEngine()
    engine.create()
    engine.add_text('Hello {{name}} from {{city}}')
    engine.add_text('{{#each items}}Item: {{this}}{{/each}}')
    path = tmp_path / 'template.docx'
    engine.save(path)
    return path
