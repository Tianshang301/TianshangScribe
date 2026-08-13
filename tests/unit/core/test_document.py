"""Unit tests for document type detection and factory creation."""

from __future__ import annotations

import pytest

from src.core.document import DocumentType, create_document, detect_document_type
from src.core.excel_engine import ExcelEngine
from src.core.ppt_engine import PptEngine
from src.core.word_engine import WordEngine


class TestDetectDocumentType:
    def test_detect_docx(self) -> None:
        assert detect_document_type('test.docx') == DocumentType.WORD

    def test_detect_xlsx(self) -> None:
        assert detect_document_type('test.xlsx') == DocumentType.EXCEL

    def test_detect_pptx(self) -> None:
        assert detect_document_type('test.pptx') == DocumentType.PPT

    def test_detect_unknown(self) -> None:
        assert detect_document_type('test.txt') == DocumentType.UNKNOWN

    def test_detect_case_insensitive(self) -> None:
        assert detect_document_type('TEST.DOCX') == DocumentType.WORD


class TestCreateDocument:
    def test_create_word(self) -> None:
        engine = create_document(DocumentType.WORD)
        assert isinstance(engine, WordEngine)
        assert engine.doc is not None

    def test_create_excel(self) -> None:
        engine = create_document(DocumentType.EXCEL)
        assert isinstance(engine, ExcelEngine)
        assert engine.wb is not None

    def test_create_ppt(self) -> None:
        engine = create_document(DocumentType.PPT)
        assert isinstance(engine, PptEngine)
        assert engine.prs is not None

    def test_create_unknown_raises(self) -> None:
        with pytest.raises(ValueError, match='Unsupported'):
            create_document(DocumentType.UNKNOWN)
