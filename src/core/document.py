"""Document-agnostic interfaces, type detection, and engine factories."""

from __future__ import annotations

import enum
from abc import ABC, abstractmethod
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from src.rendering.styles import TextStyle


class DocumentType(enum.Enum):
    """Supported office document types."""

    WORD = 'word'
    EXCEL = 'excel'
    PPT = 'ppt'
    UNKNOWN = 'unknown'


EXTENSION_MAP: dict[str, DocumentType] = {
    '.docx': DocumentType.WORD,
    '.doc': DocumentType.WORD,
    '.xlsx': DocumentType.EXCEL,
    '.xlsm': DocumentType.EXCEL,
    '.pptx': DocumentType.PPT,
    '.ppt': DocumentType.PPT,
}


def detect_document_type(path: str | Path) -> DocumentType:
    """Return the document type for ``path`` based on its file extension."""
    suffix = Path(path).suffix.lower()
    return EXTENSION_MAP.get(suffix, DocumentType.UNKNOWN)


class DocumentABC(ABC):
    """Abstract base class implemented by Word, Excel, and PowerPoint engines."""

    def __init__(self, path: str | Path | None = None) -> None:
        """Store the optional document ``path``."""
        self._path: Path | None = Path(path) if path else None

    @property
    def path(self) -> Path | None:
        """Return the document path, or ``None`` if not set."""
        return self._path

    @abstractmethod
    def create(self) -> None:
        """Create a new blank document."""

    @abstractmethod
    def open(self, path: str | Path) -> None:
        """Load an existing document from ``path``."""

    @abstractmethod
    def save(self, path: str | Path | None = None) -> None:
        """Persist the document, optionally to ``path``."""

    @abstractmethod
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
        **kwargs: Any,
    ) -> Any:
        """Add a single paragraph of styled ``text`` to the document."""

    @abstractmethod
    def add_styled_content(
        self,
        tokens: list[dict[str, Any]],
    ) -> Any:
        """Add content built from parsed LaTeX-style ``tokens``."""

    @abstractmethod
    def replace_text(self, old: str, new: str, regex: bool = False) -> int:
        """Replace ``old`` with ``new``, optionally treating ``old`` as a regex."""

    @abstractmethod
    def set_style(self, style_str: str) -> None:
        """Merge the given ``style_str`` into the document's base style."""

    @abstractmethod
    def apply_style_to_all(self) -> None:
        """Apply the current base style to every existing element."""

    @abstractmethod
    def get_base_style(self) -> TextStyle:
        """Return the current base style."""

    @abstractmethod
    def to_pdf(self, output_path: str | Path) -> None:
        """Convert the document to PDF at ``output_path``."""

    def export_csv(self, output_path: str | Path) -> None:
        """Export the document as CSV (only supported for Excel documents)."""
        raise NotImplementedError('CSV export is only supported for Excel documents')

    def export_json(self, output_path: str | Path) -> None:
        """Export the document as JSON (only supported for Excel documents)."""
        raise NotImplementedError('JSON export is only supported for Excel documents')

    def export_html(self, output_path: str | Path) -> None:
        """Export the document as HTML (only supported for Excel documents)."""
        raise NotImplementedError('HTML export is only supported for Excel documents')

    @abstractmethod
    def get_metadata(self) -> dict[str, str | None]:
        """Return document metadata such as author and title."""

    @abstractmethod
    def set_metadata(self, **kwargs: str) -> None:
        """Set document metadata fields from keyword arguments."""


def create_document(doc_type: DocumentType) -> DocumentABC:
    """Create a new blank document of the given ``doc_type`` and return its engine."""
    from src.core.excel_engine import ExcelEngine
    from src.core.ppt_engine import PptEngine
    from src.core.word_engine import WordEngine

    engine_map: dict[DocumentType, type[DocumentABC]] = {
        DocumentType.WORD: WordEngine,
        DocumentType.EXCEL: ExcelEngine,
        DocumentType.PPT: PptEngine,
    }
    engine_cls = engine_map.get(doc_type)
    if engine_cls is None:
        raise ValueError(f'Unsupported document type: {doc_type}')
    engine = engine_cls()
    engine.create()
    return engine


def open_document(path: str | Path) -> DocumentABC:
    """Open the document at ``path`` and return the matching engine."""
    from src.core.excel_engine import ExcelEngine
    from src.core.ppt_engine import PptEngine
    from src.core.word_engine import WordEngine

    doc_type = detect_document_type(path)
    engine_map: dict[DocumentType, type[DocumentABC]] = {
        DocumentType.WORD: WordEngine,
        DocumentType.EXCEL: ExcelEngine,
        DocumentType.PPT: PptEngine,
    }
    engine_cls = engine_map.get(doc_type)
    if engine_cls is None:
        raise ValueError(f'Cannot open file: unsupported format "{path}"')
    engine = engine_cls()
    engine.open(path)
    return engine
