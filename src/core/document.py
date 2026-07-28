from __future__ import annotations

import enum
from abc import ABC, abstractmethod
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from src.rendering.styles import TextStyle


class DocumentType(enum.Enum):
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
    suffix = Path(path).suffix.lower()
    return EXTENSION_MAP.get(suffix, DocumentType.UNKNOWN)


class DocumentABC(ABC):

    def __init__(self, path: str | Path | None = None) -> None:
        self._path: Path | None = Path(path) if path else None

    @property
    def path(self) -> Path | None:
        return self._path

    @abstractmethod
    def create(self) -> None:
        ...

    @abstractmethod
    def open(self, path: str | Path) -> None:
        ...

    @abstractmethod
    def save(self, path: str | Path | None = None) -> None:
        ...

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
        text_style: 'TextStyle | None' = None,
    ) -> Any:
        ...

    @abstractmethod
    def add_styled_content(
        self,
        tokens: list[dict[str, Any]],
    ) -> Any:
        ...

    @abstractmethod
    def replace_text(self, old: str, new: str, regex: bool = False) -> int:
        ...

    @abstractmethod
    def set_style(self, style_str: str) -> None:
        ...

    @abstractmethod
    def apply_style_to_all(self) -> None:
        ...

    @abstractmethod
    def get_base_style(self) -> 'TextStyle':
        ...

    @abstractmethod
    def to_pdf(self, output_path: str | Path) -> None:
        ...

    @abstractmethod
    def get_metadata(self) -> dict[str, str | None]:
        ...

    @abstractmethod
    def set_metadata(self, **kwargs: str) -> None:
        ...


def create_document(doc_type: DocumentType) -> DocumentABC:
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
