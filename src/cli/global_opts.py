from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from src.core.document import DocumentType, detect_document_type


@dataclass
class GlobalOptions:
    doc_type: DocumentType | None = None
    output_path: str | None = None
    force: bool = False
    use_stdin: bool = False
    use_stdout: bool = False
    to_pdf: bool = False
    latex_style: bool = False
    regex: bool = False
    input_path: str | None = None


def resolve_doc_type(
    explicit: DocumentType | None,
    input_path: str | None,
) -> DocumentType:
    if explicit:
        return explicit
    if input_path and input_path != '-':
        detected = detect_document_type(input_path)
        if detected != DocumentType.UNKNOWN:
            return detected
    return DocumentType.WORD


def determine_output_path(
    input_path: str | None,
    output_path: str | None,
    doc_type: DocumentType,
    to_pdf: bool = False,
) -> str:
    if output_path:
        return output_path

    if input_path and input_path != '-':
        in_path = Path(input_path)
        stem = in_path.stem

        if to_pdf:
            return str(in_path.with_suffix('.pdf'))

        suffix_map = {
            DocumentType.WORD: '.docx',
            DocumentType.EXCEL: '.xlsx',
            DocumentType.PPT: '.pptx',
        }
        ext = suffix_map.get(doc_type, '.docx')
        return str(in_path.parent / f'{stem}-out{ext}')

    if to_pdf:
        return 'output.pdf'

    suffix_map = {
        DocumentType.WORD: 'output.docx',
        DocumentType.EXCEL: 'output.xlsx',
        DocumentType.PPT: 'output.pptx',
    }
    return suffix_map.get(doc_type, 'output.docx')


def check_overwrite(path: str, force: bool) -> bool:
    if Path(path).exists() and not force:
        return False
    return True
