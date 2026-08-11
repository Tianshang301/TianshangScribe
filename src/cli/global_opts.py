from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from src.core.document import DocumentType, detect_document_type
from src.utils.file_utils import check_overwrite as _check_overwrite


def parse_table_input(spec: str) -> list[list[str]]:
    """Parse a table spec: inline ``"H1,H2|a1,a2"`` or ``@file.csv``."""
    if spec.startswith('@'):
        import csv

        with open(spec[1:], newline='', encoding='utf-8') as f:
            return [list(row) for row in csv.reader(f)]
    return [[cell.strip() for cell in line.split(',')] for line in spec.split('|')]


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


REVERSE_SOURCE_TARGET: dict[str, DocumentType] = {
    '.md': DocumentType.WORD,
    '.markdown': DocumentType.WORD,
    '.html': DocumentType.WORD,
    '.htm': DocumentType.WORD,
    '.json': DocumentType.EXCEL,
}


def is_reverse_source(input_path: str | None) -> bool:
    if not input_path:
        return False
    return Path(input_path).suffix.lower() in REVERSE_SOURCE_TARGET


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
        target = REVERSE_SOURCE_TARGET.get(Path(input_path).suffix.lower())
        if target:
            return target
    return DocumentType.WORD


def determine_output_path(
    input_path: str | None,
    output_path: str | None,
    doc_type: DocumentType,
    to_pdf: bool = False,
    to_ext: str | None = None,
) -> str:
    if output_path:
        return output_path

    if input_path and input_path != '-':
        in_path = Path(input_path)
        stem = in_path.stem

        if to_pdf:
            return str(in_path.with_suffix('.pdf'))

        if to_ext:
            return str(in_path.with_name(f'{stem}-out{to_ext}'))

        suffix_map = {
            DocumentType.WORD: '.docx',
            DocumentType.EXCEL: '.xlsx',
            DocumentType.PPT: '.pptx',
        }
        ext = suffix_map.get(doc_type, '.docx')
        return str(in_path.parent / f'{stem}-out{ext}')

    if to_pdf:
        return 'output.pdf'

    if to_ext:
        return f'output{to_ext}'

    suffix_map = {
        DocumentType.WORD: 'output.docx',
        DocumentType.EXCEL: 'output.xlsx',
        DocumentType.PPT: 'output.pptx',
    }
    return suffix_map.get(doc_type, 'output.docx')


def check_overwrite(path: str, force: bool) -> bool:
    return _check_overwrite(path, force)
