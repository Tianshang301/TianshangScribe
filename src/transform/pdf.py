"""PDF conversion — office2pdf (primary) + LibreOffice (fallback)."""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path
from typing import Optional


def _find_office2pdf() -> Optional[str]:
    """Find office2pdf binary on PATH."""
    return shutil.which('office2pdf')


def _find_libreoffice() -> Optional[str]:
    """Find LibreOffice/soffice binary on PATH."""
    for name in ['libreoffice', 'soffice']:
        lo = shutil.which(name)
        if lo:
            return lo
    for p in [
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    ]:
        if Path(p).exists():
            return p
    return None


def _convert_via_office2pdf(input_path: str | Path, output_path: str | Path) -> None:
    """Convert any OOXML document to PDF using office2pdf."""
    subprocess.run(
        ['office2pdf', str(input_path), '-o', str(output_path)],
        check=True,
        capture_output=True,
    )


def _convert_via_libreoffice(input_path: str | Path, output_path: str | Path) -> None:
    """Convert any OOXML document to PDF using LibreOffice headless."""
    lo_bin = _find_libreoffice()
    if not lo_bin:
        raise RuntimeError('No PDF engine found. Install office2pdf or LibreOffice.')
    output_dir = str(Path(output_path).parent)
    subprocess.run(
        [lo_bin, '--headless', '--convert-to', 'pdf',
         '--outdir', output_dir, str(input_path)],
        check=True,
        capture_output=True,
    )


def _convert(input_path: str | Path, output_path: str | Path) -> None:
    """Convert document to PDF — auto-select engine."""
    if _find_office2pdf():
        _convert_via_office2pdf(input_path, output_path)
        return

    if _find_libreoffice():
        _convert_via_libreoffice(input_path, output_path)
        return

    raise RuntimeError(
        'No PDF conversion engine found.\n'
        'Install office2pdf (recommended, ~2MB):\n'
        '  Download from https://github.com/XXXX/office2pdf/releases\n'
        'Or install LibreOffice (fallback, ~500MB):\n'
        '  https://www.libreoffice.org/download/'
    )


def word_to_pdf(docx_path: str | Path, pdf_path: str | Path) -> None:
    """Convert Word (.docx) to PDF."""
    _convert(docx_path, pdf_path)


def excel_to_pdf(xlsx_path: str | Path, pdf_path: str | Path) -> None:
    """Convert Excel (.xlsx) to PDF."""
    _convert(xlsx_path, pdf_path)


def ppt_to_pdf(pptx_path: str | Path, pdf_path: str | Path) -> None:
    """Convert PowerPoint (.pptx) to PDF."""
    _convert(pptx_path, pdf_path)


def word_to_markdown(docx_path: str | Path, md_path: str | Path) -> None:
    """Convert Word to Markdown using mammoth or pandoc."""
    try:
        import mammoth
        with open(docx_path, 'rb') as f:
            result = mammoth.convert_to_markdown(f)
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(result.value)
    except ImportError:
        subprocess.run(
            ['pandoc', str(docx_path), '-t', 'markdown', '-o', str(md_path)],
            check=True,
            capture_output=True,
        )


def word_to_html(docx_path: str | Path, html_path: str | Path) -> None:
    """Convert Word to HTML using mammoth or pandoc."""
    try:
        import mammoth
        with open(docx_path, 'rb') as f:
            result = mammoth.convert_to_html(f)
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(result.value)
    except ImportError:
        subprocess.run(
            ['pandoc', str(docx_path), '-t', 'html', '-o', str(html_path)],
            check=True,
            capture_output=True,
        )
