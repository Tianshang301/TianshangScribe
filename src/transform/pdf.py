from __future__ import annotations

from pathlib import Path


def word_to_pdf(docx_path: str | Path, pdf_path: str | Path) -> None:
    try:
        from docx2pdf import convert
        convert(str(docx_path), str(pdf_path))
    except ImportError:
        _fallback_libreoffice(docx_path, pdf_path)


def excel_to_pdf(xlsx_path: str | Path, pdf_path: str | Path) -> None:
    _fallback_libreoffice(xlsx_path, pdf_path)


def ppt_to_pdf(pptx_path: str | Path, pdf_path: str | Path) -> None:
    _fallback_libreoffice(pptx_path, pdf_path)


def word_to_markdown(docx_path: str | Path, md_path: str | Path) -> None:
    try:
        import mammoth
        with open(docx_path, 'rb') as f:
            result = mammoth.convert_to_markdown(f)
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(result.value)
    except ImportError:
        _fallback_pandoc(docx_path, md_path, 'markdown')


def word_to_html(docx_path: str | Path, html_path: str | Path) -> None:
    try:
        import mammoth
        with open(docx_path, 'rb') as f:
            result = mammoth.convert_to_html(f)
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(result.value)
    except ImportError:
        _fallback_pandoc(docx_path, html_path, 'html')


def _fallback_libreoffice(input_path: str | Path, output_path: str | Path) -> None:
    import shutil
    import subprocess

    lo_paths = [
        'libreoffice',
        'soffice',
        '/usr/bin/libreoffice',
        '/usr/bin/soffice',
        r'C:\Program Files\LibreOffice\program\soffice.exe',
    ]
    lo_bin = None
    for p in lo_paths:
        if shutil.which(p):
            lo_bin = p
            break

    if not lo_bin:
        raise RuntimeError(
            'LibreOffice is required for this conversion. '
            'Install from https://www.libreoffice.org/download/'
        )

    output_dir = str(Path(output_path).parent)
    subprocess.run(
        [lo_bin, '--headless', '--convert-to', 'pdf',
         '--outdir', output_dir, str(input_path)],
        check=True,
        capture_output=True,
    )


def _fallback_pandoc(input_path: str | Path, output_path: str | Path, fmt: str) -> None:
    import subprocess

    subprocess.run(
        ['pandoc', str(input_path), '-t', fmt, '-o', str(output_path)],
        check=True,
        capture_output=True,
    )
