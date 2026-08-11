"""Reverse conversions — HTML/Markdown → Word.

Both converters read the source with an explicit UTF-8 encoding and feed the
string to ``htmldocx``'s string API, avoiding ``parse_html_file``'s use of the
platform default encoding (GBK on Windows). ``table_style`` is intentionally
left at its default so python-docx's deprecated style-id lookup is not hit
(keeps the test suite clean under ``-W error``).
"""

from __future__ import annotations

from pathlib import Path


def _html_to_docx(html: str, docx_path: str | Path) -> None:
    from docx import Document
    from htmldocx import HtmlToDocx

    document = Document()
    converter = HtmlToDocx()
    converter.add_html_to_document(html, document)
    document.save(str(docx_path))


def html_to_word(html_path: str | Path, docx_path: str | Path) -> None:
    """Convert an HTML file to a Word document (via ``htmldocx``)."""
    html = Path(html_path).read_text(encoding='utf-8').lstrip('\ufeff')
    _html_to_docx(html, docx_path)


def markdown_to_word(md_path: str | Path, docx_path: str | Path) -> None:
    """Convert a Markdown file to a Word document (Markdown → HTML → docx)."""
    import markdown as markdown_lib

    text = Path(md_path).read_text(encoding='utf-8').lstrip('\ufeff')
    html = markdown_lib.markdown(
        text,
        extensions=['tables', 'fenced_code', 'nl2br'],
    )
    _html_to_docx(html, docx_path)
