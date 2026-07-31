"""create_office_document — Create Word/Excel/PPT with structured content."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from mcp.errors import McpErrorCode, _make_content, error_response, success_response
from src.core.document import DocumentType, create_document


def create_office_document(
    format: str,
    content: list[dict[str, Any]],
    template_data: dict[str, Any] | None = None,
    output_path: str = '',
    style: str | None = None,
    metadata: dict[str, str] | None = None,
    options: dict[str, Any] | None = None,
) -> dict:
    """Create a Word, Excel, or PowerPoint document."""
    fmt = format.lower().replace('pptx', 'ppt').replace('ppt', 'ppt')
    fmt_map = {'docx': DocumentType.WORD, 'xlsx': DocumentType.EXCEL, 'ppt': DocumentType.PPT}
    doc_type = fmt_map.get(fmt)
    if doc_type is None:
        return error_response(
            McpErrorCode.UNSUPPORTED_FORMAT,
            f"'{format}' is not supported. Use docx, xlsx, or pptx.",
        )

    if not output_path:
        import tempfile
        ext = format.lower()
        output_path = os.path.join(tempfile.gettempdir(), f'scribe_output.{ext}')

    opts = options or {}
    if opts.get('dry_run'):
        return success_response({
            'dry_run': True,
            'planned_content': len(content),
            'planned_items': [c.get('type', 'paragraph') for c in content],
        })

    try:
        engine = create_document(doc_type)

        if style:
            engine.set_style(style)

        for item in content:
            item_type = item.get('type', 'paragraph')
            text = item.get('text', '')
            item_style = item.get('style')

            if item_type in ('paragraph', 'text'):
                if '\\' in text or '$' in text:
                    if hasattr(engine, 'add_latex_content'):
                        engine.add_latex_content(text)
                    else:
                        engine.add_text(text)
                else:
                    engine.add_text(text)
            elif item_type == 'heading':
                level = item.get('level', 1)
                if hasattr(engine, 'add_heading'):
                    engine.add_heading(text, level=level)
                elif hasattr(engine, 'add_latex_content'):
                    engine.add_latex_content(rf'\heading{{{level}}}{{{text}}}')
                else:
                    engine.add_text(text)
            elif item_type == 'formula':
                if hasattr(engine, 'add_math_formula'):
                    engine.add_math_formula(text)
                elif hasattr(engine, 'add_latex_content'):
                    engine.add_latex_content('$$' + text + '$$')
                else:
                    engine.add_text(text)
            elif item_type == 'page_break':
                if hasattr(engine, 'add_page_break'):
                    engine.add_page_break()
                elif hasattr(engine, 'add_latex_content'):
                    engine.add_latex_content(r'\newpage')
            elif item_type == 'table':
                rows = item.get('rows', [])
                if hasattr(engine, 'add_table'):
                    engine.add_table(rows=len(rows), cols=len(rows[0]) if rows else 1)
                else:
                    for row in rows:
                        engine.add_text(' | '.join(str(c) for c in row))
            elif item_type == 'image':
                img_path = item.get('path', '')
                if img_path and hasattr(engine, 'add_latex_content'):
                    engine.add_latex_content(rf'\includegraphics{{{img_path}}}')

            if item_style and hasattr(engine, 'set_style'):
                engine.set_style(item_style)

        if template_data and hasattr(engine, 'replace_text'):
            for key, value in template_data.items():
                engine.replace_text('{{' + key + '}}', str(value))

        if metadata and hasattr(engine, 'set_metadata'):
            engine.set_metadata(**metadata)

        if opts.get('backup') and Path(output_path).exists():
            backup = output_path + '.bak'
            import shutil
            shutil.copy2(output_path, backup)

        engine.save(output_path)
        stats = _get_stats(engine, engine.doc if hasattr(engine, 'doc')
                           else (engine.wb if hasattr(engine, 'wb') else engine.prs))

        return success_response({
            'output_path': output_path,
            'format': format,
            'content_items': len(content),
            **stats,
        }, content=_make_content(output_path, f'Document created: {output_path}'))
    except Exception as e:
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))


def _get_stats(engine: Any, obj: Any) -> dict[str, int]:
    try:
        if hasattr(obj, 'paragraphs'):
            return {
                'paragraphs': len(obj.paragraphs),
                'sections': len(obj.sections),
            }
        elif hasattr(obj, 'worksheets'):
            return {
                'sheets': len(obj.worksheets),
                'total_rows': sum(ws.max_row or 0 for ws in obj.worksheets),
            }
        elif hasattr(obj, 'slides'):
            return {
                'slides': len(obj.slides),
            }
    except Exception:
        pass
    return {}
