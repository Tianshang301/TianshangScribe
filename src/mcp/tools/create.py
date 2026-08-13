"""create_office_document — Create Word/Excel/PPT with structured content."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from src.core.document import DocumentType, create_document
from src.mcp.errors import McpErrorCode, _make_content, error_response, success_response
from src.mcp.schemas import ContentBlock, ToolOptions, as_dict
from src.utils.file_utils import ensure_parent_dir


def create_office_document(
    format: Annotated[  # noqa: A002  # tool schema field name; renaming breaks MCP API
        str,
        Field(
            description=(
                'Document format:\n'
                '- "docx": Word document — reports, letters, contracts, proposals\n'
                '- "xlsx": Excel workbook — spreadsheets, data tables, charts\n'
                '- "pptx": PowerPoint — slides, presentations, pitch decks'
            ),
            examples=['docx', 'xlsx'],
        ),
    ],
    content: Annotated[list[ContentBlock], Field(description='Ordered list of content blocks.')],
    template_data: Annotated[
        dict[str, Any] | None,
        Field(description='Key-value pairs to fill {{placeholder}} in content.'),
    ] = None,
    output_path: Annotated[str, Field(description='Output file path.')] = '',
    style: Annotated[str | None, Field(description='Global document style.')] = None,
    metadata: Annotated[
        dict[str, str] | None,
        Field(description='Document metadata (title, author, etc.).'),
    ] = None,
    options: Annotated[ToolOptions | None, Field(description='Tool options.')] = None,
) -> dict[str, Any]:
    """Create a Word, Excel, or PowerPoint document."""
    content_blocks: list[dict[str, Any]] = as_dict(content)
    opts: dict[str, Any] = as_dict(options) or {}
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

    if opts.get('dry_run'):
        return success_response(
            {
                'dry_run': True,
                'planned_content': len(content_blocks),
                'planned_items': [c.get('type', 'paragraph') for c in content_blocks],
            }
        )

    try:
        engine = create_document(doc_type)

        if style:
            engine.set_style(style)

        for item in content_blocks:
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

        ensure_parent_dir(output_path)
        engine.save(output_path)
        stats = _get_stats(
            engine,
            getattr(engine, 'doc', None)
            or getattr(engine, 'wb', None)
            or getattr(engine, 'prs', None),
        )

        return success_response(
            {
                'output_path': output_path,
                'format': format,
                'content_items': len(content_blocks),
                **stats,
            },
            content=_make_content(output_path, f'Document created: {output_path}'),
        )
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
    except Exception:  # noqa: S110  # introspection of library objects is best-effort; return {} on failure
        pass
    return {}
