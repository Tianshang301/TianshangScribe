"""create_office_document — Create Word/Excel/PPT with structured content."""

from __future__ import annotations

import contextlib
import os
from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from tianshang_scribe.core.document import DocumentType, create_document
from tianshang_scribe.mcp.errors import (
    McpErrorCode,
    _make_content,
    error_response,
    success_response,
)
from tianshang_scribe.mcp.schemas import ContentBlock, ToolOptions, as_dict
from tianshang_scribe.mcp.tools._parse import (
    parse_conditional_format,
    parse_data_validation,
    parse_number_format,
    parse_ppt_chart,
    resolve_slide_index,
)
from tianshang_scribe.utils.file_utils import ensure_parent_dir


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

        current_slide_index: int | None = None

        for item in content_blocks:
            item_type = item.get('type', 'paragraph')
            text = item.get('text', '')
            item_style = item.get('style')

            if item_type in ('paragraph', 'text'):
                if doc_type == DocumentType.PPT:
                    if current_slide_index is None:
                        engine.add_text(text)
                        current_slide_index = len(engine.prs.slides) - 1 if hasattr(engine, 'prs') else 0
                    else:
                        engine.add_text(text, slide_index=current_slide_index)
                elif '\\' in text or '$' in text:
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
                if doc_type == DocumentType.PPT and hasattr(engine, 'add_slide'):
                    engine.add_slide()
                    current_slide_index = len(engine.prs.slides) - 1 if hasattr(engine, 'prs') else 0
                elif hasattr(engine, 'add_page_break'):
                    engine.add_page_break()
                elif hasattr(engine, 'add_latex_content'):
                    engine.add_latex_content(r'\newpage')
            elif item_type == 'table':
                rows = item.get('rows', []) or []
                if doc_type == DocumentType.PPT and rows and hasattr(engine, 'add_table'):
                    col_names = rows[0]
                    data = rows[1:]
                    idx = _resolve_ppt_slide(engine, item, current_slide_index)
                    engine.add_table(idx, data, col_names=col_names)
                elif hasattr(engine, 'add_table'):
                    engine.add_table(rows=len(rows), cols=len(rows[0]) if rows else 1)
                else:
                    for row in rows:
                        engine.add_text(' | '.join(str(c) for c in row))
            elif item_type == 'image':
                img_path = item.get('path', '')
                if img_path and hasattr(engine, 'add_latex_content'):
                    engine.add_latex_content(rf'\includegraphics{{{img_path}}}')

            # Excel/PPT capability fields are independent of the block type.
            _apply_mcp_capabilities(engine, doc_type, item, current_slide_index)

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


def _resolve_ppt_slide(engine: Any, item: dict[str, Any], current_slide_index: int | None) -> int:
    """Resolve the target PPT slide for a block.

    An explicit ``slide_index`` wins; otherwise, when content is being stacked
    onto a slide (``current_slide_index`` is set) that slide is used so multiple
    blocks land on the same slide. Falls back to the last slide when nothing
    else applies.
    """
    idx = item.get('slide_index')
    if idx is None and current_slide_index is not None:
        return current_slide_index
    return resolve_slide_index(engine, idx)


def _apply_mcp_capabilities(
    engine: Any, doc_type: Any, item: dict[str, Any], current_slide_index: int | None = None
) -> None:
    """Apply the Excel/PPT capability fields carried by a ``ContentBlock``.

    These fields are optional and independent of the block's primary ``type``, so
    a single block can both render text and (for example) set a formula on a cell.
    """
    if doc_type == DocumentType.EXCEL:
        sheet = item.get('sheet_name')
        if sheet and hasattr(engine, 'select_sheet'):
            with contextlib.suppress(ValueError):
                engine.select_sheet(sheet)
        cell = item.get('cell')
        if item.get('formula') is not None and cell and hasattr(engine, 'set_formula'):
            engine.set_formula(cell, item['formula'])
        if item.get('chart_type') and item.get('chart_data_range') and hasattr(engine, 'add_chart'):
            engine.add_chart(item['chart_type'], item['chart_data_range'])
        if item.get('number_format') and hasattr(engine, 'set_number_format'):
            rng, fmt = parse_number_format(item['number_format'])
            engine.set_number_format(rng, fmt)
        if item.get('conditional_format') and hasattr(engine, 'add_conditional_format'):
            rng, cf_type, opts = parse_conditional_format(item['conditional_format'])
            engine.add_conditional_format(rng, cf_type, **opts)
        if item.get('data_validation') and hasattr(engine, 'add_data_validation'):
            rng, dv_type, f1, f2 = parse_data_validation(item['data_validation'])
            engine.add_data_validation(rng, dv_type, f1, f2)
        if item.get('freeze') and hasattr(engine, 'freeze_panes'):
            engine.freeze_panes(item['freeze'])
        if item.get('hyperlink') and cell and hasattr(engine, 'add_hyperlink'):
            engine.add_hyperlink(cell, item['hyperlink'])
        if item.get('named_range') and hasattr(engine, 'set_named_range'):
            name, _, rng = item['named_range'].partition('=')
            engine.set_named_range(name.strip(), rng.strip())
    elif doc_type == DocumentType.PPT:
        if item.get('slide_layout') is not None and hasattr(engine, 'apply_layout'):
            idx = _resolve_ppt_slide(engine, item, current_slide_index)
            engine.apply_layout(idx, item['slide_layout'])
        if item.get('notes') is not None and hasattr(engine, 'add_notes'):
            idx = _resolve_ppt_slide(engine, item, current_slide_index)
            engine.add_notes(idx, item['notes'])
        if item.get('transition') is not None and hasattr(engine, 'set_transition'):
            idx = _resolve_ppt_slide(engine, item, current_slide_index)
            engine.set_transition(item['transition'], slide_index=idx)
        if item.get('chart_type') and item.get('chart_data') and hasattr(engine, 'add_chart'):
            idx = _resolve_ppt_slide(engine, item, current_slide_index)
            data = parse_ppt_chart(item['chart_data'])
            engine.add_chart(idx, item['chart_type'], data)
