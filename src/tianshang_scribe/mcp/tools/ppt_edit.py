"""edit_presentation — Edit an existing PowerPoint file with typed operations.

Document-type-specific wrapper around ``edit_office_document``. The typed
``PptEditOp`` list is normalised into ``EditOperation`` dicts and the shared
dispatch in ``edit.py`` is reused, giving Agents a precise, PPT-only surface.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from tianshang_scribe.mcp.errors import McpErrorCode, error_response
from tianshang_scribe.mcp.schemas import EditOperation, ToolOptions, as_dict
from tianshang_scribe.mcp.tools._dedicated_schemas import PptEditOp
from tianshang_scribe.mcp.tools.edit import edit_office_document


def _to_edit_op(op: dict[str, Any]) -> dict[str, Any]:
    action = op.get('action')
    idx = op.get('slide_index')
    if action == 'add_slide':
        return {'action': 'add_slide', 'layout': op.get('layout')}
    if action == 'add_text':
        return {'action': 'add', 'text': op.get('text'), 'slide_index': idx}
    if action == 'replace_text':
        return {'action': 'replace', 'old_text': op.get('old_text'), 'new_text': op.get('new_text')}
    if action == 'add_table':
        table = op.get('table') or {}
        return {
            'action': 'add_table',
            'rows': [table.get('headers', [])] + (table.get('rows') or []),
            'slide_index': idx,
        }
    if action == 'add_chart':
        chart = op.get('chart') or {}
        return {
            'action': 'add_chart',
            'chart_type': chart.get('chart_type'),
            'chart_data': chart.get('data'),
            'slide_index': idx,
        }
    if action == 'add_picture':
        pic = op.get('picture') or {}
        return {
            'action': 'add_picture',
            'path': pic.get('path'),
            'slide_index': idx,
            'left': pic.get('left'),
            'top': pic.get('top'),
            'width': pic.get('width'),
            'height': pic.get('height'),
        }
    if action == 'add_shape':
        return {
            'action': 'add_shape',
            'slide_index': idx,
            'shape_type': op.get('shape_type'),
            'fill': op.get('fill'),
            'line': op.get('line'),
        }
    if action == 'apply_layout':
        return {'action': 'apply_layout', 'slide_index': idx, 'layout': op.get('layout')}
    if action == 'set_transition':
        return {'action': 'set_transition', 'slide_index': idx, 'transition': op.get('transition')}
    if action == 'add_notes':
        return {'action': 'add_notes', 'slide_index': idx, 'notes': op.get('notes')}
    return {'action': action}


def edit_presentation(
    input_path: Annotated[str, Field(description='Path to the existing .pptx presentation.')],
    operations: Annotated[
        list[PptEditOp], Field(description='Typed PowerPoint edit operations applied in order.')
    ],
    output_path: Annotated[
        str, Field(description='Output path (defaults to the input file).')
    ] = '',
    options: Annotated[
        ToolOptions | None, Field(description='Tool options (dry_run, backup).')
    ] = None,
) -> dict[str, Any]:
    """Edit an existing PowerPoint presentation with typed, PPT-only edit operations.

    Side effects: rewrites the presentation at ``input_path`` (or ``output_path``).
    """
    ops_list: list[dict[str, Any]] = as_dict(operations)
    if not Path(input_path).exists():
        return error_response(McpErrorCode.DOCUMENT_NOT_FOUND, f"'{input_path}' not found.")
    if not input_path.lower().endswith('.pptx'):
        return error_response(
            McpErrorCode.INVALID_PARAMETER,
            'edit_presentation only accepts .pptx presentations.',
        )

    edit_ops = [EditOperation(**_to_edit_op(o)) for o in ops_list]
    return edit_office_document(input_path, edit_ops, output_path=output_path, options=options)
