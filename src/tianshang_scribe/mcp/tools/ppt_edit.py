"""edit_presentation — Edit an existing PowerPoint file with typed operations.

Document-type-specific wrapper around ``edit_office_document``. The typed
``PptEditOp`` list is normalised and dispatched through the shared dispatcher
in ``edit.py``, giving Agents a precise, PPT-only surface. Edge actions that
are deliberately frozen out of the legacy ``EditOperation`` model
(``apply_theme`` / ``set_master_options``, PLAN.md D-7) dispatch straight to
the engine here.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from tianshang_scribe.mcp.errors import McpErrorCode, error_response
from tianshang_scribe.mcp.schemas import ToolOptions, as_dict
from tianshang_scribe.mcp.tools._dedicated_schemas import PptEditOp
from tianshang_scribe.mcp.tools.edit import (
    _apply_edit_operation,
    run_edit_session,
)


def _to_edit_op(op: dict[str, Any]) -> dict[str, Any]:
    """Map a typed ``PptEditOp`` dict onto the shared dispatcher's fields."""
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
    if action == 'add_media':
        media = op.get('media') or {}
        return {
            'action': 'add_media',
            'path': media.get('path'),
            'media_type': media.get('kind') or 'movie',
            'slide_index': idx,
            'left': media.get('left'),
            'top': media.get('top'),
            'width': media.get('width'),
            'height': media.get('height'),
            'poster': media.get('poster'),
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
    if action == 'apply_theme':
        return {'action': 'apply_theme', 'theme': op.get('theme')}
    if action == 'set_master_options':
        return {
            'action': 'set_master_options',
            'slide_number': op.get('slide_number'),
            'footer_text': op.get('footer_text'),
            'date_visible': op.get('date_visible'),
            'date_text': op.get('date_text'),
        }
    return {'action': action}


def _dispatch_ppt_op(engine: Any, op: dict[str, Any]) -> int:
    """Edge actions go straight to the engine; the rest via the shared dispatcher."""
    action = op.get('action')
    if action == 'apply_theme':
        theme = op.get('theme')
        if not theme:
            raise ValueError('apply_theme requires a built-in theme name: "office" or "dark".')
        engine.apply_theme(theme)
        return 1
    if action == 'set_master_options':
        if not any([op.get('slide_number'), op.get('footer_text'), op.get('date_visible')]):
            raise ValueError(
                'set_master_options requires at least one of '
                'slide_number / footer_text / date_visible.'
            )
        engine.set_master_options(
            slide_number=bool(op.get('slide_number')),
            footer_text=op.get('footer_text'),
            date_visible=bool(op.get('date_visible')),
            date_text=op.get('date_text'),
        )
        return 1
    return _apply_edit_operation(engine, _to_edit_op(op))


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
    opts: dict[str, Any] = as_dict(options) or {}
    return run_edit_session(input_path, output_path or input_path, ops_list, opts, _dispatch_ppt_op)
