"""edit_excel_workbook — Edit an existing Excel workbook with typed operations.

This is a document-type-specific wrapper around ``edit_office_document``: the
typed ``ExcelEditOp`` list is normalised and dispatched through the shared,
well-tested dispatcher in ``edit.py``, giving Agents a precise, Excel-only
parameter surface instead of the 20+ field generic model. Edge actions that
are deliberately frozen out of the legacy ``EditOperation`` model
(``set_tab_color`` / ``set_print_area`` / ``set_page_setup``, PLAN.md D-7)
dispatch straight to the engine here.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from tianshang_scribe.mcp.errors import McpErrorCode, error_response
from tianshang_scribe.mcp.schemas import ToolOptions, as_dict
from tianshang_scribe.mcp.tools._dedicated_schemas import ExcelEditOp
from tianshang_scribe.mcp.tools.edit import (
    _apply_edit_operation,
    _try_select_sheet,
    run_edit_session,
)


def _to_edit_op(op: dict[str, Any]) -> dict[str, Any]:
    """Map a typed ``ExcelEditOp`` dict onto the shared dispatcher's fields."""
    action = op.get('action')
    if action == 'write_cell':
        return {
            'action': 'write_cell',
            'cell': op.get('cell'),
            'text': op.get('value'),
            'sheet_name': op.get('sheet_name'),
            'style': op.get('style'),
            'is_formula': op.get('is_formula'),
        }
    if action == 'set_formula':
        return {
            'action': 'set_formula',
            'cell': op.get('cell'),
            'formula': op.get('formula'),
            'sheet_name': op.get('sheet_name'),
        }
    if action == 'freeze_panes':
        return {
            'action': 'freeze_panes',
            'range': op.get('range'),
            'sheet_name': op.get('sheet_name'),
        }
    if action == 'add_chart':
        return {
            'action': 'add_chart',
            'chart_type': op.get('chart_type'),
            'chart_data_range': op.get('chart_data_range'),
            'sheet_name': op.get('sheet_name'),
        }
    if action == 'conditional_format':
        return {
            'action': 'conditional_format',
            'conditional_format': op.get('conditional_format'),
            'sheet_name': op.get('sheet_name'),
        }
    if action == 'data_validation':
        return {
            'action': 'data_validation',
            'data_validation': op.get('data_validation'),
            'sheet_name': op.get('sheet_name'),
        }
    if action == 'add_table':
        rows = op.get('rows') or []
        headers = op.get('headers')
        table = ([headers] + rows) if headers else rows  # noqa: RUF005  # headers is itself a list row
        return {'action': 'add_table', 'rows': table, 'sheet_name': op.get('sheet_name')}
    if action == 'sort':
        return {
            'action': 'sort',
            'range': op.get('range'),
            'key_columns': op.get('key_columns'),
            'orders': op.get('orders'),
            'order': op.get('order'),
        }
    if action == 'add_sheet':
        return {'action': 'add_sheet', 'sheet_name': op.get('sheet_name')}
    if action == 'set_range_style':
        return {'action': 'set_range_style', 'range': op.get('range'), 'style': op.get('style')}
    if action == 'number_format':
        return {'action': 'number_format', 'number_format': op.get('number_format')}
    if action == 'group_rows':
        return {
            'action': 'group_rows',
            'range': op.get('range'),
            'outline_level': op.get('outline_level'),
            'hidden': op.get('hidden'),
            'sheet_name': op.get('sheet_name'),
        }
    if action == 'group_columns':
        return {
            'action': 'group_columns',
            'range': op.get('range'),
            'outline_level': op.get('outline_level'),
            'hidden': op.get('hidden'),
            'sheet_name': op.get('sheet_name'),
        }
    if action == 'ungroup':
        return {
            'action': 'ungroup',
            'range': op.get('range'),
            'axis': op.get('axis'),
            'sheet_name': op.get('sheet_name'),
        }
    return {'action': action}


def _dispatch_excel_op(engine: Any, op: dict[str, Any]) -> int:
    """Edge actions go straight to the engine; the rest via the shared dispatcher."""
    action = op.get('action')
    if action == 'set_tab_color':
        color = op.get('tab_color')
        if not color:
            raise ValueError("set_tab_color requires 'tab_color' (RGB hex, e.g. 'FF0000').")
        if op.get('sheet_name'):
            _try_select_sheet(engine, op['sheet_name'])
        engine.set_tab_color(color)
        return 1
    if action == 'set_print_area':
        rng = op.get('range')
        if not rng:
            raise ValueError("set_print_area requires 'range' (e.g. 'A1:C10').")
        if op.get('sheet_name'):
            _try_select_sheet(engine, op['sheet_name'])
        engine.set_print_area(rng)
        return 1
    if action == 'set_page_setup':
        if op.get('sheet_name'):
            _try_select_sheet(engine, op['sheet_name'])
        engine.set_page_setup(
            paper_size=op.get('paper_size'),
            orientation=op.get('orientation'),
            margins=op.get('margins'),
            header=op.get('header'),
            footer=op.get('footer'),
        )
        return 1
    return _apply_edit_operation(engine, _to_edit_op(op))


def edit_excel_workbook(
    input_path: Annotated[str, Field(description='Path to the existing .xlsx workbook.')],
    operations: Annotated[
        list[ExcelEditOp], Field(description='Typed Excel edit operations applied in order.')
    ],
    output_path: Annotated[
        str, Field(description='Output path (defaults to the input file).')
    ] = '',
    options: Annotated[
        ToolOptions | None, Field(description='Tool options (dry_run, backup).')
    ] = None,
) -> dict[str, Any]:
    """Edit an existing Excel workbook with typed, Excel-only edit operations.

    Side effects: rewrites the workbook at ``input_path`` (or ``output_path``).
    For analysis without mutation use ``analyze_excel_data`` (read-only).
    """
    ops_list: list[dict[str, Any]] = as_dict(operations)
    if not Path(input_path).exists():
        return error_response(McpErrorCode.DOCUMENT_NOT_FOUND, f"'{input_path}' not found.")
    if not input_path.lower().endswith(('.xlsx', '.xlsm')):
        return error_response(
            McpErrorCode.INVALID_PARAMETER,
            'edit_excel_workbook only accepts .xlsx/.xlsm workbooks.',
        )
    opts: dict[str, Any] = as_dict(options) or {}
    return run_edit_session(
        input_path, output_path or input_path, ops_list, opts, _dispatch_excel_op
    )
