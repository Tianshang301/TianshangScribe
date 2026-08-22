"""edit_excel_workbook — Edit an existing Excel workbook with typed operations.

This is a document-type-specific wrapper around ``edit_office_document``: the
typed ``ExcelEditOp`` list is normalised into ``EditOperation`` dicts and the
shared, well-tested dispatch in ``edit.py`` is reused. Agents get a precise,
Excel-only parameter surface instead of the 20+ field generic model.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from tianshang_scribe.mcp.errors import McpErrorCode, error_response
from tianshang_scribe.mcp.schemas import EditOperation, ToolOptions, as_dict
from tianshang_scribe.mcp.tools._dedicated_schemas import ExcelEditOp
from tianshang_scribe.mcp.tools.edit import edit_office_document


def _to_edit_op(op: dict[str, Any]) -> dict[str, Any]:
    action = op.get('action')
    if action == 'write_cell':
        return {
            'action': 'write_cell',
            'cell': op.get('cell'),
            'text': op.get('value'),
            'sheet_name': op.get('sheet_name'),
            'style': op.get('style'),
        }
    if action == 'set_formula':
        return {
            'action': 'set_formula',
            'cell': op.get('cell'),
            'formula': op.get('formula'),
            'sheet_name': op.get('sheet_name'),
        }
    if action == 'freeze_panes':
        return {'action': 'freeze_panes', 'range': op.get('range'), 'sheet_name': op.get('sheet_name')}
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
    return {'action': action}


def edit_excel_workbook(
    input_path: Annotated[str, Field(description='Path to the existing .xlsx workbook.')],
    operations: Annotated[
        list[ExcelEditOp], Field(description='Typed Excel edit operations applied in order.')
    ],
    output_path: Annotated[
        str, Field(description='Output path (defaults to the input file).')
    ] = '',
    options: Annotated[ToolOptions | None, Field(description='Tool options (dry_run, backup).')] = None,
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

    edit_ops = [EditOperation(**_to_edit_op(o)) for o in ops_list]
    return edit_office_document(input_path, edit_ops, output_path=output_path, options=options)
