"""create_excel_workbook — Build a new .xlsx workbook with typed sheet specs.

Document-type-specific tool that mirrors ``create_office_document`` for Excel
only, but with a precise ``ExcelSheetSpec`` model. It builds the workbook
directly (no ``ContentBlock`` mega-model) so Agents get discoverable parameters.
"""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from tianshang_scribe.core.document import DocumentType, create_document
from tianshang_scribe.mcp.errors import McpErrorCode, error_response, success_response
from tianshang_scribe.mcp.schemas import ToolOptions, as_dict
from tianshang_scribe.mcp.tools._dedicated_schemas import ExcelSheetSpec
from tianshang_scribe.mcp.tools._parse import parse_conditional_format, parse_data_validation
from tianshang_scribe.utils.file_utils import ensure_parent_dir


def create_excel_workbook(
    output_path: Annotated[str, Field(description='Output .xlsx path to create.')],
    sheets: Annotated[list[ExcelSheetSpec], Field(description='Worksheets to build (in order).')],
    metadata: Annotated[
        dict[str, Any] | None, Field(description='Optional document properties (title/author/...).')
    ] = None,
    options: Annotated[ToolOptions | None, Field(description='Tool options.')] = None,
) -> dict[str, Any]:
    """Create a brand new Excel workbook from typed sheet specifications.

    Side effects: writes a new ``.xlsx`` file at ``output_path``. Use
    ``edit_excel_workbook`` to modify an existing workbook; use
    ``analyze_excel_data`` for read-only inspection.
    """
    specs: list[dict[str, Any]] = as_dict(sheets)
    if not output_path.lower().endswith(('.xlsx', '.xlsm')):
        return error_response(
            McpErrorCode.INVALID_PARAMETER,
            'create_excel_workbook only writes .xlsx/.xlsm workbooks.',
        )
    if not specs:
        return error_response(
            McpErrorCode.INVALID_PARAMETER,
            'create_excel_workbook requires at least one sheet spec.',
        )

    try:
        engine: Any = create_document(DocumentType.EXCEL)
        built: list[str] = []
        for raw in specs:
            name = raw.get('name') or f'Sheet{len(built) + 1}'
            ws = engine.wb.create_sheet(title=name)
            engine.select_sheet(name)
            if raw.get('headers'):
                ws.append(raw['headers'])
            for row in raw.get('rows') or []:
                ws.append(row)
            for cell, formula in (raw.get('formulas') or {}).items():
                ws[cell] = formula
            if raw.get('freeze'):
                ws.freeze_panes = raw['freeze']
            if raw.get('number_format'):
                rng, fmt = raw['number_format'].split('=', 1)
                engine.set_number_format(rng.strip(), fmt.strip())
            if raw.get('conditional_format'):
                cr, cf_type, cf_opts = parse_conditional_format(raw['conditional_format'])
                engine.add_conditional_format(cr, cf_type, **cf_opts)
            if raw.get('data_validation'):
                cr, dv_type, f1, f2 = parse_data_validation(raw['data_validation'])
                engine.add_data_validation(cr, dv_type, f1, f2)
            for col, w in (raw.get('column_widths') or {}).items():
                ws.column_dimensions[col].width = w
            built.append(name)

        if metadata and hasattr(engine, 'set_metadata'):
            engine.set_metadata(**metadata)
        ensure_parent_dir(output_path)
        engine.save(output_path)
        return success_response(
            {'output_path': output_path, 'sheets': built, 'sheet_count': len(built)}
        )
    except Exception as e:  # surface any build failure as a structured error
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))
