"""compare_excel_workbooks — Read-only diff between two Excel workbooks.

Compares workbooks without modifying either input. ``structure`` mode reports
sheet-level additions/removals/renames and dimension drift; ``data`` mode
additionally streams a cell-level value diff (with numeric ``tolerance``).
``formula`` / ``full`` modes are wired in the 0.9.0 P2 phase.
"""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Annotated, Any, Literal

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pydantic import Field

from tianshang_scribe.mcp.errors import McpErrorCode, error_response, success_response

#: Per-sheet cap on reported cell diffs before truncation (PLAN.md B.3).
MAX_CELL_DIFFS_PER_SHEET = 10_000

_MODES = ('structure', 'data', 'formula', 'full')
_PENDING_MODES = ('formula', 'full')


def _stream_hash(ws: Any) -> str:
    """Stable hash of a sheet's value stream (cheap sheet-level prefilter)."""
    digest = hashlib.sha256()
    for row in ws.iter_rows(values_only=True):
        digest.update(repr(row).encode('utf-8', 'surrogatepass'))
    return digest.hexdigest()


def _values_equal(a: Any, b: Any, tolerance: float) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= tolerance
    return bool(a == b)


def _diff_stream(
    ws_a: Any, ws_b: Any, tolerance: float, cap: int
) -> tuple[list[dict[str, Any]], bool]:
    """Stream-compare two sheets position-by-position (O(1) memory)."""
    diffs: list[dict[str, Any]] = []
    truncated = False
    rows_a = ws_a.iter_rows(values_only=True)
    rows_b = ws_b.iter_rows(values_only=True)
    max_rows = max(ws_a.max_row or 0, ws_b.max_row or 0)
    for r in range(1, max_rows + 1):
        row_a = next(rows_a, ())
        row_b = next(rows_b, ())
        width = max(len(row_a), len(row_b))
        for c in range(1, width + 1):
            va = row_a[c - 1] if c <= len(row_a) else None
            vb = row_b[c - 1] if c <= len(row_b) else None
            if _values_equal(va, vb, tolerance):
                continue
            diffs.append(
                {'cell': f'{get_column_letter(c)}{r}', 'kind': 'value', 'before': va, 'after': vb}
            )
            if len(diffs) >= cap:
                truncated = True
                return diffs, truncated
    return diffs, truncated


def compare_excel_workbooks(
    path_a: Annotated[str, Field(description='First .xlsx workbook.')],
    path_b: Annotated[str, Field(description='Second .xlsx workbook.')],
    mode: Annotated[
        Literal['structure', 'data', 'formula', 'full'],
        Field(description="Diff depth: 'structure' (sheets/dims), 'data' (+cell values)."),
    ] = 'structure',
    tolerance: Annotated[
        float, Field(description='Numeric tolerance applied to value comparison (data mode).')
    ] = 0.0,
    sheet_mapping: Annotated[
        dict[str, str] | None,
        Field(
            description="Rename map applied to path_b sheets before diffing, e.g. {'Data2024': 'Data'}."
        ),
    ] = None,
) -> dict[str, Any]:
    """Compare two Excel workbooks and report sheet/cell differences.

    Read-only — never modifies either input file. ``data`` mode compares
    cached cell values; ``formula`` and ``full`` modes arrive in a later
    0.9.0 phase. For PowerPoint inspection use extract_presentation_data.
    """
    for label, p in (('path_a', path_a), ('path_b', path_b)):
        if not Path(p).exists():
            return error_response(McpErrorCode.DOCUMENT_NOT_FOUND, f"'{p}' not found ({label}).")
        if not str(p).lower().endswith(('.xlsx', '.xlsm')):
            return error_response(
                McpErrorCode.UNSUPPORTED_FORMAT,
                f"'{p}' is not an .xlsx/.xlsm workbook.",
            )
    if mode in _PENDING_MODES:
        return error_response(
            McpErrorCode.INVALID_PARAMETER,
            f"mode {mode!r} is not available yet; use 'structure' or 'data'.",
        )

    mapping = dict(sheet_mapping or {})
    try:
        data_mode = mode == 'data'
        wb_a = load_workbook(path_a, read_only=True, data_only=data_mode)
        wb_b = load_workbook(path_b, read_only=True, data_only=data_mode)

        mapped_b: dict[str, str] = {mapping.get(n, n): n for n in wb_b.sheetnames}
        names_a = set(wb_a.sheetnames)
        added = sorted(set(mapped_b) - names_a)
        removed = sorted(names_a - set(mapped_b))
        renamed = [
            {'from': orig, 'to': eff} for eff, orig in sorted(mapped_b.items()) if orig != eff
        ]
        common = sorted(names_a & set(mapped_b))

        sheets_block: dict[str, Any] = {
            'added': added,
            'removed': removed,
            'renamed': renamed,
        }
        cells: list[dict[str, Any]] = []
        truncated = False
        identical: list[str] = []
        dims_changed: list[dict[str, Any]] = []

        for name in common:
            ws_a = wb_a[name]
            ws_b = wb_b[mapped_b[name]]
            if mode == 'data' and _stream_hash(ws_a) == _stream_hash(ws_b):
                identical.append(name)
                continue
            if mode == 'structure':
                ra, ca = ws_a.max_row or 0, ws_a.max_column or 0
                rb, cb = ws_b.max_row or 0, ws_b.max_column or 0
                if (ra, ca) != (rb, cb):
                    dims_changed.append(
                        {
                            'sheet': name,
                            'a': {'rows': ra, 'cols': ca},
                            'b': {'rows': rb, 'cols': cb},
                        }
                    )
                continue
            sheet_diffs, hit_cap = _diff_stream(ws_a, ws_b, tolerance, MAX_CELL_DIFFS_PER_SHEET)
            truncated = truncated or hit_cap
            for d in sheet_diffs:
                cells.append({'sheet': name, **d})

        wb_a.close()
        wb_b.close()
        payload: dict[str, Any] = {
            'path_a': path_a,
            'path_b': path_b,
            'mode': mode,
            'tolerance': tolerance,
            'sheets': sheets_block,
        }
        if mode == 'data':
            payload['cells'] = cells
            payload['cell_diff_count'] = len(cells)
            payload['truncated'] = truncated
            payload['identical_sheets'] = identical
        else:
            payload['dims_changed'] = dims_changed
        return success_response(payload)
    except Exception as e:  # surface any read failure as a structured error
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))
