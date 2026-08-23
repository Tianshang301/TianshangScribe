"""analyze_excel_data — AI-native, read-only Excel inspection.

Returns a structured summary an Agent can reason about: per-sheet dimensions,
headers, sample rows, and per-column type inference with numeric / categorical
statistics, plus null and duplicate-row counts. Never modifies the input file.
"""

from __future__ import annotations

from datetime import datetime
from typing import Annotated, Any, Literal

from pydantic import Field

from tianshang_scribe.mcp.errors import McpErrorCode, error_response, success_response
from tianshang_scribe.mcp.schemas import ToolOptions, as_dict


def _infer_and_summarize(
    name: str, idx: int, values: list[Any], include_stats: bool
) -> dict[str, Any]:
    """Infer a column's type and compute statistics over its values."""
    non_null = [v for v in values if v is not None and v != '']
    null_count = len(values) - len(non_null)
    null_ratio = (null_count / len(values)) if values else 0.0
    unique = set(non_null)

    numeric_vals: list[float] = []
    all_numeric = bool(non_null)
    for v in non_null:
        try:
            numeric_vals.append(float(v))
        except (TypeError, ValueError):
            all_numeric = False

    is_datetime = bool(non_null) and all(isinstance(v, datetime) for v in non_null)

    if is_datetime:
        inferred = 'datetime'
    elif all_numeric:
        inferred = 'numeric'
    elif non_null and 1 < len(unique) <= min(len(non_null), 20):
        inferred = 'categorical'
    else:
        inferred = 'text'

    col: dict[str, Any] = {
        'name': name,
        'index': idx,
        'inferred_type': inferred,
        'null_count': null_count,
        'null_ratio': round(null_ratio, 4),
        'unique_count': len(unique),
    }
    if not include_stats:
        return col

    if inferred == 'numeric' and numeric_vals:
        sorted_n = sorted(numeric_vals)
        n = len(sorted_n)
        total = sum(sorted_n)
        mean = total / n
        mid = n // 2
        median = sorted_n[mid] if n % 2 else (sorted_n[mid - 1] + sorted_n[mid]) / 2
        col['numeric'] = {
            'min': min(sorted_n),
            'max': max(sorted_n),
            'mean': round(mean, 6),
            'median': median,
            'sum': round(total, 6),
        }
    elif inferred == 'categorical':
        counts: dict[Any, int] = {}
        for v in non_null:
            counts[v] = counts.get(v, 0) + 1
        top = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:10]
        col['categorical'] = {'top_values': [{'value': str(k), 'count': v} for k, v in top]}
    return col


def _suggest_pivot(name: str, headers: list[str], data: list[list[Any]]) -> dict[str, Any]:
    """Derive a deterministic pivot-table suggestion from column-type inference.

    ``rows`` takes the first low-cardinality (categorical) column,
    ``columns`` the second, and every numeric column becomes a ``sum``
    value field. When nothing qualifies the rationale says so instead of
    inventing a layout.
    """
    columns = [
        _infer_and_summarize(headers[ci], ci, [r[ci] if ci < len(r) else None for r in data], False)
        for ci in range(len(headers))
    ]
    categoricals = [c['name'] for c in columns if c['inferred_type'] == 'categorical']
    numerics = [c['name'] for c in columns if c['inferred_type'] == 'numeric']
    notes: list[str] = []
    if not categoricals:
        notes.append('no low-cardinality dimension column found for rows/columns')
    if not numerics:
        notes.append('no numeric column found for values')
    return {
        'name': name,
        'suggested_rows': categoricals[:1],
        'suggested_columns': categoricals[1:2],
        'suggested_values': [{'field': n, 'agg': 'sum'} for n in numerics],
        'candidate_dimensions': categoricals,
        'rationale': (
            '; '.join(notes)
            if notes
            else 'rows=first categorical, columns=second categorical, '
            'values=sum of each numeric field'
        ),
    }


def analyze_excel_data(
    input_path: Annotated[str, Field(description='Path to the .xlsx workbook.')],
    mode: Annotated[
        Literal['profile', 'pivot_suggestion'],
        Field(
            description=(
                "'profile' (default) returns the full per-sheet statistical "
                "profile; 'pivot_suggestion' proposes a pivot-table layout "
                '(rows/columns/values/agg) derived from column-type inference.'
            )
        ),
    ] = 'profile',
    options: ToolOptions | None = None,
) -> dict[str, Any]:
    """Analyze an Excel workbook and return structured, Agent-friendly insights."""
    opts: dict[str, Any] = as_dict(options) or {}
    sample_rows = int(opts.get('sample_rows', 10))
    include_stats = bool(opts.get('include_stats', True))
    pivot_mode = mode == 'pivot_suggestion'

    try:
        from tianshang_scribe.core.document import open_document

        engine = open_document(input_path)
        if not hasattr(engine, 'wb'):
            return error_response(
                McpErrorCode.UNSUPPORTED_FORMAT,
                'analyze_excel_data only supports Excel (.xlsx) files; '
                'Word and PowerPoint are not yet analyzed by this tool.',
            )

        wb = engine.wb
        sheets_report: list[dict[str, Any]] = []
        total_dup = 0
        for ws in wb.worksheets:
            grid = [list(r) for r in ws.iter_rows(values_only=True)]
            headers = [('' if c is None else str(c)) for c in (grid[0] if grid else [])]
            data = grid[1:] if grid else []

            if pivot_mode:
                sheets_report.append(_suggest_pivot(ws.title, headers, data))
                continue

            seen: set[tuple[Any, ...]] = set()
            dup = 0
            for r in data:
                key = tuple(r)
                if key in seen:
                    dup += 1
                else:
                    seen.add(key)
            total_dup += dup

            sample = [list(r) for r in data[:sample_rows]]
            columns = [
                _infer_and_summarize(
                    headers[ci], ci, [r[ci] if ci < len(r) else None for r in data], include_stats
                )
                for ci in range(len(headers))
            ]
            sheets_report.append(
                {
                    'name': ws.title,
                    'max_row': ws.max_row,
                    'max_col': ws.max_column,
                    'headers': headers,
                    'row_count': len(data),
                    'duplicate_rows': dup,
                    'sample_rows': sample,
                    'columns': columns,
                }
            )

        payload: dict[str, Any] = {
            'input_path': input_path,
            'mode': mode,
            'sheet_count': len(sheets_report),
            'sheets': sheets_report,
        }
        if not pivot_mode:
            payload['duplicate_row_count'] = total_dup
        return success_response(payload)
    except Exception as e:  # surface any read failure as a structured error
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))
