"""analyze_excel_data — AI-native, read-only Excel inspection.

Returns a structured summary an Agent can reason about: per-sheet dimensions,
headers, sample rows, and per-column type inference with numeric / categorical
statistics, plus null and duplicate-row counts. Never modifies the input file.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from tianshang_scribe.mcp.errors import McpErrorCode, error_response, success_response
from tianshang_scribe.mcp.schemas import ToolOptions, as_dict


def _infer_and_summarize(name: str, idx: int, values: list[Any], include_stats: bool) -> dict[str, Any]:
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
        col['categorical'] = {
            'top_values': [{'value': str(k), 'count': v} for k, v in top]
        }
    return col


def analyze_excel_data(
    input_path: str,
    options: ToolOptions | None = None,
) -> dict[str, Any]:
    """Analyze an Excel workbook and return structured, Agent-friendly insights."""
    opts: dict[str, Any] = as_dict(options) or {}
    sample_rows = int(opts.get('sample_rows', 10))
    include_stats = bool(opts.get('include_stats', True))

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
                _infer_and_summarize(headers[ci], ci, [r[ci] if ci < len(r) else None for r in data], include_stats)
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

        return success_response(
            {
                'input_path': input_path,
                'sheet_count': len(sheets_report),
                'duplicate_row_count': total_dup,
                'sheets': sheets_report,
            }
        )
    except Exception as e:  # surface any read failure as a structured error
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))
