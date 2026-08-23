"""edit_office_document — Edit existing Office documents."""

from __future__ import annotations

import contextlib
from collections.abc import Callable
from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from tianshang_scribe.core.document import open_document
from tianshang_scribe.mcp.errors import (
    DOCUMENTATION_URL,
    McpErrorCode,
    error_response,
    success_response,
)
from tianshang_scribe.mcp.schemas import EditOperation, ToolOptions, as_dict
from tianshang_scribe.mcp.tools._dryrun import build_edit_plan, classify_engine_error
from tianshang_scribe.mcp.tools._parse import (
    parse_conditional_format,
    parse_data_validation,
    parse_ppt_chart,
    resolve_slide_index,
)
from tianshang_scribe.utils.file_utils import ensure_parent_dir


class UnknownEditActionError(Exception):
    """Raised when an operation carries an action the dispatcher does not know."""


def edit_office_document(
    input_path: Annotated[str, Field(description='Path to the existing document.')],
    operations: Annotated[
        list[EditOperation], Field(description='List of edit operations applied in order.')
    ],
    output_path: Annotated[
        str, Field(description='Output path (defaults to the input file).')
    ] = '',
    options: Annotated[ToolOptions | None, Field(description='Tool options.')] = None,
) -> dict[str, Any]:
    """Edit an existing Office document with replace/delete/modify/style operations."""
    ops_list: list[dict[str, Any]] = as_dict(operations)
    opts: dict[str, Any] = as_dict(options) or {}
    if not Path(input_path).exists():
        return error_response(McpErrorCode.DOCUMENT_NOT_FOUND, f"'{input_path}' not found.")

    output_path = output_path or input_path

    return run_edit_session(input_path, output_path, ops_list, opts, _apply_edit_operation)


def run_edit_session(
    input_path: str,
    output_path: str,
    ops_list: list[dict[str, Any]],
    opts: dict[str, Any],
    dispatch: Callable[[Any, dict[str, Any]], int],
) -> dict[str, Any]:
    """Shared edit shell: dry-run plan, ordered dispatch, save, error mapping.

    Used by ``edit_office_document`` and re-used by the dedicated Excel/PPT
    wrappers whose edge actions dispatch outside the legacy ``EditOperation``
    model (capability freeze, PLAN.md D-7).
    """
    if opts.get('dry_run'):
        plan = build_edit_plan(input_path, ops_list)
        return success_response(
            {
                'dry_run': True,
                'file': input_path,
                'operations': len(ops_list),
                'op_types': [o.get('action') for o in ops_list],
                **plan,
            }
        )

    op_index = 0
    try:
        engine = open_document(input_path)
        changes = 0
        for i, op in enumerate(ops_list):
            op_index = i
            changes += dispatch(engine, op)

        if opts.get('backup') and input_path == output_path:
            import shutil

            shutil.copy2(input_path, input_path + '.bak')

        ensure_parent_dir(output_path)
        engine.save(output_path)
        return success_response(
            {
                'output_path': output_path,
                'operations': len(ops_list),
                'total_changes': changes,
            }
        )
    except UnknownEditActionError as e:
        return error_response(
            McpErrorCode.INVALID_PARAMETER,
            str(e),
            field=f'operations[{op_index}].action',
            documentation_url=DOCUMENTATION_URL,
        )
    except (ValueError, IndexError) as e:
        refined = classify_engine_error(e)
        if refined is not None:
            code, fld = refined
            return error_response(
                code,
                str(e),
                field=f'operations[{op_index}].{fld}',
                documentation_url=DOCUMENTATION_URL,
            )
        if isinstance(e, ValueError):
            return error_response(McpErrorCode.DOCUMENT_LOCKED, str(e))
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))
    except Exception as e:
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))


def _apply_edit_operation(engine: Any, op: dict[str, Any]) -> int:
    """Execute one normalised edit operation against ``engine``; return changes."""
    action = op.get('action', '')
    changes = 0
    if action == 'replace':
        old = op.get('old_text', '')
        new = op.get('new_text', '')
        regex = op.get('regex', False)
        count = engine.replace_text(old, new, regex=regex)
        changes += count
    elif action == 'delete':
        target = op.get('target', '')
        count = engine.replace_text(target, '', regex=op.get('regex', False))
        changes += count
    elif action == 'modify':
        old = op.get('old_text', '')
        new = op.get('new_text', '')
        count = engine.replace_text(old, new, regex=False)
        changes += count
    elif action == 'style':
        style_str = op.get('style', '')
        if style_str:
            engine.set_style(style_str)
        if op.get('apply_all', True):
            engine.apply_style_to_all()
        changes += 1
    elif action == 'add':
        text = op.get('text', '')
        column = op.get('column', 1)
        slide_index = op.get('slide_index')
        engine.add_text(text, column=column, slide_index=slide_index)
        changes += 1
    elif action == 'clear':
        if hasattr(engine, 'clear_content'):
            engine.clear_content()
            changes += 1
    # --- Excel / PPT capability actions (P0) ---
    elif action == 'write_cell':
        _edit_write_cell(engine, op)
        changes += 1
    elif action == 'set_formula':
        cell = op.get('cell')
        formula = op.get('formula')
        if cell and formula is not None and hasattr(engine, 'set_formula'):
            if op.get('sheet_name') and hasattr(engine, 'select_sheet'):
                _try_select_sheet(engine, op['sheet_name'])
            engine.set_formula(cell, formula)
            changes += 1
    elif action == 'freeze_panes':
        rng = op.get('range')
        if rng and hasattr(engine, 'freeze_panes'):
            engine.freeze_panes(rng)
            changes += 1
    elif action == 'add_chart':
        chart_type = op.get('chart_type')
        if op.get('chart_data_range') and hasattr(engine, 'add_chart'):
            if op.get('sheet_name') and hasattr(engine, 'select_sheet'):
                _try_select_sheet(engine, op['sheet_name'])
            engine.add_chart(chart_type, op['chart_data_range'])
            changes += 1
        elif op.get('chart_data') and hasattr(engine, 'add_chart'):
            idx = resolve_slide_index(engine, op.get('slide_index'))
            data = parse_ppt_chart(op['chart_data'])
            engine.add_chart(idx, chart_type, data)
            changes += 1
    elif action == 'conditional_format':
        spec = op.get('conditional_format') or op.get('range')
        if spec and hasattr(engine, 'add_conditional_format'):
            cr, cf_type, cf_opts = parse_conditional_format(spec)
            engine.add_conditional_format(cr, cf_type, **cf_opts)
            changes += 1
    elif action == 'data_validation':
        spec = op.get('data_validation') or op.get('range')
        if spec and hasattr(engine, 'add_data_validation'):
            cr, dv_type, f1, f2 = parse_data_validation(spec)
            engine.add_data_validation(cr, dv_type, f1, f2)
            changes += 1
    elif action == 'add_table':
        rows = op.get('rows') or []
        if rows and hasattr(engine, 'add_table'):
            idx = resolve_slide_index(engine, op.get('slide_index'))
            engine.add_table(idx, rows[1:], col_names=rows[0])
            changes += 1
    elif action == 'add_picture':
        path = op.get('path')
        if path and hasattr(engine, 'add_picture'):
            idx = resolve_slide_index(engine, op.get('slide_index'))
            engine.add_picture(idx, path)
            changes += 1
    elif action == 'add_shape':
        if hasattr(engine, 'add_shape'):
            idx = resolve_slide_index(engine, op.get('slide_index'))
            engine.add_shape(
                idx,
                op.get('shape_type', 'rectangle'),
                fill=op.get('fill'),
                line=op.get('line'),
            )
            changes += 1
    elif action == 'sort':
        rng = op.get('range')
        if rng and hasattr(engine, 'sort'):
            engine.sort(
                rng,
                key_columns=op.get('key_columns'),
                orders=op.get('orders'),
                order=op.get('order') or 'asc',
            )
            changes += 1
    elif action == 'add_sheet':
        if op.get('sheet_name') and hasattr(engine, 'add_sheet'):
            engine.add_sheet(op['sheet_name'])
            changes += 1
    elif action == 'set_range_style':
        rng = op.get('range')
        if rng and op.get('style') and hasattr(engine, 'set_range_style'):
            engine.set_range_style(rng, op['style'])
            changes += 1
    elif action == 'number_format':
        spec = op.get('number_format')
        if spec and '=' in spec and hasattr(engine, 'set_number_format'):
            rng, fmt = spec.split('=', 1)
            engine.set_number_format(rng.strip(), fmt.strip())
            changes += 1
    elif action == 'add_slide':
        if hasattr(engine, 'add_slide'):
            layout = op.get('layout')
            if isinstance(layout, str) and hasattr(engine, 'prs'):
                for i, lay in enumerate(engine.prs.slide_layouts):
                    if lay.name == layout:
                        layout = i
                        break
            engine.add_slide(layout if isinstance(layout, int) else 1)
            changes += 1
    elif action == 'apply_layout':
        if op.get('layout') is not None and hasattr(engine, 'apply_layout'):
            idx = resolve_slide_index(engine, op.get('slide_index'))
            engine.apply_layout(idx, op['layout'])
            changes += 1
    elif action == 'set_transition':
        if op.get('transition') is not None and hasattr(engine, 'set_transition'):
            engine.set_transition(op['transition'], slide_index=op.get('slide_index'))
            changes += 1
    elif action == 'add_notes':
        if op.get('notes') is not None and hasattr(engine, 'add_notes'):
            idx = resolve_slide_index(engine, op.get('slide_index'))
            engine.add_notes(idx, op['notes'])
            changes += 1
    # --- 0.9.0 core actions (legacy surface frozen at these four, D-7) ---
    elif action == 'group_rows':
        rng = op.get('range')
        if rng and hasattr(engine, 'group_rows'):
            if op.get('sheet_name'):
                _try_select_sheet(engine, op['sheet_name'])
            engine.group_rows(
                rng,
                outline_level=op.get('outline_level') or 1,
                hidden=bool(op.get('hidden')),
            )
            changes += 1
    elif action == 'group_columns':
        rng = op.get('range')
        if rng and hasattr(engine, 'group_columns'):
            if op.get('sheet_name'):
                _try_select_sheet(engine, op['sheet_name'])
            engine.group_columns(
                rng,
                outline_level=op.get('outline_level') or 1,
                hidden=bool(op.get('hidden')),
            )
            changes += 1
    elif action == 'ungroup':
        rng = op.get('range')
        if rng and hasattr(engine, 'ungroup'):
            if op.get('sheet_name'):
                _try_select_sheet(engine, op['sheet_name'])
            engine.ungroup(rng, axis=op.get('axis') or 'rows')
            changes += 1
    elif action == 'add_media':
        path = op.get('path')
        if path and hasattr(engine, 'add_movie'):
            idx = resolve_slide_index(engine, op.get('slide_index'))
            left = op.get('left')
            top = op.get('top')
            width = op.get('width')
            height = op.get('height')
            if op.get('media_type') == 'audio':
                engine.add_audio(
                    idx,
                    path,
                    left=left if left is not None else 1.0,
                    top=top if top is not None else 1.0,
                )
            else:
                engine.add_movie(
                    idx,
                    path,
                    left=left if left is not None else 1.0,
                    top=top if top is not None else 1.0,
                    width=width if width is not None else 6.0,
                    height=height if height is not None else 4.5,
                    poster=op.get('poster'),
                )
            changes += 1
    else:
        raise UnknownEditActionError(f'Unknown edit action: {action!r}')
    return changes


def _try_select_sheet(engine: Any, sheet: str) -> None:
    with contextlib.suppress(ValueError):
        engine.select_sheet(sheet)


def _edit_write_cell(engine: Any, op: dict[str, Any]) -> None:
    """Write a value into an Excel cell (formula-aware).

    Resolution is per-operation: an explicit ``sheet_name`` targets that sheet
    without mutating the engine's persistent sheet selection, so later ops
    without ``sheet_name`` still default to the engine's active/target sheet.

    ``is_formula`` disambiguates strings that start with ``=``: ``True`` stores
    the text as a formula (it must start with ``=``), ``False`` forces a
    literal string cell even when it starts with ``=``, and omitted keeps the
    automatic openpyxl behaviour.
    """
    cell = op.get('cell')
    val = op.get('text', '')
    if not cell:
        return
    ws = None
    sheet = op.get('sheet_name')
    if sheet and hasattr(engine, 'wb') and sheet in engine.wb.sheetnames:
        ws = engine.wb[sheet]
    elif hasattr(engine, '_ws'):
        ws = engine._ws()
    is_formula = op.get('is_formula')
    if ws is not None:
        if is_formula is True:
            if not isinstance(val, str) or not val.startswith('='):
                raise ValueError(
                    f'write_cell is_formula=true requires a string starting with "=", got {val!r}'
                )
            ws[cell] = val
        else:
            ws[cell] = val
            if is_formula is False and isinstance(val, str) and val.startswith('='):
                ws[cell].data_type = 's'
    style = op.get('style')
    if style and ws is not None and hasattr(engine, 'set_range_style'):
        prev = getattr(engine, '_selected_sheet', None)
        try:
            if sheet:
                with contextlib.suppress(ValueError):
                    engine.select_sheet(sheet)
            engine.set_range_style(f'{cell}:{cell}', style)
        finally:
            if prev is None:
                engine._selected_sheet = None
            else:
                with contextlib.suppress(ValueError):
                    engine.select_sheet(prev)
