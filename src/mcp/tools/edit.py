"""edit_office_document — Edit existing Office documents."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from src.core.document import open_document
from src.mcp.errors import McpErrorCode, error_response, success_response
from src.mcp.schemas import EditOperation, ToolOptions, as_dict
from src.utils.file_utils import ensure_parent_dir


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

    if opts.get('dry_run'):
        return success_response(
            {
                'dry_run': True,
                'file': input_path,
                'operations': len(ops_list),
                'op_types': [o.get('action') for o in ops_list],
            }
        )

    try:
        engine = open_document(input_path)
        changes = 0

        for op in ops_list:
            action = op.get('action', '')
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
                engine.add_text(text, column=column)
                changes += 1
            elif action == 'clear':
                if hasattr(engine, 'clear_content'):
                    engine.clear_content()
                    changes += 1

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
    except ValueError as e:
        return error_response(McpErrorCode.DOCUMENT_LOCKED, str(e))
    except Exception as e:
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))
