"""fill_template — Fill document templates with structured data."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from src.mcp.errors import McpErrorCode, error_response, success_response
from src.mcp.schemas import ToolOptions, as_dict
from src.rendering.template import TemplateEngine
from src.utils.file_utils import ensure_parent_dir


def fill_template(
    template_path: Annotated[
        str, Field(description='Path to the template document containing placeholders.')
    ],
    data: Annotated[
        dict[str, Any],
        Field(description='Key-value data to fill placeholders. Supports nested objects.'),
    ],
    output_path: Annotated[str, Field(description='Output file path.')] = '',
    options: Annotated[ToolOptions | None, Field(description='Tool options.')] = None,
) -> dict[str, Any]:
    """Fill a document template with data from a JSON object."""
    opts: dict[str, Any] = as_dict(options) or {}
    if not Path(template_path).exists():
        return error_response(McpErrorCode.DOCUMENT_NOT_FOUND, f"'{template_path}' not found.")

    output_path = output_path or template_path.replace('.docx', '_filled.docx').replace(
        '.xlsx', '_filled.xlsx'
    ).replace('.pptx', '_filled.pptx')

    if opts.get('dry_run'):
        return success_response(
            {
                'dry_run': True,
                'template': template_path,
                'keys_to_fill': list(data.keys()),
                'output': output_path,
            }
        )

    try:
        import tempfile

        tmp = Path(tempfile.gettempdir()) / 'mcp_template_data.json'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)

        engine = TemplateEngine(str(tmp))
        from src.core.document import open_document

        doc = open_document(template_path)

        if opts.get('backup'):
            import shutil

            shutil.copy2(template_path, template_path + '.bak')

        count = engine.fill(doc)
        ensure_parent_dir(output_path)
        doc.save(output_path)

        tmp.unlink(missing_ok=True)

        return success_response(
            {
                'output_path': output_path,
                'placeholders_filled': count,
                'keys_processed': len(data),
            }
        )
    except Exception as e:
        return error_response(McpErrorCode.TEMPLATE_ERROR, str(e))
