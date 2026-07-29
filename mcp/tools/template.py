"""fill_template — Fill document templates with structured data."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from mcp.errors import McpErrorCode, error_response, success_response
from src.rendering.template import TemplateEngine


def fill_template(
    template_path: str,
    data: dict[str, Any],
    output_path: str = '',
    options: dict[str, Any] | None = None,
) -> dict:
    """Fill a document template with data from a JSON object."""
    if not Path(template_path).exists():
        return error_response(McpErrorCode.DOCUMENT_NOT_FOUND, f"'{template_path}' not found.")

    output_path = output_path or template_path.replace('.docx', '_filled.docx').replace(
        '.xlsx', '_filled.xlsx'
    ).replace('.pptx', '_filled.pptx')
    opts = options or {}

    if opts.get('dry_run'):
        return success_response({
            'dry_run': True,
            'template': template_path,
            'keys_to_fill': list(data.keys()),
            'output': output_path,
        })

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
        doc.save(output_path)

        tmp.unlink(missing_ok=True)

        return success_response({
            'output_path': output_path,
            'placeholders_filled': count,
            'keys_processed': len(data),
        })
    except Exception as e:
        return error_response(McpErrorCode.TEMPLATE_ERROR, str(e))
