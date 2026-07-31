"""validate_template — Check template placeholders against data source."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from mcp.errors import McpErrorCode, error_response, success_response


def validate_template(
    template_path: str,
    data: dict[str, Any],
) -> dict:
    """Validate that all template placeholders can be filled."""
    if not Path(template_path).exists():
        return error_response(
            McpErrorCode.DOCUMENT_NOT_FOUND,
            f"'{template_path}' not found.",
        )

    try:
        from src.core.document import open_document

        engine = open_document(template_path)
        text_parts: list[str] = []
        if hasattr(engine, 'doc'):
            text_parts = [p.text for p in engine.doc.paragraphs]
        elif hasattr(engine, 'wb'):
            for ws in engine.wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    for cell in row:
                        if cell:
                            text_parts.append(str(cell))
        else:
            return error_response(
                McpErrorCode.UNSUPPORTED_FORMAT,
                'Template validation supports Word and Excel.',
            )

        full_text = '\n'.join(text_parts)
        placeholders: list[str] = re.findall(
            r'\{\{(?!#\/)([^}]+)\}\}', full_text
        )
        unique_placeholders = list(dict.fromkeys(placeholders))

        flat_data = _flatten(data)
        missing: list[str] = []
        filled: list[str] = []
        loop_keys: list[str] = []
        cond_keys: list[str] = []

        for ph in unique_placeholders:
            if ph.startswith('#each '):
                key = ph.replace('#each ', '').strip()
                loop_keys.append(key)
                if key not in flat_data or not isinstance(flat_data.get(key), list):
                    missing.append(f'{{{{#each {key}}}}} — key missing or not a list')
            elif ph.startswith('#if ') or ph.startswith('#unless '):
                key = ph.replace('#if ', '').replace('#unless ', '').strip()
                cond_keys.append(key)
                if '=' in key:
                    key = key.split('=')[0].strip()
                if key not in flat_data:
                    missing.append(f'{{{{{ph}}}}} — key not in data')
            else:
                if ph in flat_data:
                    filled.append(ph)
                else:
                    missing.append(f'{{{{{ph}}}}} — key not in data')

        warnings: list[str] = []
        for key in loop_keys:
            val = flat_data.get(key)
            if isinstance(val, list) and val:
                first_item = val[0]
                if isinstance(first_item, dict):
                    sub_keys = set(first_item.keys())
                    for ph in unique_placeholders:
                        clean = ph.strip('#').strip()
                        dots = clean.count('.')
                        if dots > 0 and clean not in sub_keys:
                            warnings.append(
                                f'{{{{{ph}}}}} may not resolve in loop {key}'
                            )

        return success_response({
            'template_path': template_path,
            'placeholders_total': len(unique_placeholders),
            'missing': missing,
            'filled': filled,
            'loops_detected': loop_keys,
            'conditions_detected': cond_keys,
            'warnings': warnings,
            'valid': len(missing) == 0,
        })

    except Exception as e:
        return error_response(McpErrorCode.TEMPLATE_ERROR, str(e))


def _flatten(data: dict[str, Any], prefix: str = '') -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in data.items():
        full_key = f'{prefix}.{key}' if prefix else key
        if isinstance(value, dict):
            result.update(_flatten(value, full_key))
        elif isinstance(value, list):
            result[full_key] = value
        else:
            result[full_key] = value
    return result
