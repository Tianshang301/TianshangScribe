"""compare_documents — Compare two Word documents for differences."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from src.mcp.errors import McpErrorCode, error_response, success_response
from src.mcp.schemas import ToolOptions


def compare_documents(
    path_a: Annotated[str, Field(description='Path to the first document.')],
    path_b: Annotated[str, Field(description='Path to the second document.')],
    options: Annotated[ToolOptions | None, Field(description='Tool options.')] = None,
) -> dict[str, Any]:
    """Compare two Word documents and report differences."""
    for p, label in [(path_a, 'path_a'), (path_b, 'path_b')]:
        if not Path(p).exists():
            return error_response(
                McpErrorCode.DOCUMENT_NOT_FOUND,
                f"'{label}': '{p}' not found.",
            )

    try:
        from src.core.document import open_document

        engine_a = open_document(path_a)
        engine_b = open_document(path_b)

        if not hasattr(engine_a, 'doc') or not hasattr(engine_b, 'doc'):
            return error_response(
                McpErrorCode.UNSUPPORTED_FORMAT,
                'Document comparison is only supported for Word (.docx) files.',
            )

        paras_a = [p.text for p in engine_a.doc.paragraphs]
        paras_b = [p.text for p in engine_b.doc.paragraphs]

        added: list[dict[str, Any]] = []
        removed: list[dict[str, Any]] = []
        changed: list[dict[str, Any]] = []

        max_len = max(len(paras_a), len(paras_b))
        for i in range(max_len):
            text_a = paras_a[i] if i < len(paras_a) else ''
            text_b = paras_b[i] if i < len(paras_b) else ''
            if not text_a and text_b:
                added.append({'index': i, 'text': text_b[:200]})
            elif text_a and not text_b:
                removed.append({'index': i, 'text': text_a[:200]})
            elif text_a != text_b:
                changed.append(
                    {
                        'index': i,
                        'old': text_a[:200],
                        'new': text_b[:200],
                    }
                )

        return success_response(
            {
                'path_a': path_a,
                'path_b': path_b,
                'paragraphs_a': len(paras_a),
                'paragraphs_b': len(paras_b),
                'added': added,
                'removed': removed,
                'changed': changed,
                'identical': (len(added) == 0 and len(removed) == 0 and len(changed) == 0),
            }
        )
    except Exception as e:
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))
