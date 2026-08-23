"""TianshangScribe MCP Server — Structured error types."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

_MIME_MAP: dict[str, str] = {
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'pdf': 'application/pdf',
    'csv': 'text/csv',
    'json': 'application/json',
    'html': 'text/html',
    'md': 'text/markdown',
}


def _mime_for_format(fmt: str) -> str:
    return _MIME_MAP.get(fmt.lstrip('.').lower(), 'application/octet-stream')


_notify_writer: Callable[[str], object] | None = None


def _get_notify_writer() -> Callable[[str], object] | None:
    return _notify_writer


def _set_notify_writer(writer: Callable[[str], object] | None) -> None:
    global _notify_writer
    _notify_writer = writer


def send_progress(progress: int, total: int, message: str = '') -> None:
    """Send an MCP progress notification if a writer is configured."""
    writer = _notify_writer
    if writer is None:
        return
    import json

    notification = json.dumps(
        {
            'jsonrpc': '2.0',
            'method': 'notifications/progress',
            'params': {
                'progress': progress,
                'total': total,
                'message': message,
            },
        },
        ensure_ascii=False,
    )
    writer(notification)


def _make_content(output_path: str, message: str) -> list[dict[str, Any]]:
    """Build MCP content array with text + resource URI."""
    path = Path(output_path)
    content: list[dict[str, Any]] = [{'type': 'text', 'text': message}]
    if path.exists():
        content.append(
            {
                'type': 'resource',
                'resource': {
                    'uri': path.resolve().as_uri(),
                    'mimeType': _mime_for_format(path.suffix),
                    'title': path.name,
                    'size': path.stat().st_size,
                },
            }
        )
    return content


class McpErrorCode:
    """Numeric error codes returned by MCP tools."""

    SUCCESS = 0
    DOCUMENT_NOT_FOUND = 1001
    DOCUMENT_LOCKED = 1002
    UNSUPPORTED_FORMAT = 1003
    TEMPLATE_ERROR = 1004
    CONVERSION_FAILED = 1005
    INVALID_PARAMETER = 1006
    EXCEL_INVALID_CELL_REF = 1007
    EXCEL_INVALID_RANGE = 1008
    PPT_INVALID_SLIDE_INDEX = 1009
    EXCEL_SHEET_NOT_FOUND = 1010
    INTERNAL_ERROR = 9999


#: Canonical documentation anchor for structured errors (docs/mcp/README.md,
#: "Error Handling"). Call sites attach it to refined error responses so Agents
#: can self-serve without an extra lookup.
DOCUMENTATION_URL = (
    'https://github.com/Tianshang301/TianshangScribe/blob/main/docs/mcp/README.md#error-handling'
)


ERROR_DESCRIPTIONS: dict[int, tuple[str, str]] = {
    McpErrorCode.DOCUMENT_NOT_FOUND: (
        'The document file was not found.',
        'Check the file path and ensure the file exists.',
    ),
    McpErrorCode.DOCUMENT_LOCKED: (
        'The document is password-protected.',
        'Provide the password or unlock the document first.',
    ),
    McpErrorCode.UNSUPPORTED_FORMAT: (
        'The document format is not supported.',
        'Use one of: docx, xlsx, pptx, pdf, csv, json, html, md.',
    ),
    McpErrorCode.TEMPLATE_ERROR: (
        'Template data could not be applied.',
        'Verify the template file and data structure.',
    ),
    McpErrorCode.CONVERSION_FAILED: (
        'Document conversion failed.',
        'Install office2pdf (~2MB) or LibreOffice for PDF conversion.',
    ),
    McpErrorCode.INVALID_PARAMETER: (
        'A required parameter is missing or invalid.',
        'Check the input schema and provide all required fields.',
    ),
    McpErrorCode.EXCEL_INVALID_CELL_REF: (
        'The Excel cell reference is invalid.',
        'Use A1-style references like "B2" (columns A-XFD, rows 1-1048576).',
    ),
    McpErrorCode.EXCEL_INVALID_RANGE: (
        'The Excel range is invalid.',
        'Use "START:END" form like "A1:C10", or row/column form ("2:5"/"B:D") for grouping.',
    ),
    McpErrorCode.PPT_INVALID_SLIDE_INDEX: (
        'The slide index is out of range.',
        'Use a 0-based index within the deck; extract_presentation_data reports '
        'the current slide count.',
    ),
    McpErrorCode.EXCEL_SHEET_NOT_FOUND: (
        'The target worksheet does not exist.',
        'Check the sheet name with analyze_excel_data, or create it with an '
        'add_sheet operation first.',
    ),
}


def error_response(
    error_code: int,
    detail: str = '',
    *,
    field: str | None = None,
    documentation_url: str | None = None,
) -> dict[str, Any]:
    """Build a structured error response with description and suggested fix.

    ``field`` and ``documentation_url`` are optional refinements: ``field``
    names the offending parameter (e.g. ``operations[2].range``) and
    ``documentation_url`` links to canonical docs. Both are included in the
    payload only when provided, keeping responses backward compatible.
    """
    desc, fix = ERROR_DESCRIPTIONS.get(
        error_code,
        ('An unexpected error occurred.', 'Try again or check the logs.'),
    )
    response: dict[str, Any] = {
        'success': False,
        'error_code': error_code,
        'error_message': desc + (' ' + detail if detail else ''),
        'suggested_fix': fix,
        'retryable': error_code != McpErrorCode.INTERNAL_ERROR,
    }
    if field is not None:
        response['field'] = field
    if documentation_url is not None:
        response['documentation_url'] = documentation_url
    return response


def success_response(
    data: dict[str, Any] | None = None,
    content: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Build a structured success response with optional data and content."""
    result: dict[str, Any] = {'success': True}
    if data:
        result['data'] = data
    if content:
        result['content'] = content
    return result
