"""TianshangScribe MCP Server — Structured error types."""

from __future__ import annotations


class McpErrorCode:
    SUCCESS = 0
    DOCUMENT_NOT_FOUND = 1001
    DOCUMENT_LOCKED = 1002
    UNSUPPORTED_FORMAT = 1003
    TEMPLATE_ERROR = 1004
    CONVERSION_FAILED = 1005
    INVALID_PARAMETER = 1006
    INTERNAL_ERROR = 9999


ERROR_DESCRIPTIONS = {
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
        'Ensure LibreOffice or pandoc is installed for this conversion.',
    ),
    McpErrorCode.INVALID_PARAMETER: (
        'A required parameter is missing or invalid.',
        'Check the input schema and provide all required fields.',
    ),
}


def error_response(error_code: int, detail: str = '') -> dict:
    desc, fix = ERROR_DESCRIPTIONS.get(
        error_code,
        ('An unexpected error occurred.', 'Try again or check the logs.'),
    )
    return {
        'success': False,
        'error_code': error_code,
        'error_message': desc + (' ' + detail if detail else ''),
        'suggested_fix': fix,
        'retryable': error_code != McpErrorCode.INTERNAL_ERROR,
    }


def success_response(data: dict | None = None) -> dict:
    return {'success': True, **({'data': data} if data else {})}
