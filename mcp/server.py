"""TianshangScribe MCP Server — stdio JSON-RPC entry point.

Protocol: MCP (Model Context Protocol) over stdio JSON-RPC 2.0

Usage:
  python -m mcp.server
  # Install: pip install -e ".[dev]"
"""

from __future__ import annotations

import json
import os
import sys

from mcp.errors import McpErrorCode, _set_notify_writer, error_response

SERVER_NAME = 'tianshang-scribe'
SERVER_VERSION = '0.2.0'

TOOLS = [
    {
        'name': 'create_office_document',
        'description': (
            'Create a Word, Excel, or PowerPoint document with structured content. '
            'Supports LaTeX-style formatting (e.g., \\bfseries{bold}, \\itshape{italic}) '
            'and math formulas (e.g., \\frac{a}{b}, \\sum_{i=0}^{n}). '
            'Use for generating reports, contracts, spreadsheets, or presentations.'
        ),
        'inputSchema': {
            'type': 'object',
            'properties': {
                'format': {
                    'type': 'string',
                    'enum': ['docx', 'xlsx', 'pptx'],
                    'default': 'docx',
                    'description': (
                        'Document format:\n'
                        '- "docx": Word document — reports, letters, contracts, proposals\n'
                        '- "xlsx": Excel workbook — spreadsheets, data tables, charts\n'
                        '- "pptx": PowerPoint — slides, presentations, pitch decks'
                    ),
                    'examples': ['docx', 'xlsx'],
                },
                'content': {
                    'type': 'array',
                    'description': 'Ordered list of content blocks forming the document.',
                    'items': {
                        'type': 'object',
                        'properties': {
                            'type': {
                                'type': 'string',
                                'enum': [
                                'paragraph', 'heading', 'formula',
                                'table', 'image', 'page_break',
                            ],
                                'description': 'Content block type.',
                            },
                            'text': {
                                'type': 'string',
                                'description': (
                                    'Text content. Supports LaTeX markup '
                                    'and inline math $...$.\n'
                                    'Examples:\n'
                                    '- \\bfseries{Bold text}\n'
                                    '- \\itshape{Italic text}\n'
                                    '- \\color{FF0000}{Red text}\n'
                                    '- \\fontsize{20}{Large title}\n'
                                    '- $x^2 + y^2 = 1$\n'
                                    '- $$\\sum_{i=0}^{n} x_i$$'
                                ),
                            },
                            'level': {
                                'type': 'integer',
                                'minimum': 1,
                                'maximum': 6,
                                'default': 1,
                                'description': (
                                    'Heading level 1-6. '
                                    'Level 1 = document title, '
                                    'Level 2 = chapter, Level 3 = section.'
                                ),
                            },
                            'style': {
                                'type': 'string',
                                'description': (
                            'Style: font=Times,size=14,bold,'
                            'color=FF0000,align=center'
                        ),
                            },
                            'rows': {
                                'type': 'array',
                                'description': 'Table data as 2D array. Only for type=table.',
                            },
                            'path': {
                                'type': 'string',
                                'description': 'Image file path. Only for type=image.',
                            },
                        },
                        'required': ['type'],
                    },
                },
                'template_data': {
                    'type': 'object',
                    'description': 'Key-value pairs to fill {{placeholder}} in content.',
                },
                'output_path': {
                    'type': 'string',
                    'description': 'Output file path.',
                },
                'style': {
                    'type': 'string',
                    'description': 'Global document style.',
                },
                'metadata': {
                    'type': 'object',
                    'description': 'Document metadata (title, author, etc.).',
                },
                'options': {
                    'type': 'object',
                    'properties': {
                        'dry_run': {'type': 'boolean'},
                        'backup': {'type': 'boolean'},
                        'deterministic_id': {'type': 'string'},
                    },
                },
            },
            'required': ['format', 'content'],
        },
    },
    {
        'name': 'edit_office_document',
        'description': (
            'Edit an existing Office document. Supports replace, delete, modify, style, '
            'add, and clear operations. Operations are applied in order.'
        ),
        'inputSchema': {
            'type': 'object',
            'properties': {
                'input_path': {
                    'type': 'string',
                    'description': 'Path to the existing document.',
                },
                'operations': {
                    'type': 'array',
                    'description': 'List of edit operations to apply.',
                    'items': {
                        'type': 'object',
                        'properties': {
                            'action': {
                                'type': 'string',
                                'enum': ['replace', 'delete', 'modify', 'style', 'add', 'clear'],
                            },
                            'old_text': {'type': 'string'},
                            'new_text': {'type': 'string'},
                            'target': {'type': 'string'},
                            'text': {'type': 'string'},
                            'style': {'type': 'string'},
                            'regex': {'type': 'boolean'},
                            'apply_all': {'type': 'boolean'},
                            'column': {'type': 'integer'},
                        },
                        'required': ['action'],
                    },
                },
                'output_path': {'type': 'string'},
                'options': {
                    'type': 'object',
                    'properties': {
                        'dry_run': {'type': 'boolean'},
                        'backup': {'type': 'boolean'},
                    },
                },
            },
            'required': ['input_path', 'operations'],
        },
    },
    {
        'name': 'fill_template',
        'description': (
            'Fill a document template with structured data. Replaces '
            '{{key}} placeholders and expands {{#each list}}...{{/each}} loops.'
        ),
        'inputSchema': {
            'type': 'object',
            'properties': {
                'template_path': {
                    'type': 'string',
                    'description': 'Path to the template document containing placeholders.',
                },
                'data': {
                    'type': 'object',
                    'description': 'Key-value data to fill placeholders. Supports nested objects.',
                },
                'output_path': {'type': 'string'},
                'options': {
                    'type': 'object',
                    'properties': {
                        'dry_run': {'type': 'boolean'},
                        'backup': {'type': 'boolean'},
                    },
                },
            },
            'required': ['template_path', 'data'],
        },
    },
    {
        'name': 'convert_document',
        'description': (
            'Convert a document between formats. Converts document content '
            'while preserving structure where possible.\n'
            '- Word (.docx) → PDF, Markdown, HTML\n'
            '- Excel (.xlsx) → PDF, CSV, JSON, HTML\n'
            '- PowerPoint (.pptx) → PDF'
        ),
        'inputSchema': {
            'type': 'object',
            'properties': {
                'input_path': {
                    'type': 'string',
                    'description': 'Path to the source document.',
                },
                'target_format': {
                    'type': 'string',
                    'enum': ['pdf', 'csv', 'json', 'html', 'md'],
                    'description': (
                        'Target output format:\n'
                        '- "pdf": PDF document (requires office2pdf or LibreOffice)\n'
                        '- "csv": Comma-separated values (Excel only)\n'
                        '- "json": JSON array of rows (Excel only)\n'
                        '- "html": HTML table (Excel) / styled HTML (Word)\n'
                        '- "md": Markdown (Word only)'
                    ),
                    'examples': ['pdf', 'md'],
                },
                'output_path': {'type': 'string'},
                'options': {
                    'type': 'object',
                    'properties': {
                        'dry_run': {'type': 'boolean'},
                    },
                },
            },
            'required': ['input_path', 'target_format'],
        },
    },
    {
        'name': 'extract_document_data',
        'description': (
            'Extract metadata, text, or structure from a document.'
        ),
        'inputSchema': {
            'type': 'object',
            'properties': {
                'input_path': {'type': 'string'},
                'mode': {
                    'type': 'string',
                    'enum': ['metadata', 'text', 'structure'],
                    'description': 'What to extract: metadata, text, or structure.',
                },
            },
            'required': ['input_path'],
        },
    },
    {
        'name': 'validate_template',
        'description': (
            'Validate a template document against provided data. '
            'Checks if all {{placeholder}} variables, {{#each}} loops, '
            'and {{#if}}/{{#unless}} conditions can be resolved. '
            'Use BEFORE calling fill_template to catch missing keys early.'
        ),
        'inputSchema': {
            'type': 'object',
            'properties': {
                'template_path': {
                    'type': 'string',
                    'description': 'Path to the template document (.docx/.xlsx).',
                },
                'data': {
                    'type': 'object',
                    'description': 'Key-value data to validate against placeholders.',
                },
                'options': {
                    'type': 'object',
                    'properties': {
                        'dry_run': {'type': 'boolean'},
                    },
                },
            },
            'required': ['template_path', 'data'],
        },
    },
]

RESOURCE_REGISTRY: dict[str, list[dict]] = {}


def _register_resource(session_id: str, uri: str, name: str,
                       mime_type: str = 'application/octet-stream',
                       description: str = '') -> None:
    """Register a document as a readable resource for a session."""
    if session_id not in RESOURCE_REGISTRY:
        RESOURCE_REGISTRY[session_id] = []
    RESOURCE_REGISTRY[session_id].append({
        'uri': uri,
        'name': name,
        'mimeType': mime_type,
        'description': description or f'Document: {name}',
    })


def _auto_register(session_id: str, result: dict) -> None:
    """Auto-register resource from tool result if it produced a file."""
    data = result.get('data', result)
    output_path = data.get('output_path', '') if isinstance(data, dict) else ''
    if output_path:
        from pathlib import Path

        from mcp.errors import _mime_for_format
        path = Path(output_path)
        if path.exists():
            _register_resource(
                session_id,
                path.resolve().as_uri(),
                path.name,
                _mime_for_format(path.suffix),
            )


def _dispatch_tool(name: str, args: dict) -> dict:
    try:
        if name == 'create_office_document':
            from mcp.tools.create import create_office_document
            return create_office_document(**args)
        elif name == 'edit_office_document':
            from mcp.tools.edit import edit_office_document
            return edit_office_document(**args)
        elif name == 'fill_template':
            from mcp.tools.template import fill_template
            return fill_template(**args)
        elif name == 'convert_document':
            from mcp.tools.convert import convert_document
            return convert_document(**args)
        elif name == 'extract_document_data':
            from mcp.tools.convert import extract_document_data
            return extract_document_data(**args)
        elif name == 'validate_template':
            from mcp.tools.validate import validate_template
            return validate_template(**args)
        else:
            return error_response(McpErrorCode.INVALID_PARAMETER, f"Unknown tool: {name}")
    except TypeError as e:
        return error_response(McpErrorCode.INVALID_PARAMETER, str(e))


def _handle_request(request: dict, _notify=None) -> dict | None:
    method = request.get('method', '')
    req_id = request.get('id')
    params = request.get('params', {})

    if method == 'initialize':
        return {
            'jsonrpc': '2.0',
            'id': req_id,
            'result': {
                'protocolVersion': '2024-11-05',
                'serverInfo': {
                    'name': SERVER_NAME,
                    'version': SERVER_VERSION,
                },
                'capabilities': {
                    'tools': {},
                    'resources': {'listChanged': False},
                },
            },
        }

    if method == 'tools/list':
        return {
            'jsonrpc': '2.0',
            'id': req_id,
            'result': {'tools': TOOLS},
        }

    if method == 'tools/call':
        tool_name = params.get('name', '')
        tool_args = params.get('arguments', {})
        if _notify:
            _set_notify_writer(_notify)
        try:
            result = _dispatch_tool(tool_name, tool_args)
        finally:
            if _notify:
                _set_notify_writer(None)
        session_id = params.get('_meta', {}).get('session_id', 'default')
        if result.get('success'):
            _auto_register(session_id, result)
        content_blocks = [
            {'type': 'text',
             'text': json.dumps(result, ensure_ascii=False, indent=2)},
        ]
        resource_blocks = result.get('content', [])
        for block in resource_blocks:
            if block.get('type') == 'resource':
                content_blocks.append(block)
        return {
            'jsonrpc': '2.0',
            'id': req_id,
            'result': {'content': content_blocks},
        }

    if method == 'resources/list':
        session_id = params.get('_meta', {}).get('session_id', 'default')
        resources = RESOURCE_REGISTRY.get(session_id, [])
        return {
            'jsonrpc': '2.0',
            'id': req_id,
            'result': {'resources': resources},
        }

    if method == 'resources/read':
        uri = params.get('uri', '')
        session_id = params.get('_meta', {}).get('session_id', 'default')
        resources = RESOURCE_REGISTRY.get(session_id, [])
        for res in resources:
            if res['uri'] == uri:
                from urllib.parse import urlparse
                parsed = urlparse(uri)
                file_path = parsed.path
                if file_path:
                    try:
                        from src.core.document import open_document
                        engine = open_document(file_path)
                        text_parts = []
                        if hasattr(engine, 'doc'):
                            text_parts = [
                                p.text for p in engine.doc.paragraphs
                                if p.text.strip()
                            ]
                        elif hasattr(engine, 'wb'):
                            text_parts = engine.wb.sheetnames
                        elif hasattr(engine, 'prs'):
                            text_parts = [
                                s.shapes[0].text_frame.text
                                for s in engine.prs.slides
                                if s.shapes
                            ]
                        return {
                            'jsonrpc': '2.0',
                            'id': req_id,
                            'result': {
                                'contents': [{
                                    'uri': uri,
                                    'mimeType': res.get('mimeType',
                                                        'text/plain'),
                                    'text': '\n'.join(text_parts),
                                }],
                            },
                        }
                    except Exception:
                        pass
                return {
                    'jsonrpc': '2.0',
                    'id': req_id,
                    'result': {
                        'contents': [{
                            'uri': uri,
                            'mimeType': res.get('mimeType',
                                                'text/plain'),
                            'text': f'[Binary resource: {res["name"]}]',
                        }],
                    },
                }
        return {
            'jsonrpc': '2.0',
            'id': req_id,
            'error': {
                'code': -32002,
                'message': f'Resource not found: {uri}',
            },
        }

    if method == 'notifications/initialized':
        return None

    return {
        'jsonrpc': '2.0',
        'id': req_id,
        'error': {'code': -32601, 'message': f'Method not found: {method}'},
    }


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description='TianshangScribe MCP Server',
    )
    parser.add_argument(
        '--transport',
        choices=['stdio', 'sse'],
        default='stdio',
        help='Transport protocol (default: stdio)',
    )
    parser.add_argument(
        '--host',
        default='127.0.0.1',
        help='SSE server host (default: 127.0.0.1)',
    )
    parser.add_argument(
        '--port',
        type=int,
        default=8080,
        help='SSE server port (default: 8080)',
    )
    parser.add_argument(
        '--auth-token',
        default=os.environ.get('SCRIBE_AUTH_TOKEN'),
        help='Bearer token for SSE auth (env: SCRIBE_AUTH_TOKEN)',
    )
    parser.add_argument(
        '--cors-origins',
        default=os.environ.get('SCRIBE_CORS_ORIGINS'),
        help='CORS allowed origins, comma-separated (env: SCRIBE_CORS_ORIGINS)',
    )
    args = parser.parse_args()

    if args.transport == 'sse':
        from mcp.transport_sse import run_sse
        run_sse(args.host, args.port, auth_token=args.auth_token,
                cors_origins=args.cors_origins)
        return

    import io

    sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

    while True:
        try:
            line = sys.stdin.readline()
            if not line:
                break

            request = json.loads(line.strip())
            response = _handle_request(
                request,
                _notify=lambda msg: sys.stdout.write(msg + '\n'),
            )

            if response is not None:
                sys.stdout.write(json.dumps(response, ensure_ascii=False) + '\n')
                sys.stdout.flush()

        except json.JSONDecodeError:
            continue
        except EOFError:
            break
        except Exception:
            err_resp = {
                'jsonrpc': '2.0',
                'id': None,
                'error': {'code': -32603, 'message': 'Internal error'},
            }
            sys.stdout.write(json.dumps(err_resp) + '\n')
            sys.stdout.flush()


if __name__ == '__main__':
    main()
