"""TianshangScribe MCP Server — stdio JSON-RPC entry point.

Protocol: MCP (Model Context Protocol) over stdio JSON-RPC 2.0

Usage:
  python -m mcp.server
  # Install: pip install -e ".[dev]"
"""

from __future__ import annotations

import json
import sys

from mcp.errors import McpErrorCode, error_response

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
                    'description': (
                        'Document format. docx for reports, xlsx for spreadsheets, '
                        'pptx for presentations.'
                    ),
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
                            'and inline math $...$.'
                        ),
                            },
                            'level': {
                                'type': 'integer',
                                'description': 'Heading level (1-6). Only for type=heading.',
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
            'Fill a document template with structured data. Replaces {{key}} placeholders '
            'and expands {{#each list}}...{{/each}} loops.'
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
            'Convert a document to another format. Supported: pdf, csv, json, html, md.'
        ),
        'inputSchema': {
            'type': 'object',
            'properties': {
                'input_path': {'type': 'string'},
                'target_format': {
                    'type': 'string',
                    'enum': ['pdf', 'csv', 'json', 'html', 'md'],
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
]


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
        else:
            return error_response(McpErrorCode.INVALID_PARAMETER, f"Unknown tool: {name}")
    except TypeError as e:
        return error_response(McpErrorCode.INVALID_PARAMETER, str(e))


def _handle_request(request: dict) -> dict | None:
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
        result = _dispatch_tool(tool_name, tool_args)
        return {
            'jsonrpc': '2.0',
            'id': req_id,
            'result': {
                'content': [{
                    'type': 'text',
                    'text': json.dumps(result, ensure_ascii=False, indent=2),
                }],
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
    args = parser.parse_args()

    if args.transport == 'sse':
        from mcp.transport_sse import run_sse
        run_sse(args.host, args.port)
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
            response = _handle_request(request)

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
