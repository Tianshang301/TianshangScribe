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

PROMPTS = [
    {
        'name': 'generate_report',
        'description': (
            'Generate a professional Word report with table of contents, '
            'headings, data tables, and formatted text.'
        ),
        'arguments': [
            {'name': 'topic', 'description': 'Report topic or title',
             'required': True},
            {'name': 'sections',
             'description': 'Comma-separated section names',
             'required': False},
            {'name': 'data_hint',
             'description': 'Brief description of data to include',
             'required': False},
        ],
    },
    {
        'name': 'batch_fill_templates',
        'description': (
            'Fill a Word template with data from a CSV file. '
            'Each CSV row produces one output document.'
        ),
        'arguments': [
            {'name': 'template_path',
             'description': 'Path to the .docx template with {{placeholders}}',
             'required': True},
            {'name': 'csv_path',
             'description': 'Path to the CSV data file',
             'required': True},
            {'name': 'output_dir',
             'description': 'Directory for generated documents',
             'required': False},
        ],
    },
    {
        'name': 'convert_and_archive',
        'description': (
            'Convert a batch of Office documents to PDF format.'
        ),
        'arguments': [
            {'name': 'input_pattern',
             'description': 'Glob pattern or comma-separated file paths',
             'required': True},
            {'name': 'watermark',
             'description': 'Optional watermark text for all PDFs',
             'required': False},
        ],
    },
    {
        'name': 'extract_and_analyze',
        'description': (
            'Extract metadata, structure, and text from a document '
            'for analysis or data migration.'
        ),
        'arguments': [
            {'name': 'document_path',
             'description': 'Path to the document to analyze',
             'required': True},
        ],
    },
    {
        'name': 'create_presentation',
        'description': (
            'Create a PowerPoint presentation from a text outline. '
            'Each top-level line becomes a slide title with body content.'
        ),
        'arguments': [
            {'name': 'title', 'description': 'Presentation title',
             'required': True},
            {'name': 'outline',
             'description': 'Slide-by-slide outline (one slide per line)',
             'required': True},
            {'name': 'theme',
             'description': 'Optional: font/fontsize for styling',
             'required': False},
        ],
    },
]

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
    {
        'name': 'compare_documents',
        'description': (
            'Compare two Word (.docx) documents and report paragraph-level '
            'differences: additions, removals, and changes. Returns a '
            'structured diff with paragraph indices.'
        ),
        'inputSchema': {
            'type': 'object',
            'properties': {
                'path_a': {
                    'type': 'string',
                    'description': 'Path to the first document.',
                },
                'path_b': {
                    'type': 'string',
                    'description': 'Path to the second document.',
                },
            },
            'required': ['path_a', 'path_b'],
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
        elif name == 'compare_documents':
            from mcp.tools.compare import compare_documents
            return compare_documents(**args)
        else:
            return error_response(McpErrorCode.INVALID_PARAMETER, f"Unknown tool: {name}")
    except TypeError as e:
        return error_response(McpErrorCode.INVALID_PARAMETER, str(e))


_PROMPT_TEMPLATES = {
    'generate_report': {
        'role': 'user',
        'content': (
            'Create a professional Word document titled "{topic}". '
            'Include a table of contents, section headings, and formatted content. '
            'Use create_office_document with format="docx". '
            'Add a centered title heading, numbered sections, and a summary table. '
            'Include metadata with author and date.{sections_hint}{data_hint}'
        ),
    },
    'batch_fill_templates': {
        'role': 'user',
        'content': (
            'Fill the template at "{template_path}" with data from CSV '
            'at "{csv_path}". '
            'Extract CSV data, then call fill_template for each row. '
            'Save output files to "{output_dir}".'
        ),
    },
    'convert_and_archive': {
        'role': 'user',
        'content': (
            'Convert documents matching "{input_pattern}" to PDF format. '
            'For each file, call convert_document with target_format="pdf". '
            'If watermark "{watermark}" is provided, add it to each PDF '
            'using edit_office_document.{watermark_hint}'
        ),
    },
    'extract_and_analyze': {
        'role': 'user',
        'content': (
            'Extract all data from "{document_path}". '
            'Call extract_document_data with mode="metadata", mode="structure", '
            'and mode="text". Summarize the findings in a structured report.'
        ),
    },
    'create_presentation': {
        'role': 'user',
        'content': (
            'Create a PowerPoint presentation titled "{title}" from this '
            'outline:\n\n{outline}\n\n'
            'Use create_office_document with format="pptx". '
            'Each line becomes one slide: the text before ":" is the title, '
            'the text after is the body.{theme_hint}'
        ),
    },
}


def _build_prompt_messages(prompt: dict,
                           arguments: dict[str, str]) -> list[dict]:
    """Build message array from a prompt template and user arguments."""
    template = _PROMPT_TEMPLATES.get(prompt['name'], {})
    if not template:
        return []

    content = template.get('content', '')

    if prompt['name'] == 'generate_report':
        content = content.format(
            topic=arguments.get('topic', 'Report'),
            sections_hint=(
                f' Sections: {arguments["sections"]}.'
                if arguments.get('sections') else ''
            ),
            data_hint=(
                f' Include data about: {arguments["data_hint"]}.'
                if arguments.get('data_hint') else ''
            ),
        )
    elif prompt['name'] == 'batch_fill_templates':
        content = content.format(
            template_path=arguments.get('template_path', 'template.docx'),
            csv_path=arguments.get('csv_path', 'data.csv'),
            output_dir=arguments.get('output_dir', './output'),
        )
    elif prompt['name'] == 'convert_and_archive':
        watermark = arguments.get('watermark', '')
        content = content.format(
            input_pattern=arguments.get('input_pattern', '*.docx'),
            watermark=watermark,
            watermark_hint=(
                f' If watermark is set, add text "{watermark}" to each PDF.'
                if watermark else ''
            ),
        )
    elif prompt['name'] == 'extract_and_analyze':
        content = content.format(
            document_path=arguments.get('document_path', 'document.docx'),
        )
    elif prompt['name'] == 'create_presentation':
        content = content.format(
            title=arguments.get('title', 'Presentation'),
            outline=arguments.get('outline', 'Slide 1\nSlide 2'),
            theme_hint=(
                f' Style: {arguments["theme"]}.'
                if arguments.get('theme') else ''
            ),
        )

    return [{'role': template.get('role', 'user'), 'content': content}]


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
                    'prompts': {},
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

    if method == 'prompts/list':
        return {
            'jsonrpc': '2.0',
            'id': req_id,
            'result': {
                'prompts': [
                    {'name': p['name'], 'description': p['description']}
                    for p in PROMPTS
                ],
            },
        }

    if method == 'prompts/get':
        name = params.get('name', '')
        arguments = params.get('arguments', {})
        for p in PROMPTS:
            if p['name'] == name:
                messages = _build_prompt_messages(p, arguments)
                return {
                    'jsonrpc': '2.0',
                    'id': req_id,
                    'result': {
                        'description': p['description'],
                        'messages': messages,
                    },
                }
        return {
            'jsonrpc': '2.0',
            'id': req_id,
            'error': {
                'code': -32602,
                'message': f'Prompt not found: {name}',
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
