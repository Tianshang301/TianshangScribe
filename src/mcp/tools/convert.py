"""convert_document + extract_document_data — Format conversion and data extraction."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any, Literal

from pydantic import Field

from src.core.document import open_document
from src.mcp.errors import McpErrorCode, _make_content, error_response, success_response
from src.mcp.schemas import ToolOptions, as_dict


def convert_document(
    input_path: Annotated[str, Field(description='Path to the source document.')],
    target_format: Annotated[
        Literal['pdf', 'csv', 'json', 'html', 'md', 'markdown'],
        Field(
            description=(
                'Target output format:\n'
                '- "pdf": PDF document (requires office2pdf or LibreOffice)\n'
                '- "csv": Comma-separated values (Excel only)\n'
                '- "json": JSON array of rows (Excel only)\n'
                '- "html": HTML table (Excel) / styled HTML (Word)\n'
                '- "md"/"markdown": Markdown (Word only)'
            ),
            examples=['pdf', 'md'],
        ),
    ],
    output_path: Annotated[str, Field(description='Output file path.')] = '',
    options: Annotated[ToolOptions | None, Field(description='Tool options.')] = None,
) -> dict[str, Any]:
    """Convert a document to another format."""
    opts: dict[str, Any] = as_dict(options) or {}
    if not Path(input_path).exists():
        return error_response(McpErrorCode.DOCUMENT_NOT_FOUND, f"'{input_path}' not found.")

    fmt = target_format.lower()
    supported = {'pdf', 'csv', 'json', 'html', 'md', 'markdown'}
    if fmt not in supported:
        return error_response(
            McpErrorCode.UNSUPPORTED_FORMAT,
            f"Target format '{target_format}' not supported. Use: {', '.join(supported)}.",
        )

    stem = Path(input_path).stem
    output_path = output_path or f'{stem}.{fmt}'

    if opts.get('dry_run'):
        return success_response(
            {
                'dry_run': True,
                'from': input_path,
                'to': output_path,
                'format': fmt,
            }
        )

    try:
        engine = open_document(input_path)

        if fmt == 'pdf':
            engine.to_pdf(output_path)
        elif fmt == 'csv':
            engine.export_csv(output_path)
        elif fmt == 'json':
            engine.export_json(output_path)
        elif fmt == 'html':
            engine.export_html(output_path)
        elif fmt in ('md', 'markdown'):
            from src.transform.pdf import word_to_markdown

            engine.save()
            word_to_markdown(str(engine._path), str(output_path))

        return success_response(
            {
                'output_path': output_path,
                'source_format': Path(input_path).suffix,
                'target_format': fmt,
            },
            content=_make_content(output_path, f'Converted {input_path} → {output_path} ({fmt})'),
        )
    except NotImplementedError as e:
        return error_response(McpErrorCode.CONVERSION_FAILED, str(e))
    except Exception as e:
        return error_response(McpErrorCode.CONVERSION_FAILED, str(e))


def extract_document_data(
    input_path: Annotated[str, Field(description='Path to the source document.')],
    mode: Annotated[
        Literal['metadata', 'text', 'structure'],
        Field(description='What to extract: metadata, text, or structure.'),
    ] = 'metadata',
    options: Annotated[ToolOptions | None, Field(description='Tool options.')] = None,
) -> dict[str, Any]:
    """Extract data from a document."""
    if not Path(input_path).exists():
        return error_response(McpErrorCode.DOCUMENT_NOT_FOUND, f"'{input_path}' not found.")

    try:
        engine = open_document(input_path)

        if mode == 'metadata':
            meta = engine.get_metadata()
            return success_response({'metadata': meta})
        elif mode == 'text':
            text_parts = []
            if hasattr(engine, 'doc'):
                text_parts = [p.text for p in engine.doc.paragraphs if p.text.strip()]
            elif hasattr(engine, 'wb'):
                for ws in engine.wb.worksheets:
                    for row in ws.iter_rows(values_only=True):
                        text_parts.append(' | '.join(str(c) if c is not None else '' for c in row))
            elif hasattr(engine, 'prs'):
                for slide in engine.prs.slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            t = shape.text_frame.text.strip()
                            if t:
                                text_parts.append(t)
            return success_response(
                {
                    'text': '\n'.join(text_parts),
                    'text_blocks': len(text_parts),
                }
            )
        elif mode == 'structure':
            info = {}
            if hasattr(engine, 'doc'):
                info['paragraphs'] = len(engine.doc.paragraphs)
                info['sections'] = len(engine.doc.sections)
            elif hasattr(engine, 'wb'):
                info['sheets'] = engine.wb.sheetnames
                info['sheet_count'] = len(engine.wb.sheetnames)
            elif hasattr(engine, 'prs'):
                info['slides'] = len(engine.prs.slides)
            return success_response(info)
        else:
            return error_response(
                McpErrorCode.INVALID_PARAMETER,
                f"Extract mode '{mode}' not supported. Use: metadata, text, structure.",
            )
    except Exception as e:
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))
