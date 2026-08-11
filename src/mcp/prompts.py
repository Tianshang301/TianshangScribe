"""Prompt templates registered on the TianshangScribe MCP Server.

Each prompt is a small orchestration hint that tells an agent how to chain the
document tools. The MCP SDK derives the prompt's ``arguments`` from the handler
function signature and converts the returned instruction string into a user
message.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Annotated

from pydantic import Field

if TYPE_CHECKING:
    from mcp.server.mcpserver import MCPServer


def _generate_report(
    topic: Annotated[str, Field(description='Report topic or title')],
    sections: Annotated[str, Field(description='Comma-separated section names')] = '',
    data_hint: Annotated[str, Field(description='Brief description of data to include')] = '',
) -> str:
    """Generate a professional Word report with table of contents, headings, data tables, and formatted text."""
    sections_hint = f' Sections: {sections}.' if sections else ''
    data_hint_text = f' Include data about: {data_hint}.' if data_hint else ''
    return (
        f'Create a professional Word document titled "{topic}". '
        'Include a table of contents, section headings, and formatted content. '
        'Use create_office_document with format="docx". '
        'Add a centered title heading, numbered sections, and a summary table. '
        f'Include metadata with author and date.{sections_hint}{data_hint_text}'
    )


def _batch_fill_templates(
    template_path: Annotated[
        str, Field(description='Path to the .docx template with {{placeholders}}')
    ],
    csv_path: Annotated[str, Field(description='Path to the CSV data file')],
    output_dir: Annotated[str, Field(description='Directory for generated documents')] = './output',
) -> str:
    """Fill a Word template with data from a CSV file. Each CSV row produces one output document."""
    return (
        f'Fill the template at "{template_path}" with data from CSV '
        f'at "{csv_path}". '
        'Extract CSV data, then call fill_template for each row. '
        f'Save output files to "{output_dir}".'
    )


def _convert_and_archive(
    input_pattern: Annotated[str, Field(description='Glob pattern or comma-separated file paths')],
    watermark: Annotated[str, Field(description='Optional watermark text for all PDFs')] = '',
) -> str:
    """Convert a batch of Office documents to PDF format."""
    watermark_hint = (
        f' If watermark is set, add text "{watermark}" to each PDF.' if watermark else ''
    )
    return (
        f'Convert documents matching "{input_pattern}" to PDF format. '
        'For each file, call convert_document with target_format="pdf". '
        f'If watermark "{watermark}" is provided, add it to each PDF '
        f'using edit_office_document.{watermark_hint}'
    )


def _extract_and_analyze(
    document_path: Annotated[str, Field(description='Path to the document to analyze')],
) -> str:
    """Extract metadata, structure, and text from a document for analysis or data migration."""
    return (
        f'Extract all data from "{document_path}". '
        'Call extract_document_data with mode="metadata", mode="structure", '
        'and mode="text". Summarize the findings in a structured report.'
    )


def _create_presentation(
    title: Annotated[str, Field(description='Presentation title')],
    outline: Annotated[str, Field(description='Slide-by-slide outline (one slide per line)')],
    theme: Annotated[str, Field(description='Optional: font/fontsize for styling')] = '',
) -> str:
    """Create a PowerPoint presentation from a text outline."""
    theme_hint = f' Style: {theme}.' if theme else ''
    return (
        f'Create a PowerPoint presentation titled "{title}" from this '
        f'outline:\n\n{outline}\n\n'
        'Use create_office_document with format="pptx". '
        'Each line becomes one slide: the text before ":" is the title, '
        f'the text after is the body.{theme_hint}'
    )


def register_prompts(server: MCPServer) -> None:
    """Register all prompts on the given MCP server."""
    from mcp.server.mcpserver.prompts import Prompt

    server.add_prompt(
        Prompt.from_function(
            _generate_report,
            name='generate_report',
            description=(
                'Generate a professional Word report with table of contents, '
                'headings, data tables, and formatted text.'
            ),
        )
    )
    server.add_prompt(
        Prompt.from_function(
            _batch_fill_templates,
            name='batch_fill_templates',
            description=(
                'Fill a Word template with data from a CSV file. '
                'Each CSV row produces one output document.'
            ),
        )
    )
    server.add_prompt(
        Prompt.from_function(
            _convert_and_archive,
            name='convert_and_archive',
            description='Convert a batch of Office documents to PDF format.',
        )
    )
    server.add_prompt(
        Prompt.from_function(
            _extract_and_analyze,
            name='extract_and_analyze',
            description=(
                'Extract metadata, structure, and text from a document for analysis or data migration.'
            ),
        )
    )
    server.add_prompt(
        Prompt.from_function(
            _create_presentation,
            name='create_presentation',
            description=(
                'Create a PowerPoint presentation from a text outline. '
                'Each top-level line becomes a slide title with body content.'
            ),
        )
    )
