"""Central tool registry for the TianshangScribe MCP Server.

Each entry binds a tool name to its implementation function, a rich
description (used verbatim as the ``description`` in ``tools/list``) and the
SDK annotations derived from :mod:`tianshang_scribe.mcp.security`.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, cast

from tianshang_scribe.mcp.tools.compare import compare_documents
from tianshang_scribe.mcp.tools.convert import convert_document, extract_document_data
from tianshang_scribe.mcp.tools.create import create_office_document
from tianshang_scribe.mcp.tools.edit import edit_office_document
from tianshang_scribe.mcp.tools.template import fill_template
from tianshang_scribe.mcp.tools.validate import validate_template

ToolEntry = dict[str, Any]

TOOLS: list[ToolEntry] = [
    {
        'name': 'create_office_document',
        'fn': create_office_document,
        'description': (
            'Create a NEW Word (.docx), Excel (.xlsx), or PowerPoint (.pptx) file '
            'from a structured content list of typed blocks (heading, paragraph, '
            'formula, table, image), with LaTeX-style markup (\\bfseries{bold}) and '
            'math formulas (\\frac{a}{b}). Writes to output_path (default: a temp '
            'file) and persists. To edit an existing file, use edit_office_document.\n'
            'Excel blocks may carry: sheet_name, cell+formula, freeze, '
            'chart_type+chart_data_range, number_format, conditional_format, '
            'data_validation, hyperlink, named_range. PPT blocks may carry: '
            'slide_index, slide_layout, notes, transition, chart_type+chart_data, '
            'rows (table), path (picture), fill/line (shape). Multiple PPT '
            'text/table/chart blocks stack onto one slide unless slide_index is set.'
        ),
    },
    {
        'name': 'edit_office_document',
        'fn': edit_office_document,
        'description': (
            'Modify an existing Office document by applying an ordered list of '
            'operations. Word/Excel/PPT share: replace, delete, modify, style, add, '
            'clear. Excel-only actions: write_cell (cell/text/sheet_name/style), '
            'set_formula (cell/formula/sheet_name), freeze_panes (range), add_chart '
            '(chart_type/chart_data_range/sheet_name), conditional_format, '
            'data_validation. PowerPoint-only actions: add_table (rows/slide_index), '
            'add_picture (path/slide_index), add_shape (fill/line/slide_index), '
            'apply_layout (layout/slide_index), set_transition (transition/slide_index), '
            'add_notes (notes/slide_index). Writes the result to output_path; when '
            'output_path is omitted the INPUT FILE IS OVERWRITTEN IN PLACE — set '
            'output_path or pass options {"backup": true} to keep a .bak copy. To '
            'generate a new document, use create_office_document.'
        ),
    },
    {
        'name': 'fill_template',
        'fn': fill_template,
        'description': (
            'Fill {{key}} placeholders in a template document with data, expanding '
            '{{#each list}} loops and {{#if}}/{{#unless}} conditions. Writes a NEW '
            'file to output_path (default: <template>_filled.<ext>); the input '
            'template is left unchanged. Run validate_template first to catch '
            'missing keys early.'
        ),
    },
    {
        'name': 'convert_document',
        'fn': convert_document,
        'description': (
            'Convert a document between formats while preserving structure where '
            'possible: Word (.docx) to PDF/Markdown/HTML, Excel (.xlsx) to '
            'PDF/CSV/JSON/HTML, PowerPoint (.pptx) to PDF. Writes the converted '
            'file to output_path (default: <stem>.<target_format>). PDF conversion '
            'requires office2pdf or LibreOffice to be installed. To read content '
            'for analysis instead, use extract_document_data.'
        ),
    },
    {
        'name': 'extract_document_data',
        'fn': extract_document_data,
        'description': (
            'Read data from a document: metadata (author/title/etc.), text (plain '
            'text plus a block count), or structure (paragraphs/sheets/slides). '
            'Read-only — never modifies the input file. To compare two documents, '
            'use compare_documents.'
        ),
    },
    {
        'name': 'validate_template',
        'fn': validate_template,
        'description': (
            'Validate that all {{placeholder}} variables, {{#each}} loops, and '
            '{{#if}}/{{#unless}} conditions in a template can be resolved against '
            'data, reporting missing keys. Read-only — never modifies the file. '
            'Call this BEFORE fill_template to catch missing keys early.'
        ),
    },
    {
        'name': 'compare_documents',
        'fn': compare_documents,
        'description': (
            'Compare two Word (.docx) documents and report paragraph-level '
            'differences: additions, removals, and changes with paragraph indices. '
            'Note: comparison and snapshots are ONLY supported for Word (.docx); '
            'Excel (.xlsx) and PowerPoint (.pptx) comparison is not yet available '
            '(you will receive an UNSUPPORTED_FORMAT error). Also manages document '
            'snapshots via options.action: "snapshot" records path_a state, '
            '"list_snapshots" lists recorded snapshots, and "restore" writes a '
            'snapshot back to path_b (snapshot store default: '
            '~/.tianshang-scribe/snapshots/). The compare mode never modifies its '
            'inputs; snapshot/restore write to the snapshot store. To read a single '
            'document, use extract_document_data.'
        ),
    },
]


def get_tools() -> list[ToolEntry]:
    """Return a copy of the tool registry."""
    return [dict(entry) for entry in TOOLS]


def get_tool(name: str) -> Callable[..., Any] | None:
    """Return the callable registered under ``name``, or ``None``."""
    for entry in TOOLS:
        if entry['name'] == name:
            return cast(Callable[..., Any], entry['fn'])
    return None
