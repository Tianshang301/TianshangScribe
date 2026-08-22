"""Pydantic input models for the TianshangScribe MCP tools.

The MCP SDK derives each tool's ``inputSchema`` from the tool function's
signature. Parameters annotated with these models (via ``Annotated[..., Field]``)
therefore produce rich, validated schemas: content blocks and edit operations
are validated against the ``Literal`` enums below, while unknown keys are kept
so the underlying document engines can still see them.
"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


def as_dict(value: Any) -> Any:
    """Recursively convert pydantic models to plain dicts.

    The MCP SDK validates tool arguments against the signature-derived model
    and passes validated sub-models straight into the tool function. The tools
    themselves work with plain dicts, so every model-valued parameter is
    normalised here before it reaches a tool body.
    """
    if isinstance(value, BaseModel):
        return value.model_dump()
    if isinstance(value, list):
        return [as_dict(v) for v in value]
    if isinstance(value, dict):
        return {key: as_dict(v) for key, v in value.items()}
    return value


class ContentBlock(BaseModel):
    """One element of the ordered ``content`` array used to build a document."""

    model_config = ConfigDict(extra='allow')

    type: Literal['paragraph', 'heading', 'formula', 'table', 'image', 'page_break'] = Field(
        description=(
            'Content block type:\n'
            '- "paragraph": plain or LaTeX-marked-up text\n'
            '- "heading": section heading with a level\n'
            '- "formula": LaTeX math formula (Word only)\n'
            '- "table": 2D array of rows\n'
            '- "image": image file inserted via its path\n'
            '- "page_break": page break'
        ),
    )
    text: str = Field(
        default='',
        description=(
            'Text content. Supports LaTeX-style markup and inline math $...$.\n'
            'Examples: \\bfseries{Bold}, \\itshape{Italic}, \\color{FF0000}{Red}, '
            '\\fontsize{20}{Title}, $x^2+y^2=1$.'
        ),
    )
    level: int | None = Field(
        default=None,
        ge=1,
        le=6,
        description='Heading level 1-6 (1 = document title). Only for type=heading.',
    )
    style: str | None = Field(
        default=None,
        description='Per-block style: font=Times,size=14,bold,color=FF0000,align=center.',
    )
    rows: list[list[Any]] | None = Field(
        default=None,
        description='Table data as a 2D array of cell values. Only for type=table.',
    )
    path: str | None = Field(
        default=None,
        description='Image file path. Only for type=image.',
    )
    # --- Excel / PPT specific (optional, backward compatible) ---
    sheet_name: str | None = Field(
        default=None,
        description='Target worksheet name (Excel). Routes write/formula/style to this sheet.',
    )
    cell: str | None = Field(
        default=None,
        description='Target cell reference, e.g. "A1" (Excel). Used by formula/hyperlink/write_cell.',
    )
    formula: str | None = Field(
        default=None,
        description='Excel formula string, e.g. "=SUM(B1:B10)". Requires cell.',
    )
    chart_type: str | None = Field(
        default=None,
        description='Chart type: bar, line, pie, area, doughnut, scatter (Excel/PPT).',
    )
    chart_data_range: str | None = Field(
        default=None,
        description='Excel chart data range, e.g. "Sheet1!A1:B10".',
    )
    chart_data: list[list[Any]] | None = Field(
        default=None,
        description='PPT chart data: first row series names, then [category, *values] rows.',
    )
    number_format: str | None = Field(
        default=None,
        description='Excel number format spec, e.g. "A1:A10=0.00%".',
    )
    conditional_format: str | None = Field(
        default=None,
        description='Excel conditional format spec, e.g. "B2:B100=color_scale" or "C1:C5=cell_is:greaterThan:20".',
    )
    data_validation: str | None = Field(
        default=None,
        description='Excel data validation spec, e.g. "C2:C50=list:yes,no" or "B1:B10=whole:1:100".',
    )
    freeze: str | None = Field(
        default=None,
        description='Excel freeze panes cell, e.g. "A2".',
    )
    hyperlink: str | None = Field(
        default=None,
        description='URL to hyperlink the cell given by `cell` to (Excel).',
    )
    named_range: str | None = Field(
        default=None,
        description='Excel named range spec, e.g. "MyRange=A1:B2".',
    )
    slide_index: int | None = Field(
        default=None,
        description='Target slide index (0-based) for PPT content (tables/charts/text). None = current/last slide.',
    )
    slide_layout: str | None = Field(
        default=None,
        description='PPT slide layout name or index (applied on slide creation).',
    )
    notes: str | None = Field(
        default=None,
        description='PPT speaker notes text for the target slide.',
    )
    transition: str | None = Field(
        default=None,
        description='PPT slide transition name, e.g. "fade".',
    )


class EditOperation(BaseModel):
    """One edit operation inside the ``operations`` array."""

    model_config = ConfigDict(extra='allow')

    action: Literal[
        'replace',
        'delete',
        'modify',
        'style',
        'add',
        'clear',
        'write_cell',
        'set_formula',
        'freeze_panes',
        'add_chart',
        'conditional_format',
        'data_validation',
        'add_table',
        'add_picture',
        'add_shape',
        'apply_layout',
        'set_transition',
        'add_notes',
    ] = Field(
        description=(
            'Operation type. Only the listed fields are meaningful per action:\n'
            '- "replace": old_text, new_text, regex (all doc types)\n'
            '- "delete": target, regex (all doc types)\n'
            '- "modify": old_text, new_text (all doc types)\n'
            '- "style": style, apply_all (all doc types)\n'
            '- "add": text, column (Excel append; Word/PPT append paragraph)\n'
            '- "clear": (no fields)\n'
            '- "write_cell": cell, text (value, formula if starts with "="), sheet_name, style (Excel)\n'
            '- "set_formula": cell, formula, sheet_name (Excel)\n'
            '- "freeze_panes": range (Excel)\n'
            '- "add_chart": chart_type + chart_data_range (Excel) OR chart_type + chart_data (PPT), sheet_name (Excel)\n'
            '- "conditional_format": conditional_format OR range + cell_is opts (Excel)\n'
            '- "data_validation": data_validation OR range (Excel)\n'
            '- "add_table": rows (first row = header), slide_index (PPT)\n'
            '- "add_picture": path, slide_index (PPT)\n'
            '- "add_shape": slide_index, fill, line (PPT)\n'
            '- "apply_layout": slide_index, layout (PPT)\n'
            '- "set_transition": slide_index, transition (PPT)\n'
            '- "add_notes": slide_index, notes (PPT)\n'
            'Unused fields for an action are ignored.'
        ),
    )
    old_text: str | None = Field(default=None, description='Text to find (replace/modify).')
    new_text: str | None = Field(default=None, description='Replacement text (replace/modify).')
    target: str | None = Field(default=None, description='Text to delete (delete).')
    text: str | None = Field(default=None, description='Text to add (add) or value to write (write_cell).')
    style: str | None = Field(default=None, description='Style string (style).')
    regex: bool | None = Field(default=None, description='Treat old_text/target as a regex.')
    apply_all: bool | None = Field(
        default=None,
        description='Apply style to every paragraph (style).',
    )
    column: int | None = Field(default=None, description='Target column, 1-based (add).')
    # --- Excel / PPT specific (optional) ---
    sheet_name: str | None = Field(default=None, description='Target worksheet (Excel).')
    cell: str | None = Field(default=None, description='Target cell reference, e.g. "A1" (Excel).')
    formula: str | None = Field(default=None, description='Excel formula string (set_formula).')
    range: str | None = Field(
        default=None,
        description='Cell range for freeze_panes/conditional_format/data_validation/add_chart (Excel).',
    )
    cf_type: str | None = Field(default=None, description='Conditional format type (color_scale/data_bar/cell_is/formula).')
    dv_type: str | None = Field(default=None, description='Data validation type (list/whole/decimal/date/text_length).')
    formula1: str | None = Field(default=None, description='Data validation formula1 (e.g. "yes,no").')
    formula2: str | None = Field(default=None, description='Data validation formula2 (e.g. upper bound).')
    chart_type: str | None = Field(default=None, description='Chart type for add_chart.')
    chart_data_range: str | None = Field(default=None, description='Excel chart data range for add_chart.')
    chart_data: list[list[Any]] | None = Field(default=None, description='PPT chart data for add_chart.')
    rows: list[list[Any]] | None = Field(
        default=None,
        description='Table rows for add_table (PPT). First row is the header.',
    )
    slide_index: int | None = Field(default=None, description='Target slide index (PPT add_table/add_chart).')
    layout: str | None = Field(default=None, description='Slide layout name/index (apply_layout).')
    transition: str | None = Field(default=None, description='Transition name (set_transition).')
    notes: str | None = Field(default=None, description='Speaker notes text (add_notes).')
    path: str | None = Field(default=None, description='Image path (add_picture).')
    fill: str | None = Field(default=None, description='Shape fill color hex (add_shape).')
    line: str | None = Field(default=None, description='Shape line color hex (add_shape).')


class ToolOptions(BaseModel):
    """Optional behavioural switches shared by several tools."""

    model_config = ConfigDict(extra='allow')

    action: Literal['compare', 'snapshot', 'list_snapshots', 'restore'] | None = Field(
        default=None,
        description=(
            'Sub-operation for compare_documents: "compare" (default) diffs two '
            'documents; "snapshot" records path_a state; "list_snapshots" lists '
            'recorded snapshots; "restore" writes a snapshot back to path_b.'
        ),
    )
    dry_run: bool | None = Field(
        default=None,
        description='Validate inputs and report the plan without writing any file.',
    )
    backup: bool | None = Field(
        default=None,
        description='Create a .bak copy of the target file before overwriting.',
    )
    snapshot_dir: str | None = Field(
        default=None,
        description='Directory to store/read document snapshots (default: ~/.tianshang-scribe/snapshots/).',
    )
    snapshot_id: str | None = Field(
        default=None,
        description='Snapshot identifier used by the restore action.',
    )
    deterministic_id: str | None = Field(
        default=None,
        description='Stable identifier attached to the operation result.',
    )
