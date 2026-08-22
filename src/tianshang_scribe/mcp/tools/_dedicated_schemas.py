"""Typed parameter models for the dedicated MCP tools.

These replace the generic ``ContentBlock`` / ``EditOperation`` models for the
document-type-specific tools (``create_excel_workbook``, ``edit_excel_workbook``,
``create_presentation``, ``edit_presentation``). Each model carries only the
fields relevant to its document type, giving Agents precise, discoverable
parameters instead of one mega-model with 20+ optional fields.
"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

_allow = ConfigDict(extra='allow')


# --------------------------------------------------------------------------- #
# Excel: create
# --------------------------------------------------------------------------- #
class ExcelSheetSpec(BaseModel):
    """One worksheet to build inside ``create_excel_workbook``."""

    model_config = _allow

    name: str = Field(description='Worksheet name.')
    headers: list[str] | None = Field(default=None, description='Column headers written as the first row.')
    rows: list[list[Any]] | None = Field(default=None, description='Data rows (list of cell lists).')
    formulas: dict[str, str] | None = Field(
        default=None, description='Cell->formula map, e.g. {"B2": "=SUM(A2:A10)"}.'
    )
    freeze: str | None = Field(default=None, description='Freeze panes at a cell, e.g. "A2".')
    number_format: str | None = Field(
        default=None, description='Number format "RANGE=FORMAT", e.g. "A1:A10=0.00%".'
    )
    conditional_format: str | None = Field(
        default=None, description='Conditional format spec, e.g. "B2:B100=color_scale".'
    )
    data_validation: str | None = Field(
        default=None, description='Data validation spec, e.g. "C1:C50=list:yes,no".'
    )
    column_widths: dict[str, float] | None = Field(
        default=None, description='Column letter -> width map, e.g. {"A": 20, "B": 12}.'
    )


# --------------------------------------------------------------------------- #
# Excel: edit
# --------------------------------------------------------------------------- #
class ExcelEditOp(BaseModel):
    """One Excel edit operation.

    Only the listed fields are meaningful per action:
    - "write_cell": cell, value (formula if starts with "="), sheet_name, style
    - "set_formula": cell, formula, sheet_name
    - "freeze_panes": range (freeze anchor), sheet_name
    - "add_chart": chart_type, chart_data_range, sheet_name
    - "conditional_format": conditional_format (or range + cell_is opts)
    - "data_validation": data_validation (or range)
    - "add_table": headers, rows, sheet_name
    - "sort": range, key_columns, orders (or order)
    - "add_sheet": sheet_name
    - "set_range_style": range, style
    """

    model_config = _allow

    action: Literal[
        'write_cell',
        'set_formula',
        'freeze_panes',
        'add_chart',
        'conditional_format',
        'data_validation',
        'add_table',
        'sort',
        'add_sheet',
        'set_range_style',
    ] = Field(description='Excel edit operation type (see class docstring for fields).')
    sheet_name: str | None = Field(default=None, description='Target worksheet.')
    cell: str | None = Field(default=None, description='Target cell, e.g. "A1".')
    value: Any | None = Field(default=None, description='Value to write (write_cell).')
    formula: str | None = Field(default=None, description='Excel formula (set_formula).')
    range: str | None = Field(default=None, description='Cell range for freeze/chart/style/sort.')
    chart_type: str | None = Field(default=None, description='Chart type (bar/line/pie/...).')
    chart_data_range: str | None = Field(default=None, description='Source range for an Excel chart.')
    conditional_format: str | None = Field(default=None, description='Conditional format spec.')
    data_validation: str | None = Field(default=None, description='Data validation spec.')
    headers: list[str] | None = Field(default=None, description='Table headers (add_table).')
    rows: list[list[Any]] | None = Field(default=None, description='Table data rows (add_table).')
    key_columns: list[int] | None = Field(default=None, description='0-based sort key columns.')
    orders: list[str] | None = Field(default=None, description='Per-key orders asc/desc.')
    order: str | None = Field(default=None, description='Single sort order asc/desc.')
    style: str | None = Field(default=None, description='Style string (write_cell/set_range_style).')


# --------------------------------------------------------------------------- #
# PowerPoint: create
# --------------------------------------------------------------------------- #
class PptTextBlock(BaseModel):
    """A precisely positioned text box on a slide."""

    model_config = _allow

    text: str = Field(description='Text content (LaTeX-style markup supported).')
    left: float = Field(default=1.0, description='Left position in inches.')
    top: float = Field(default=1.0, description='Top position in inches.')
    width: float | None = Field(default=None, description='Width in inches (defaults to slide width).')
    height: float = Field(default=1.0, description='Height in inches.')
    style: str | None = Field(default=None, description='Optional style string.')


class PptTableSpec(BaseModel):
    """A table placed on a slide."""

    model_config = _allow

    headers: list[str] = Field(description='Column headers.')
    rows: list[list[Any]] = Field(description='Data rows.')


class PptChartSpec(BaseModel):
    """A chart placed on a slide."""

    model_config = _allow

    chart_type: str = Field(description='Chart type: bar/column/line/pie/area/doughnut.')
    data: list[list[Any]] = Field(
        description='Chart data: first row holds series names (first cell ignored), '
        'each subsequent row is [category, *series values].'
    )


class PptPictureSpec(BaseModel):
    """A picture placed on a slide."""

    model_config = _allow

    path: str = Field(description='Image file path.')
    left: float = Field(default=1.0, description='Left position in inches.')
    top: float = Field(default=1.0, description='Top position in inches.')
    width: float | None = Field(default=None, description='Width in inches (keeps ratio if only one set).')
    height: float | None = Field(default=None, description='Height in inches.')


class PptSlideSpec(BaseModel):
    """One slide to build inside ``create_presentation``."""

    model_config = _allow

    layout: str | None = Field(default=None, description='Slide layout name, e.g. "Title and Content".')
    title: str | None = Field(default=None, description='Slide title text.')
    bullets: list[str] | None = Field(default=None, description='Bullet points for the body placeholder.')
    text_blocks: list[PptTextBlock] | None = Field(default=None, description='Precisely positioned text boxes.')
    table: PptTableSpec | None = Field(default=None, description='Table to add to the slide.')
    chart: PptChartSpec | None = Field(default=None, description='Chart to add to the slide.')
    picture: PptPictureSpec | None = Field(default=None, description='Picture to add to the slide.')
    notes: str | None = Field(default=None, description='Speaker notes.')
    transition: str | None = Field(default=None, description='Slide transition name, e.g. "fade".')


# --------------------------------------------------------------------------- #
# PowerPoint: edit
# --------------------------------------------------------------------------- #
class PptEditOp(BaseModel):
    """One PowerPoint edit operation.

    Only the listed fields are meaningful per action:
    - "add_slide": layout (optional)
    - "add_text": slide_index, text
    - "add_table": slide_index, table (headers+rows)
    - "add_chart": slide_index, chart (chart_type+data)
    - "add_picture": slide_index, picture (path/left/top/width/height)
    - "add_shape": slide_index, shape_type, fill, line
    - "apply_layout": slide_index, layout
    - "set_transition": slide_index, transition
    - "add_notes": slide_index, notes
    - "replace_text": old_text, new_text
    """

    model_config = _allow

    action: Literal[
        'add_slide',
        'add_text',
        'add_table',
        'add_chart',
        'add_picture',
        'add_shape',
        'apply_layout',
        'set_transition',
        'add_notes',
        'replace_text',
    ] = Field(description='PowerPoint edit operation type (see class docstring for fields).')
    slide_index: int | None = Field(default=None, description='Target slide index (0-based).')
    text: str | None = Field(default=None, description='Text content (add_text/replace_text new).')
    old_text: str | None = Field(default=None, description='Text to find (replace_text).')
    new_text: str | None = Field(default=None, description='Replacement text (replace_text).')
    table: PptTableSpec | None = Field(default=None, description='Table spec (add_table).')
    chart: PptChartSpec | None = Field(default=None, description='Chart spec (add_chart).')
    picture: PptPictureSpec | None = Field(default=None, description='Picture spec (add_picture).')
    shape_type: str | None = Field(default=None, description='Autoshape type (add_shape).')
    fill: str | None = Field(default=None, description='Fill color (add_shape).')
    line: str | None = Field(default=None, description='Line color (add_shape).')
    layout: str | None = Field(default=None, description='Layout name (apply_layout/add_slide).')
    transition: str | None = Field(default=None, description='Transition name (set_transition).')
    notes: str | None = Field(default=None, description='Speaker notes (add_notes).')
