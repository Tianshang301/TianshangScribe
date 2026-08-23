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
    headers: list[str] | None = Field(
        default=None, description='Column headers written as the first row.'
    )
    rows: list[list[Any]] | None = Field(
        default=None, description='Data rows (list of cell lists).'
    )
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
    - "write_cell": cell, value, sheet_name, style, is_formula
    - "set_formula": cell, formula, sheet_name
    - "freeze_panes": range (freeze anchor), sheet_name
    - "add_chart": chart_type, chart_data_range, sheet_name
    - "conditional_format": conditional_format (or range + cell_is opts)
    - "data_validation": data_validation (or range)
    - "add_table": headers, rows, sheet_name
    - "sort": range, key_columns, orders (or order)
    - "add_sheet": sheet_name
    - "set_range_style": range, style
    - "number_format": number_format ("RANGE=FORMAT")
    - "group_rows": range ("2:5"), outline_level, hidden, sheet_name
    - "group_columns": range ("B:D"), outline_level, hidden, sheet_name
    - "ungroup": range, axis ("rows"/"columns"), sheet_name
    - "set_tab_color": tab_color ("FF0000"), sheet_name
    - "set_print_area": range ("A1:C10"), sheet_name
    - "set_page_setup": paper_size, orientation, margins, header, footer, sheet_name
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
        'number_format',
        'group_rows',
        'group_columns',
        'ungroup',
        'set_tab_color',
        'set_print_area',
        'set_page_setup',
    ] = Field(description='Excel edit operation type (see class docstring for fields).')
    sheet_name: str | None = Field(default=None, description='Target worksheet.')
    cell: str | None = Field(default=None, description='Target cell, e.g. "A1".')
    value: Any | None = Field(default=None, description='Value to write (write_cell).')
    is_formula: bool | None = Field(
        default=None,
        description=(
            'write_cell only: true stores text as a formula (must start with "="), '
            'false forces a literal string even when it starts with "=", '
            'omitted keeps the automatic behaviour.'
        ),
    )
    formula: str | None = Field(default=None, description='Excel formula (set_formula).')
    range: str | None = Field(default=None, description='Cell range for freeze/chart/style/sort.')
    chart_type: str | None = Field(default=None, description='Chart type (bar/line/pie/...).')
    chart_data_range: str | None = Field(
        default=None, description='Source range for an Excel chart.'
    )
    conditional_format: str | None = Field(default=None, description='Conditional format spec.')
    data_validation: str | None = Field(default=None, description='Data validation spec.')
    headers: list[str] | None = Field(default=None, description='Table headers (add_table).')
    rows: list[list[Any]] | None = Field(default=None, description='Table data rows (add_table).')
    key_columns: list[int] | None = Field(default=None, description='0-based sort key columns.')
    orders: list[str] | None = Field(default=None, description='Per-key orders asc/desc.')
    order: str | None = Field(default=None, description='Single sort order asc/desc.')
    style: str | None = Field(
        default=None, description='Style string (write_cell/set_range_style).'
    )
    number_format: str | None = Field(
        default=None, description='Number format spec "RANGE=FORMAT".'
    )
    outline_level: int | None = Field(
        default=None,
        ge=1,
        le=7,
        description='Outline level 1-7 (group_rows/group_columns).',
    )
    hidden: bool | None = Field(
        default=None, description='Collapse the grouped rows/columns (group_*).'
    )
    axis: str | None = Field(default=None, description='"rows" or "columns" (ungroup).')
    tab_color: str | None = Field(
        default=None, description='Tab color as RGB hex, e.g. "FF0000" (set_tab_color).'
    )
    paper_size: str | int | None = Field(
        default=None, description='Paper size name (a4/letter/...) or raw int (set_page_setup).'
    )
    orientation: str | None = Field(
        default=None, description='"portrait" or "landscape" (set_page_setup).'
    )
    margins: dict[str, float] | None = Field(
        default=None,
        description='Margin inches, subset of left/right/top/bottom/header/footer (set_page_setup).',
    )
    header: str | None = Field(
        default=None, description='Centred page header text (set_page_setup).'
    )
    footer: str | None = Field(
        default=None, description='Centred page footer text (set_page_setup).'
    )


# --------------------------------------------------------------------------- #
# PowerPoint: create
# --------------------------------------------------------------------------- #
class PptTextBlock(BaseModel):
    """A precisely positioned text box on a slide."""

    model_config = _allow

    text: str = Field(description='Text content (LaTeX-style markup supported).')
    left: float = Field(default=1.0, description='Left position in inches.')
    top: float | None = Field(
        default=None,
        description=(
            'Top position in inches; when omitted the box auto-stacks below '
            'the previous block on the same slide instead of overlapping it.'
        ),
    )
    width: float | None = Field(
        default=None, description='Width in inches (defaults to slide width).'
    )
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
    width: float | None = Field(
        default=None, description='Width in inches (keeps ratio if only one set).'
    )
    height: float | None = Field(default=None, description='Height in inches.')


class PptMediaSpec(BaseModel):
    """A video or audio clip placed on a slide (click-to-play)."""

    model_config = _allow

    path: str = Field(
        description='Media file path (.mp4/.mov/.avi/.mkv video, .mp3/.wav/.m4a audio).'
    )
    kind: Literal['movie', 'audio'] = Field(
        default='movie', description='"movie" renders a video frame, "audio" a speaker icon.'
    )
    left: float = Field(default=1.0, description='Left position in inches.')
    top: float = Field(default=1.0, description='Top position in inches.')
    width: float | None = Field(
        default=None, description='Frame width in inches (movies; default 6.0).'
    )
    height: float | None = Field(
        default=None, description='Frame height in inches (movies; default 4.5).'
    )
    poster: str | None = Field(
        default=None, description='Poster frame image path shown before playback.'
    )


class PptSlideSpec(BaseModel):
    """One slide to build inside ``create_presentation``."""

    model_config = _allow

    layout: str | None = Field(
        default=None, description='Slide layout name, e.g. "Title and Content".'
    )
    title: str | None = Field(default=None, description='Slide title text.')
    bullets: list[str] | None = Field(
        default=None, description='Bullet points for the body placeholder.'
    )
    text_blocks: list[PptTextBlock] | None = Field(
        default=None, description='Precisely positioned text boxes.'
    )
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
    - "add_media": slide_index, media (path/kind/left/top/width/height/poster)
    - "add_shape": slide_index, shape_type, fill, line
    - "apply_layout": slide_index, layout
    - "set_transition": slide_index, transition
    - "add_notes": slide_index, notes
    - "replace_text": old_text, new_text
    - "apply_theme": theme ("office"/"dark")
    - "set_master_options": slide_number, footer_text, date_visible, date_text
    """

    model_config = _allow

    action: Literal[
        'add_slide',
        'add_text',
        'add_table',
        'add_chart',
        'add_picture',
        'add_media',
        'add_shape',
        'apply_layout',
        'set_transition',
        'add_notes',
        'replace_text',
        'apply_theme',
        'set_master_options',
    ] = Field(description='PowerPoint edit operation type (see class docstring for fields).')
    slide_index: int | None = Field(default=None, description='Target slide index (0-based).')
    text: str | None = Field(default=None, description='Text content (add_text/replace_text new).')
    old_text: str | None = Field(default=None, description='Text to find (replace_text).')
    new_text: str | None = Field(default=None, description='Replacement text (replace_text).')
    table: PptTableSpec | None = Field(default=None, description='Table spec (add_table).')
    chart: PptChartSpec | None = Field(default=None, description='Chart spec (add_chart).')
    picture: PptPictureSpec | None = Field(default=None, description='Picture spec (add_picture).')
    media: PptMediaSpec | None = Field(default=None, description='Media clip spec (add_media).')
    shape_type: str | None = Field(default=None, description='Autoshape type (add_shape).')
    fill: str | None = Field(default=None, description='Fill color (add_shape).')
    line: str | None = Field(default=None, description='Line color (add_shape).')
    layout: str | None = Field(default=None, description='Layout name (apply_layout/add_slide).')
    transition: str | None = Field(default=None, description='Transition name (set_transition).')
    notes: str | None = Field(default=None, description='Speaker notes (add_notes).')
    theme: str | None = Field(
        default=None, description='Built-in theme name: "office" or "dark" (apply_theme).'
    )
    slide_number: bool | None = Field(
        default=None, description='Show master-level slide numbers (set_master_options).'
    )
    footer_text: str | None = Field(
        default=None, description='Master-level footer text (set_master_options).'
    )
    date_visible: bool | None = Field(
        default=None, description='Show master-level date placeholder (set_master_options).'
    )
    date_text: str | None = Field(
        default=None,
        description='Fixed date text instead of an auto-updating field (set_master_options).',
    )
