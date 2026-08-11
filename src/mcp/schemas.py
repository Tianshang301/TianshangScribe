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


class EditOperation(BaseModel):
    """One edit operation inside the ``operations`` array."""

    model_config = ConfigDict(extra='allow')

    action: Literal['replace', 'delete', 'modify', 'style', 'add', 'clear'] = Field(
        description=(
            'Operation type:\n'
            '- "replace": find old_text, replace with new_text (regex optional)\n'
            '- "delete": remove text matching target\n'
            '- "modify": find old_text, replace with new_text\n'
            '- "style": apply a style string to the whole document\n'
            '- "add": append text at a column (Excel)\n'
            '- "clear": clear document content'
        ),
    )
    old_text: str | None = Field(default=None, description='Text to find (replace/modify).')
    new_text: str | None = Field(default=None, description='Replacement text (replace/modify).')
    target: str | None = Field(default=None, description='Text to delete (delete).')
    text: str | None = Field(default=None, description='Text to add (add).')
    style: str | None = Field(default=None, description='Style string (style).')
    regex: bool | None = Field(default=None, description='Treat old_text/target as a regex.')
    apply_all: bool | None = Field(
        default=None,
        description='Apply style to every paragraph (style).',
    )
    column: int | None = Field(default=None, description='Target column, 1-based (add).')


class ToolOptions(BaseModel):
    """Optional behavioural switches shared by several tools."""

    model_config = ConfigDict(extra='allow')

    dry_run: bool | None = Field(
        default=None,
        description='Validate inputs and report the plan without writing any file.',
    )
    backup: bool | None = Field(
        default=None,
        description='Create a .bak copy of the target file before overwriting.',
    )
    deterministic_id: str | None = Field(
        default=None,
        description='Stable identifier attached to the operation result.',
    )
