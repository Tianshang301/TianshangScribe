"""create_presentation — Build a new .pptx presentation with typed slide specs.

Document-type-specific tool that mirrors ``create_office_document`` for
PowerPoint only, using a precise ``PptSlideSpec`` model instead of the generic
``ContentBlock`` mega-model. It builds the deck directly for discoverable,
slide-by-slide parameters.
"""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from tianshang_scribe.core.document import DocumentType, create_document
from tianshang_scribe.mcp.errors import McpErrorCode, error_response, success_response
from tianshang_scribe.mcp.schemas import ToolOptions, as_dict
from tianshang_scribe.mcp.tools._dedicated_schemas import PptSlideSpec
from tianshang_scribe.utils.file_utils import ensure_parent_dir


def create_presentation(
    output_path: Annotated[str, Field(description='Output .pptx path to create.')],
    slides: Annotated[list[PptSlideSpec], Field(description='Slides to build (in order).')],
    metadata: Annotated[
        dict[str, Any] | None, Field(description='Optional document properties (title/author/...).')
    ] = None,
    options: Annotated[ToolOptions | None, Field(description='Tool options.')] = None,
) -> dict[str, Any]:
    """Create a brand new PowerPoint presentation from typed slide specifications.

    Side effects: writes a new ``.pptx`` file at ``output_path``. Use
    ``edit_presentation`` to modify an existing deck.
    """
    specs: list[dict[str, Any]] = as_dict(slides)
    if not output_path.lower().endswith('.pptx'):
        return error_response(
            McpErrorCode.INVALID_PARAMETER,
            'create_presentation only writes .pptx presentations.',
        )
    if not specs:
        return error_response(
            McpErrorCode.INVALID_PARAMETER,
            'create_presentation requires at least one slide spec.',
        )

    try:
        engine: Any = create_document(DocumentType.PPT)
        built: list[int] = []
        for raw in specs:
            engine.add_slide()
            idx = len(engine.prs.slides) - 1 if hasattr(engine, 'prs') else 0
            if raw.get('layout'):
                engine.apply_layout(idx, raw['layout'])
            if raw.get('title'):
                engine.add_text(raw['title'], slide_index=idx)
            for bullet in raw.get('bullets') or []:
                engine.add_text(bullet, slide_index=idx)
            for tb in raw.get('text_blocks') or []:
                engine.add_textbox(
                    idx,
                    tb.get('text', ''),
                    left=tb.get('left', 1.0),
                    top=tb.get('top', 1.0),
                    width=tb.get('width'),
                    height=tb.get('height', 1.0),
                )
            if raw.get('table'):
                tbl = raw['table']
                engine.add_table(
                    idx,
                    [tbl.get('headers', [])] + (tbl.get('rows') or []),
                    col_names=tbl.get('headers'),
                )
            if raw.get('chart'):
                ch = raw['chart']
                engine.add_chart(idx, ch.get('chart_type'), ch.get('data'))
            if raw.get('picture'):
                pic = raw['picture']
                engine.add_picture(
                    idx,
                    pic.get('path'),
                    left=pic.get('left', 1.0),
                    top=pic.get('top', 1.0),
                    width=pic.get('width'),
                    height=pic.get('height'),
                )
            if raw.get('notes'):
                engine.add_notes(idx, raw['notes'])
            if raw.get('transition'):
                engine.set_transition(raw['transition'], slide_index=idx)
            built.append(idx)

        if metadata and hasattr(engine, 'set_metadata'):
            engine.set_metadata(**metadata)
        ensure_parent_dir(output_path)
        engine.save(output_path)
        return success_response(
            {'output_path': output_path, 'slides': built, 'slide_count': len(built)}
        )
    except Exception as e:  # surface any build failure as a structured error
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))
