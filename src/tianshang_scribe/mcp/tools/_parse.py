"""Parsing helpers for MCP tool arguments (Excel/PPT capability specs).

These are intentionally independent of the CLI package so the MCP layer does
not depend on ``tianshang_scribe.cli``. They turn the string specs accepted by
the new ``ContentBlock`` / ``EditOperation`` fields into engine call arguments.
"""

from __future__ import annotations

from typing import Any


def parse_number_format(spec: str) -> tuple[str, str]:
    """Parse ``"A1:A10=0.00%"`` into ``(range, fmt)``."""
    cell_range, _, fmt = spec.partition('=')
    if not cell_range or not fmt:
        raise ValueError(f'Invalid number_format spec: {spec!r} (expected RANGE=FORMAT)')
    return cell_range.strip(), fmt.strip()


def parse_conditional_format(spec: str) -> tuple[str, str, dict[str, str]]:
    """Parse ``"B2:B100=color_scale"`` or ``"C1:C5=cell_is:greaterThan:20"``.

    The remainder is split at most twice so a ``formula`` containing colons
    (e.g. a time literal like ``10:00``) survives intact.
    """
    cell_range, _, rest = spec.partition('=')
    if not cell_range or not rest:
        raise ValueError(f'Invalid conditional_format spec: {spec!r}')
    parts = rest.split(':', 2)
    cf_type = parts[0].strip()
    opts: dict[str, str] = {}
    if len(parts) > 1:
        opts['operator'] = parts[1].strip()
    if len(parts) > 2:
        opts['formula'] = parts[2].strip()
    return cell_range.strip(), cf_type, opts


def parse_data_validation(spec: str) -> tuple[str, str, str | None, str | None]:
    """Parse ``"C2:C50=list:yes,no"`` or ``"B1:B10=whole:1:100"``.

    The remainder is split at most twice (``dv_type``, ``formula1``,
    ``formula2``) so formulas containing colons stay in one piece.
    """
    cell_range, _, rest = spec.partition('=')
    if not cell_range or not rest:
        raise ValueError(f'Invalid data_validation spec: {spec!r}')
    parts = rest.split(':', 2)
    dv_type = parts[0].strip()
    formula1 = parts[1].strip() if len(parts) > 1 else None
    formula2 = parts[2].strip() if len(parts) > 2 else None
    return cell_range.strip(), dv_type, formula1, formula2


def parse_ppt_chart(spec: list[list[Any]]) -> list[list[Any]]:
    """Normalise PPT chart data.

    ``spec`` is a list of rows where ``spec[0]`` holds series names (its first
    cell is ignored) and each subsequent row is ``[category, *values]``. Returns
    the normalised data matching the shape expected by
    :meth:`PptEngine.add_chart` (first row ``[None, *series]``). The chart type
    is taken from the ``chart_type`` field of the calling tool, not inferred here.
    """
    if not spec:
        raise ValueError('chart_data must contain at least one row of series names')
    series = list(spec[0][1:]) if len(spec[0]) > 1 else []
    data: list[list[Any]] = [[None, *series]]
    for row in spec[1:]:
        data.append(list(row))
    return data


def resolve_slide_index(engine: Any, slide_index: int | None) -> int:
    """Resolve a slide index for PPT operations.

    ``None`` means the last slide; if the deck has no slides yet, one is created
    and its index returned.
    """
    slides = getattr(getattr(engine, 'prs', None), 'slides', None)
    if slides is None:
        raise ValueError('Engine does not support slides (not a presentation).')
    if slide_index is not None:
        if not (0 <= slide_index < len(slides)):
            raise IndexError(f'slide_index out of range: {slide_index}')
        return slide_index
    if len(slides) == 0:
        engine.add_slide()
        return len(engine.prs.slides) - 1
    return len(slides) - 1
