"""Baseline tests for MCP parsing helpers (P0-2)."""

from __future__ import annotations

from tianshang_scribe.core.ppt_engine import PptEngine
from tianshang_scribe.mcp.tools._parse import (
    parse_conditional_format,
    parse_data_validation,
    parse_number_format,
    parse_ppt_chart,
    resolve_slide_index,
)


def test_parse_number_format() -> None:
    assert parse_number_format('A1:A10=0.00%') == ('A1:A10', '0.00%')


def test_parse_conditional_format() -> None:
    assert parse_conditional_format('B2:B100=color_scale') == ('B2:B100', 'color_scale', {})
    assert parse_conditional_format('C1:C5=cell_is:greaterThan:20') == (
        'C1:C5',
        'cell_is',
        {'operator': 'greaterThan', 'formula': '20'},
    )


def test_parse_data_validation() -> None:
    assert parse_data_validation('C2:C50=list:yes,no') == ('C2:C50', 'list', 'yes,no', None)
    assert parse_data_validation('B1:B10=whole:1:100') == ('B1:B10', 'whole', '1', '100')


def test_parse_ppt_chart() -> None:
    data = parse_ppt_chart([['', 'S1', 'S2'], ['Cat1', 1, 2], ['Cat2', 3, 4]])
    assert data[0] == [None, 'S1', 'S2']
    assert data[1] == ['Cat1', 1, 2]
    assert data[2] == ['Cat2', 3, 4]


def test_resolve_slide_index_last_when_none() -> None:
    e = PptEngine()
    e.create()
    e.add_slide()
    e.add_slide()
    assert resolve_slide_index(e, None) == 1
    assert resolve_slide_index(e, 0) == 0


def test_resolve_slide_index_creates_when_empty() -> None:
    e = PptEngine()
    e.create()
    assert resolve_slide_index(e, None) == 0
    assert len(e.prs.slides) == 1


def test_resolve_slide_index_out_of_range() -> None:
    e = PptEngine()
    e.create()
    e.add_slide()
    try:
        resolve_slide_index(e, 9)
        raise AssertionError('expected IndexError')
    except IndexError:
        pass
