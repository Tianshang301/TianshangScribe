"""Tests for compare_excel_workbooks and extract_presentation_data (0.9.0 P0)."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from tianshang_scribe.mcp.tools.compare_excel import (
    MAX_CELL_DIFFS_PER_SHEET,
    compare_excel_workbooks,
)
from tianshang_scribe.mcp.tools.extract_ppt import extract_presentation_data
from tianshang_scribe.mcp.tools.ppt_create import create_presentation


def _wb(path: Path, sheets: dict[str, list[list]]) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(str(path))


# --------------------------------------------------------------------------- #
# compare_excel_workbooks
# --------------------------------------------------------------------------- #
class TestCompareExcelWorkbooks:
    def test_structure_baseline(self, tmp_path: Path) -> None:
        a, b = tmp_path / 'a.xlsx', tmp_path / 'b.xlsx'
        _wb(a, {'Base': [[1]], 'Extra': [[2]]})
        _wb(b, {'Base': [[1], [2]], 'New': [[3]]})
        res = compare_excel_workbooks(str(a), str(b), mode='structure')
        assert res['success'] is True, res
        data = res['data']
        assert data['sheets']['added'] == ['New']
        assert data['sheets']['removed'] == ['Extra']
        assert data['sheets']['renamed'] == []
        base_dim = next(d for d in data['dims_changed'] if d['sheet'] == 'Base')
        assert base_dim['a']['rows'] == 1 and base_dim['b']['rows'] == 2

    def test_structure_no_change(self, tmp_path: Path) -> None:
        a, b = tmp_path / 'a.xlsx', tmp_path / 'b.xlsx'
        _wb(a, {'S': [[1, 2]]})
        _wb(b, {'S': [[1, 2]]})
        res = compare_excel_workbooks(str(a), str(b), mode='structure')
        data = res['data']
        assert data['sheets'] == {'added': [], 'removed': [], 'renamed': []}
        assert data['dims_changed'] == []

    def test_renamed_via_sheet_mapping(self, tmp_path: Path) -> None:
        a, b = tmp_path / 'a.xlsx', tmp_path / 'b.xlsx'
        _wb(a, {'Data': [[1, 2], [3, 4]]})
        _wb(b, {'Data2024': [[1, 2], [3, 4]]})
        res = compare_excel_workbooks(
            str(a), str(b), mode='data', sheet_mapping={'Data2024': 'Data'}
        )
        data = res['data']
        assert data['sheets']['renamed'] == [{'from': 'Data2024', 'to': 'Data'}]
        assert data['sheets']['added'] == [] and data['sheets']['removed'] == []
        assert data['cell_diff_count'] == 0
        assert data['identical_sheets'] == ['Data']

    def test_data_mode_value_diffs_and_tolerance(self, tmp_path: Path) -> None:
        a, b = tmp_path / 'a.xlsx', tmp_path / 'b.xlsx'
        _wb(a, {'S': [['k', 10], ['x', 'same']]})
        _wb(b, {'S': [['k', 10.004], ['x', 'same']]})
        strict = compare_excel_workbooks(str(a), str(b), mode='data')
        assert strict['success'] is True
        assert strict['data']['cell_diff_count'] == 1
        cell = strict['data']['cells'][0]
        assert (cell['sheet'], cell['cell']) == ('S', 'B1')
        assert cell['before'] == 10 and cell['after'] == 10.004

        loose = compare_excel_workbooks(str(a), str(b), mode='data', tolerance=0.01)
        assert loose['data']['cell_diff_count'] == 0
        # values are tolerance-equal but NOT hash-identical, so no prefilter hit
        assert loose['data']['identical_sheets'] == []
        assert loose['data']['cells'] == []

    def test_missing_cell_becomes_none_diff(self, tmp_path: Path) -> None:
        a, b = tmp_path / 'a.xlsx', tmp_path / 'b.xlsx'
        _wb(a, {'S': [[1, 2]]})
        _wb(b, {'S': [[1]]})
        res = compare_excel_workbooks(str(a), str(b), mode='data')
        cells = res['data']['cells']
        assert cells == [{'sheet': 'S', 'cell': 'B1', 'kind': 'value', 'before': 2, 'after': None}]

    def test_truncation_cap(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        a, b = tmp_path / 'a.xlsx', tmp_path / 'b.xlsx'
        _wb(a, {'S': [[i] for i in range(20)]})
        _wb(b, {'S': [[i + 100] for i in range(20)]})
        monkeypatch.setattr('tianshang_scribe.mcp.tools.compare_excel.MAX_CELL_DIFFS_PER_SHEET', 5)
        res = compare_excel_workbooks(str(a), str(b), mode='data')
        data = res['data']
        assert data['truncated'] is True
        assert data['cell_diff_count'] == 5
        assert MAX_CELL_DIFFS_PER_SHEET == 10_000  # real constant untouched

    def test_error_missing_file(self, tmp_path: Path) -> None:
        res = compare_excel_workbooks(str(tmp_path / 'nope.xlsx'), str(tmp_path / 'nope2.xlsx'))
        assert res['success'] is False
        assert res['error_code'] == 1001

    def test_error_unsupported_format(self, tmp_path: Path) -> None:
        doc = tmp_path / 'a.docx'
        doc.write_text('x')
        res = compare_excel_workbooks(str(doc), str(doc))
        assert res['success'] is False
        assert res['error_code'] == 1003

    def test_error_formula_mode_pending(self, tmp_path: Path) -> None:
        a, b = tmp_path / 'a.xlsx', tmp_path / 'b.xlsx'
        _wb(a, {'S': [[1]]})
        _wb(b, {'S': [[1]]})
        res = compare_excel_workbooks(str(a), str(b), mode='formula')
        assert res['success'] is False
        assert res['error_code'] == 1006
        assert 'not available yet' in res['error_message']


# --------------------------------------------------------------------------- #
# extract_presentation_data
# --------------------------------------------------------------------------- #
class TestExtractPresentationData:
    def _deck(self, path: Path) -> None:
        res = create_presentation(
            str(path),
            slides=[
                {
                    'title': 'Q3 Review',
                    'bullets': ['Revenue +12%', 'Churn -2%'],
                    'notes': 'Pause here',
                    'transition': 'fade',
                }
            ],
        )
        assert res['success'] is True, res

    def test_outline_mode(self, tmp_path: Path) -> None:
        deck = tmp_path / 'deck.pptx'
        self._deck(deck)
        res = extract_presentation_data(str(deck), mode='outline')
        assert res['success'] is True, res
        slide = res['data']['slides'][0]
        assert slide['index'] == 0
        assert slide['title'] == 'Q3 Review'
        assert slide['bullets'] == ['Revenue +12%', 'Churn -2%']
        assert slide['notes'] == 'Pause here'
        assert slide['transition'] == 'fade'
        assert slide['layout']

    def test_structure_mode(self, tmp_path: Path) -> None:
        deck = tmp_path / 'deck.pptx'
        create_presentation(
            str(deck),
            slides=[
                {
                    'title': 'Mixed',
                    'text_blocks': [{'text': 'tb1'}],
                    'table': {'headers': ['H'], 'rows': [['v']]},
                }
            ],
        )
        res = extract_presentation_data(str(deck), mode='structure')
        assert res['success'] is True, res
        data = res['data']
        slide = data['slides'][0]
        assert slide['text'] >= 1 and slide['table'] == 1
        assert data['totals']['table'] == 1
        assert data['slide_count'] == 1

    def test_error_missing_file_and_bad_ext(self, tmp_path: Path) -> None:
        missing = extract_presentation_data(str(tmp_path / 'nope.pptx'))
        assert missing['success'] is False
        assert missing['error_code'] == 1001

        doc = tmp_path / 'a.docx'
        doc.write_text('x')
        bad = extract_presentation_data(str(doc))
        assert bad['success'] is False
        assert bad['error_code'] == 1003

    def test_error_notes_mode_pending(self, tmp_path: Path) -> None:
        deck = tmp_path / 'deck.pptx'
        self._deck(deck)
        res = extract_presentation_data(str(deck), mode='notes')
        assert res['success'] is False
        assert res['error_code'] == 1006
        assert 'not available yet' in res['error_message']
