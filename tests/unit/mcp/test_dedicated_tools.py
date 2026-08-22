"""Baseline tests for the dedicated MCP tools (v0.8.0 expansion)."""

from __future__ import annotations

from pathlib import Path

from tianshang_scribe.core.document import create_document, open_document, DocumentType
from tianshang_scribe.mcp.tools.analyze_excel import analyze_excel_data


def _make_workbook(path: Path) -> None:
    e = create_document(DocumentType.EXCEL)
    e.add_sheet('Data')
    ws = e.wb['Data']
    ws.append(['name', 'score', 'group'])
    ws.append(['alice', 10, 'a'])
    ws.append(['bob', 20, 'b'])
    ws.append(['alice', 10, 'a'])  # duplicate row
    ws.append(['carol', 30, 'a'])
    e.save(str(path))


def test_analyze_excel_data_baseline(tmp_path: Path) -> None:
    book = tmp_path / 'data.xlsx'
    _make_workbook(book)
    res = analyze_excel_data(str(book))
    assert res['success'] is True, res
    data = res['data']
    assert data['sheet_count'] == 2  # default 'Sheet' + 'Data'
    assert data['duplicate_row_count'] == 1
    sheet = next(s for s in data['sheets'] if s['name'] == 'Data')
    assert sheet['row_count'] == 4
    assert sheet['headers'] == ['name', 'score', 'group']
    by_name = {c['name']: c for c in sheet['columns']}
    assert by_name['score']['inferred_type'] == 'numeric'
    assert by_name['score']['numeric']['min'] == 10
    assert by_name['score']['numeric']['max'] == 30
    assert by_name['group']['inferred_type'] == 'categorical'
    assert by_name['name']['null_count'] == 0


def test_analyze_excel_rejects_non_excel(tmp_path: Path) -> None:
    from tianshang_scribe.core.document import DocumentType
    from tianshang_scribe.mcp.tools.analyze_excel import analyze_excel_data

    doc = tmp_path / 'a.docx'
    e = create_document(DocumentType.WORD)
    e.add_text('hello')
    e.save(str(doc))
    res = analyze_excel_data(str(doc))
    assert res['success'] is False
    assert res['error_code'] == 1003  # UNSUPPORTED_FORMAT
