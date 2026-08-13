"""Unit tests for the Excel engine."""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from src.core.document import open_document
from src.core.excel_engine import ExcelEngine


class TestExcelEngine:
    @pytest.fixture
    def engine(self) -> ExcelEngine:
        e = ExcelEngine()
        e.create()
        return e

    def test_create_empty(self, engine: ExcelEngine) -> None:
        ws = engine.wb.active
        assert ws is not None

    def test_add_text(self, engine: ExcelEngine) -> None:
        engine.add_text('Cell Content')
        ws = engine.wb.active
        assert ws.cell(row=1, column=1).value == 'Cell Content'

    def test_save_and_reopen(self, engine: ExcelEngine) -> None:
        engine.add_text('Excel Test')
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / 'test.xlsx'
            engine.save(path)
            assert path.exists()
            reopened = open_document(path)
            assert isinstance(reopened, ExcelEngine)
            ws = reopened.wb.active
            assert ws.cell(row=1, column=1).value == 'Excel Test'

    def test_add_text_with_column(self, engine: ExcelEngine) -> None:
        engine.add_text('Col2', column=2)
        ws = engine.wb.active
        assert ws.cell(row=1, column=2).value == 'Col2'

    def test_add_text_multiline(self, engine: ExcelEngine) -> None:
        engine.add_text('A\nB\nC')
        ws = engine.wb.active
        assert ws.cell(row=1, column=1).value == 'A'
        assert ws.cell(row=2, column=1).value == 'B'
        assert ws.cell(row=3, column=1).value == 'C'

    def test_replace_text_exact(self, engine: ExcelEngine) -> None:
        engine.add_text('Hello World')
        count = engine.replace_text('World', 'Earth')
        assert count >= 1
        assert engine.wb.active.cell(row=1, column=1).value == 'Hello Earth'

    def test_get_metadata(self, engine: ExcelEngine) -> None:
        metadata = engine.get_metadata()
        assert 'author' in metadata
        assert 'title' in metadata

    def test_set_metadata(self, engine: ExcelEngine) -> None:
        engine.set_metadata(author='Test Author', title='Test')
        metadata = engine.get_metadata()
        assert metadata['author'] == 'Test Author'
        assert metadata['title'] == 'Test'

    def test_set_style(self, engine: ExcelEngine) -> None:
        engine.set_style('font=Courier New,size=14')
        assert engine._base_style.font_name == 'Courier New'
        assert engine._base_style.font_size == 14

    def test_add_sheet(self, engine: ExcelEngine) -> None:
        count_before = len(engine.wb.sheetnames)
        engine.add_sheet('NewSheet')
        assert len(engine.wb.sheetnames) == count_before + 1
        assert 'NewSheet' in engine.wb.sheetnames

    def test_delete_sheet(self, engine: ExcelEngine) -> None:
        engine.add_sheet('ToDelete')
        engine.delete_sheet('ToDelete')
        assert 'ToDelete' not in engine.wb.sheetnames

    def test_rename_sheet(self, engine: ExcelEngine) -> None:
        engine.add_sheet('OldName')
        engine.rename_sheet('OldName', 'NewName')
        assert 'NewName' in engine.wb.sheetnames
        assert 'OldName' not in engine.wb.sheetnames

    def test_set_column_width(self, engine: ExcelEngine) -> None:
        engine.set_column_width(2, 25.5)
        ws = engine.wb.active
        assert ws.column_dimensions['B'].width == 25.5

    def test_set_row_height(self, engine: ExcelEngine) -> None:
        engine.set_row_height(3, 30.0)
        ws = engine.wb.active
        assert ws.row_dimensions[3].height == 30.0

    def test_set_formula(self, engine: ExcelEngine) -> None:
        engine.set_formula('A1', '=SUM(B1:B10)')
        ws = engine.wb.active
        assert ws['A1'].value == '=SUM(B1:B10)'

    def test_import_csv(self, engine: ExcelEngine) -> None:
        import csv

        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / 'test.csv'
            with open(csv_path, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Name', 'Age'])
                writer.writerow(['Alice', '30'])
            engine.import_csv(str(csv_path))
            ws = engine.wb.active
            assert ws.cell(row=1, column=1).value == 'Name'
            assert ws.cell(row=2, column=2).value == '30'

    def test_export_csv(self, engine: ExcelEngine) -> None:
        engine.add_text('Name', column=1)
        engine.add_text('Alice', column=1)
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / 'out.csv'
            engine.export_csv(str(out_path))
            assert out_path.exists()
            content = out_path.read_text()
            assert 'Name' in content
            assert 'Alice' in content

    def test_export_json(self, engine: ExcelEngine) -> None:
        engine.add_text('Name', column=1)
        engine.add_text('Alice', column=1)
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / 'out.json'
            engine.export_json(str(out_path))
            assert out_path.exists()
            import json

            data = json.loads(out_path.read_text())
            assert isinstance(data, list)
            assert data[0]['Name'] == 'Alice'

    def test_export_html(self, engine: ExcelEngine) -> None:
        engine.add_text('Header', column=1)
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / 'out.html'
            engine.export_html(str(out_path))
            assert out_path.exists()
            content = out_path.read_text()
            assert '<table>' in content
            assert 'Header' in content

    def test_sort_asc(self, engine: ExcelEngine) -> None:
        engine.add_text('B', column=1)
        engine.add_text('A', column=1)
        engine.add_text('C', column=1)
        engine.sort('A1:A3', 'asc')
        ws = engine.wb.active
        assert ws.cell(row=1, column=1).value == 'A'
        assert ws.cell(row=2, column=1).value == 'B'
        assert ws.cell(row=3, column=1).value == 'C'

    def test_sort_desc(self, engine: ExcelEngine) -> None:
        engine.add_text('A', column=1)
        engine.add_text('B', column=1)
        engine.add_text('C', column=1)
        engine.sort('A1:A3', 'desc')
        ws = engine.wb.active
        assert ws.cell(row=1, column=1).value == 'C'
        assert ws.cell(row=3, column=1).value == 'A'

    def test_add_chart_bar(self, engine: ExcelEngine) -> None:
        ws = engine.wb.active
        engine.add_text('Cats', column=1)
        engine.add_text('10', column=1)
        engine.add_text('Dogs', column=2)
        engine.add_text('20', column=2)
        sheet_name = ws.title
        engine.add_chart('bar', f"'{sheet_name}'!A1:B2")
        assert len(ws._charts) >= 1

    def test_add_chart_unsupported_raises(self, engine: ExcelEngine) -> None:
        with pytest.raises(ValueError, match='Unsupported chart type'):
            engine.add_chart('scatter', 'A1:B2')

    def test_merge_workbooks(self, engine: ExcelEngine) -> None:
        engine.add_sheet('Data')
        engine.add_text('Hello', column=1)
        with tempfile.TemporaryDirectory() as tmpdir:
            p1 = Path(tmpdir) / 'src.xlsx'
            engine.save(str(p1))
            e2 = ExcelEngine()
            e2.create()
            e2.merge_workbooks([str(p1)])
            assert 'Data' in e2.wb.sheetnames

    def test_split_by_sheet(self, engine: ExcelEngine) -> None:
        engine.add_sheet('Sheet1')
        engine.add_sheet('Sheet2')
        with tempfile.TemporaryDirectory() as tmpdir:
            results = engine.split_by_sheet(tmpdir)
            assert len(results) == 3
            for p in results:
                assert p.exists()

    def test_clear_content(self, engine: ExcelEngine) -> None:
        engine.add_text('Data', column=1)
        engine.add_text('More', column=1)
        engine.clear_content()
        ws = engine.wb.active
        assert ws.cell(row=1, column=1).value is None
        assert ws.cell(row=2, column=1).value is None

    def test_set_protection(self, engine: ExcelEngine) -> None:
        engine.set_protection('secret123')
        assert engine.wb.security.workbook_password is not None

    def test_unprotect(self, engine: ExcelEngine) -> None:
        engine.set_protection('secret123')
        engine.unprotect()
        assert engine.wb.security.workbook_password is not None

    def test_add_comment(self, engine: ExcelEngine) -> None:
        engine.add_text('Cell', column=1)
        engine.add_comment('A1', 'A note')
        ws = engine.wb.active
        assert ws['A1'].comment is not None
        assert ws['A1'].comment.text == 'A note'


class TestExcelImportJson:
    def test_import_json_dicts(self, tmp_path: Path) -> None:
        data = [{'name': 'Alice', 'age': 30}, {'name': 'Bob', 'age': 25}]
        json_path = tmp_path / 'data.json'
        json_path.write_text(__import__('json').dumps(data), encoding='utf-8')
        e = ExcelEngine()
        e.create()
        e.import_json(str(json_path))
        rows = [list(r) for r in e.wb.active.iter_rows(values_only=True)]
        assert rows == [['name', 'age'], ['Alice', 30], ['Bob', 25]]

    def test_import_json_arrays(self, tmp_path: Path) -> None:
        data = [[1, 2], [3, 4]]
        json_path = tmp_path / 'data.json'
        json_path.write_text(__import__('json').dumps(data), encoding='utf-8')
        e = ExcelEngine()
        e.create()
        e.import_json(str(json_path))
        rows = [list(r) for r in e.wb.active.iter_rows(values_only=True)]
        assert rows == [[1, 2], [3, 4]]

    def test_import_json_empty_raises(self, tmp_path: Path) -> None:
        json_path = tmp_path / 'data.json'
        json_path.write_text('[]', encoding='utf-8')
        e = ExcelEngine()
        e.create()
        with pytest.raises(ValueError, match='non-empty'):
            e.import_json(str(json_path))


class TestExcelExtract:
    def test_extract_text(self) -> None:
        e = ExcelEngine()
        e.create()
        ws = e.wb.active
        ws['A1'] = 'x'
        ws['B1'] = 2
        text = e.extract_text()
        assert '[Sheet] x | 2' in text

    def test_extract_tables(self) -> None:
        e = ExcelEngine()
        e.create()
        e.wb.active['A1'] = 'a'
        tables = e.extract_tables()
        assert tables == [[['a']]]

    def test_extract_structure(self) -> None:
        e = ExcelEngine()
        e.create()
        struct = e.extract_structure()
        assert struct['sheets'] == ['Sheet']
