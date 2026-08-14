"""Extended unit tests for the template engine (loops, Excel, YAML, edge cases)."""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

import pytest

from tianshang_scribe.core.document import DocumentType, create_document
from tianshang_scribe.rendering.template import TemplateEngine, defaultdict_from_yaml


def _tmp_json(data: dict, name: str = 'data.json') -> Path:
    tmpdir = tempfile.mkdtemp()
    path = Path(tmpdir) / name
    path.write_text(json.dumps(data))
    return path


class TestLoadFormats:
    def test_load_yaml(self) -> None:
        tmpdir = tempfile.mkdtemp()
        path = Path(tmpdir) / 'data.yaml'
        path.write_text('name: Alice\nitems:\n  - a\n  - b')
        engine = TemplateEngine(str(path))
        assert engine._data['name'] == 'Alice'
        assert engine._data['items'] == ['a', 'b']

    def test_load_yaml_import_error(self, monkeypatch) -> None:
        tmpdir = tempfile.mkdtemp()
        path = Path(tmpdir) / 'data.yaml'
        path.write_text('x: 1')
        monkeypatch.setattr('sys.modules', {**dict(__import__('sys').modules), 'yaml': None})
        import builtins

        real_import = builtins.__import__

        def fake_import(name: str, *a: object, **k: object) -> object:
            if name == 'yaml':
                raise ImportError(name)
            return real_import(name, *a, **k)

        monkeypatch.setattr(builtins, '__import__', fake_import)
        with pytest.raises(ImportError, match='PyYAML is required'):
            TemplateEngine(str(path))


class TestYamlHelper:
    def test_defaultdict_from_yaml_nested(self) -> None:
        data = {'a': {'b': [1, {'c': 2}]}, 'd': 'x'}
        out = defaultdict_from_yaml(data)
        assert out == {'a': {'b': [1, {'c': 2}]}, 'd': 'x'}
        assert isinstance(out, dict)
        assert isinstance(out['a'], dict)

    def test_scalar_passthrough(self) -> None:
        assert defaultdict_from_yaml('str') == 'str'
        assert defaultdict_from_yaml(3) == 3
        assert defaultdict_from_yaml(None) is None


class TestWordEachLoop:
    def test_each_loop_expands(self) -> None:
        doc = create_document(DocumentType.WORD)
        doc.add_text('{{#each items}}')
        doc.add_text('Item: {{this}}')
        doc.add_text('{{/each}}')

        data = _tmp_json({'items': ['A', 'B', 'C']})
        engine = TemplateEngine(str(data))
        count = engine.fill(doc)
        assert count >= 3
        text = '\n'.join(p.text.strip() for p in doc.doc.paragraphs if p.text.strip())
        assert 'Item: A' in text
        assert 'Item: B' in text
        assert 'Item: C' in text

    def test_each_loop_dict_items(self) -> None:
        doc = create_document(DocumentType.WORD)
        doc.add_text('{{#each people}}')
        doc.add_text('Name: {{name}}')
        doc.add_text('{{/each}}')

        data = _tmp_json({'people': [{'name': 'Alice'}, {'name': 'Bob'}]})
        engine = TemplateEngine(str(data))
        engine.fill(doc)
        text = '\n'.join(p.text.strip() for p in doc.doc.paragraphs if p.text.strip())
        assert 'Name: Alice' in text
        assert 'Name: Bob' in text

    def test_each_loop_missing_key_returns_zero(self) -> None:
        doc = create_document(DocumentType.WORD)
        doc.add_text('{{#each missing}}')
        doc.add_text('X')
        doc.add_text('{{/each}}')
        engine = TemplateEngine(str(_tmp_json({'other': 1})))
        assert engine._process_word_loop(doc.doc, 0, 'missing', {'other': 1}) == 0

    def test_each_loop_non_list_value_returns_zero(self) -> None:
        doc = create_document(DocumentType.WORD)
        doc.add_text('{{#each items}}')
        doc.add_text('X')
        doc.add_text('{{/each}}')
        engine = TemplateEngine(str(_tmp_json({'items': ['A']})))
        assert engine._process_word_loop(doc.doc, 0, 'items', {'items': 'notalist'}) == 0

    def test_each_loop_no_end_tag(self) -> None:
        doc = create_document(DocumentType.WORD)
        doc.add_text('{{#each items}}')
        doc.add_text('X')
        doc.add_text('no end tag here')
        engine = TemplateEngine(str(_tmp_json({'items': ['A']})))
        flat = engine._flatten_data(engine._data)
        assert engine._process_word_loop(doc.doc, 0, 'items', flat) == 0


class TestWordIfEdgeCases:
    def test_if_no_end_tag_returns_zero(self) -> None:
        doc = create_document(DocumentType.WORD)
        doc.add_text('{{#if show}}')
        doc.add_text('No end tag')
        engine = TemplateEngine(str(_tmp_json({'show': True})))
        assert engine._process_word_if(doc.doc, 0, 'show', None, {'show': True}, negate=False) == 0

    def test_unless_true_hides_content(self) -> None:
        doc = create_document(DocumentType.WORD)
        doc.add_text('{{#unless paid}}Hidden{{/unless}}')
        data = _tmp_json({'paid': True})
        engine = TemplateEngine(str(data))
        engine.fill(doc)
        text = '\n'.join(p.text.strip() for p in doc.doc.paragraphs if p.text.strip())
        assert 'Hidden' not in text

    def test_if_empty_paragraph_cleared(self) -> None:
        doc = create_document(DocumentType.WORD)
        doc.add_text('{{#if show}}')
        doc.add_text('')
        doc.add_text('{{/if}}')
        data = _tmp_json({'show': True})
        engine = TemplateEngine(str(data))
        engine.fill(doc)
        assert not any(p.text.strip() for p in doc.doc.paragraphs)


class TestExcelFill:
    def test_excel_simple_placeholder(self) -> None:
        wb = create_document(DocumentType.EXCEL)
        ws = wb.wb.active
        ws.cell(row=1, column=1, value='Hello {{name}}')
        data = _tmp_json({'name': 'Alice'})
        engine = TemplateEngine(str(data))
        count = engine.fill(wb)
        assert count >= 1
        assert ws.cell(row=1, column=1).value == 'Hello Alice'

    def test_excel_each_loop(self) -> None:
        wb = create_document(DocumentType.EXCEL)
        ws = wb.wb.active
        ws.cell(row=1, column=1, value='{{#each items}}')
        ws.cell(row=2, column=1, value='Item: {{this}}')
        data = _tmp_json({'items': ['X', 'Y']})
        engine = TemplateEngine(str(data))
        engine.fill(wb)
        values = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
        assert 'Item: X' in values
        assert 'Item: Y' in values

    def test_excel_each_dict_loop(self) -> None:
        wb = create_document(DocumentType.EXCEL)
        ws = wb.wb.active
        ws.cell(row=1, column=1, value='{{#each people}}')
        ws.cell(row=2, column=1, value='{{name}}')
        data = _tmp_json({'people': [{'name': 'A'}, {'name': 'B'}]})
        engine = TemplateEngine(str(data))
        engine.fill(wb)
        values = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
        assert 'A' in values
        assert 'B' in values

    def test_excel_no_loop_keys_falls_back(self) -> None:
        wb = create_document(DocumentType.EXCEL)
        ws = wb.wb.active
        ws.cell(row=1, column=1, value='hi {{who}}')
        data = _tmp_json({'who': 'world', 'items': []})
        engine = TemplateEngine(str(data))
        count = engine.fill(wb)
        assert count >= 1
        assert ws.cell(row=1, column=1).value == 'hi world'


class TestSimpleFill:
    class _FakeEngine:
        def __init__(self) -> None:
            self.calls: list[tuple[str, str]] = []

        def replace_text(self, placeholder: str, value: str) -> int:
            self.calls.append((placeholder, value))
            return 1

    def test_fill_simple_replaces_non_lists(self) -> None:
        engine = TemplateEngine(str(_tmp_json({'a': '1', 'b': 2, 'items': ['x']})))
        fake = self._FakeEngine()
        count = engine.fill(fake)
        assert count == 2
        assert ('{{a}}', '1') in fake.calls
        assert ('{{b}}', '2') in fake.calls

    def test_fill_simple_skips_lists(self) -> None:
        engine = TemplateEngine(str(_tmp_json({'items': ['x'], 'n': 1})))
        fake = self._FakeEngine()
        engine._fill_simple(fake, {'items': ['x'], 'n': 1})
        assert ('{{items}}', 'x') not in fake.calls
