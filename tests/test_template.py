from __future__ import annotations

import json
import tempfile
from pathlib import Path

from src.core.document import DocumentType, create_document
from src.rendering.template import TemplateEngine


class TestTemplateEngine:
    def test_load_json(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            data_path = Path(tmpdir) / 'data.json'
            data_path.write_text(json.dumps({'name': 'Alice', 'age': 30}))
            engine = TemplateEngine(str(data_path))
            assert engine._data == {'name': 'Alice', 'age': 30}

    def test_load_csv(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            data_path = Path(tmpdir) / 'data.csv'
            data_path.write_text('name,age\nAlice,30\nBob,25')
            engine = TemplateEngine(str(data_path))
            assert engine._data['count'] == 2
            assert len(engine._data['rows']) == 2

    def test_flatten_nested(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            data_path = Path(tmpdir) / 'nested.json'
            data_path.write_text(json.dumps({
                'user': {'name': 'Alice', 'details': {'city': 'NYC'}},
                'count': 5
            }))
            engine = TemplateEngine(str(data_path))
            flat = engine._flatten_data(engine._data)
            assert flat['user.name'] == 'Alice'
            assert flat['user.details.city'] == 'NYC'
            assert flat['count'] == 5

    def test_fill_word_document(self) -> None:
        doc = create_document(DocumentType.WORD)
        doc.add_text('Hello {{name}}, you are {{age}} years old.')

        with tempfile.TemporaryDirectory() as tmpdir:
            data_path = Path(tmpdir) / 'data.json'
            data_path.write_text(json.dumps({'name': 'Alice', 'age': '30'}))

            engine = TemplateEngine(str(data_path))
            count = engine.fill(doc)
            assert count >= 2

    def test_unsupported_format_raises(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            data_path = Path(tmpdir) / 'data.xml'
            data_path.write_text('<root/>')
            with __import__('pytest').raises(ValueError, match='Unsupported'):
                TemplateEngine(str(data_path))
