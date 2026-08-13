"""Direct-drive tests for the 7 MCP tool implementations.

The tools are plain callables whose parameters are validated by the MCP SDK
at call time; here we invoke them directly with pydantic models so the full
response schemas (success/error/structured content) are exercised without a
running server.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from src.core.excel_engine import ExcelEngine
from src.core.word_engine import WordEngine
from src.mcp.errors import (
    McpErrorCode,
    _make_content,
    _set_notify_writer,
    error_response,
    send_progress,
    success_response,
)
from src.mcp.schemas import ContentBlock, EditOperation, ToolOptions
from src.mcp.tools.compare import compare_documents
from src.mcp.tools.convert import convert_document, extract_document_data
from src.mcp.tools.create import create_office_document
from src.mcp.tools.edit import edit_office_document
from src.mcp.tools.template import fill_template
from src.mcp.tools.validate import validate_template


@pytest.fixture
def docx_with_text(tmp_path: Path) -> Path:
    e = WordEngine()
    e.create()
    e.add_text('hello world')
    p = tmp_path / 'src.docx'
    e.save(str(p))
    return p


@pytest.fixture
def xlsx_with_data(tmp_path: Path) -> Path:
    e = ExcelEngine()
    e.create()
    e.add_text('Alice', column=1)
    e.add_text('30', column=1)
    p = tmp_path / 'src.xlsx'
    e.save(str(p))
    return p


class TestCreateDocument:
    def test_create_docx(self, tmp_path: Path) -> None:
        out = tmp_path / 'out.docx'
        result = create_office_document(
            'docx',
            [
                ContentBlock(type='paragraph', text='Hello'),
                ContentBlock(type='heading', text='Title', level=1),
                ContentBlock(type='formula', text=r'\frac{a}{b}'),
                ContentBlock(type='table', rows=[['A', 'B'], ['1', '2']]),
            ],
            output_path=str(out),
            metadata={'author': 'Tester'},
        )
        assert result['success'] is True
        data = result['data']
        assert Path(data['output_path']).exists()
        assert data['format'] == 'docx'
        assert data['content_items'] == 4

    def test_create_xlsx(self, tmp_path: Path) -> None:
        out = tmp_path / 'out.xlsx'
        result = create_office_document(
            'xlsx',
            [ContentBlock(type='paragraph', text='cell0')],
            output_path=str(out),
        )
        assert result['success'] is True
        assert 'sheets' in result['data']

    def test_create_pptx(self, tmp_path: Path) -> None:
        out = tmp_path / 'out.pptx'
        result = create_office_document(
            'pptx',
            [ContentBlock(type='paragraph', text='Slide text')],
            output_path=str(out),
        )
        assert result['success'] is True
        assert 'slides' in result['data']

    def test_unsupported_format(self, tmp_path: Path) -> None:
        result = create_office_document(
            'txt',
            [ContentBlock(type='paragraph', text='x')],
            output_path=str(tmp_path / 'o.txt'),
        )
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.UNSUPPORTED_FORMAT

    def test_dry_run(self, tmp_path: Path) -> None:
        out = tmp_path / 'out.docx'
        result = create_office_document(
            'docx',
            [ContentBlock(type='paragraph', text='Hi')],
            output_path=str(out),
            options=ToolOptions(dry_run=True),
        )
        assert result['success'] is True
        assert result['data']['dry_run'] is True
        assert not out.exists()

    def test_default_output_path_in_temp(self) -> None:
        import tempfile

        result = create_office_document(
            'docx',
            [ContentBlock(type='paragraph', text='Hi')],
        )
        assert result['success'] is True
        assert result['data']['output_path'].startswith(tempfile.gettempdir())

    def test_template_data_fill(self, tmp_path: Path) -> None:
        out = tmp_path / 'out.docx'
        result = create_office_document(
            'docx',
            [ContentBlock(type='paragraph', text='Hi {{name}}')],
            output_path=str(out),
            template_data={'name': 'Bob'},
        )
        assert result['success'] is True
        reopened = WordEngine()
        reopened.open(str(out))
        assert 'Bob' in reopened.extract_text()


class TestEditDocument:
    def test_replace_operation(self, tmp_path: Path, docx_with_text: Path) -> None:
        out = tmp_path / 'edited.docx'
        result = edit_office_document(
            str(docx_with_text),
            [EditOperation(action='replace', old_text='world', new_text='earth')],
            output_path=str(out),
        )
        assert result['success'] is True
        assert result['data']['total_changes'] >= 1
        e = WordEngine()
        e.open(str(out))
        assert 'earth' in e.extract_text()

    def test_add_and_style_operations(self, tmp_path: Path, docx_with_text: Path) -> None:
        out = tmp_path / 'edited.docx'
        result = edit_office_document(
            str(docx_with_text),
            [
                EditOperation(action='add', text=' appended'),
                EditOperation(action='style', style='font=Arial', apply_all=True),
            ],
            output_path=str(out),
        )
        assert result['success'] is True
        assert result['data']['total_changes'] == 2

    def test_clear_operation(self, tmp_path: Path, docx_with_text: Path) -> None:
        out = tmp_path / 'cleared.docx'
        result = edit_office_document(
            str(docx_with_text),
            [EditOperation(action='clear')],
            output_path=str(out),
        )
        assert result['success'] is True
        e = WordEngine()
        e.open(str(out))
        assert e.extract_text().strip() == ''

    def test_missing_document(self, tmp_path: Path) -> None:
        result = edit_office_document(
            str(tmp_path / 'nope.docx'),
            [EditOperation(action='replace', old_text='a', new_text='b')],
        )
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.DOCUMENT_NOT_FOUND

    def test_dry_run(self, tmp_path: Path, docx_with_text: Path) -> None:
        result = edit_office_document(
            str(docx_with_text),
            [EditOperation(action='replace', old_text='world', new_text='earth')],
            options=ToolOptions(dry_run=True),
        )
        assert result['success'] is True
        assert result['data']['dry_run'] is True

    def test_backup_option(self, tmp_path: Path, docx_with_text: Path) -> None:
        result = edit_office_document(
            str(docx_with_text),
            [EditOperation(action='replace', old_text='world', new_text='earth')],
            options=ToolOptions(backup=True),
        )
        assert result['success'] is True
        assert (tmp_path / 'src.docx.bak').exists()


class TestConvertDocument:
    def test_excel_to_csv(self, tmp_path: Path, xlsx_with_data: Path) -> None:
        out = tmp_path / 'out.csv'
        result = convert_document(str(xlsx_with_data), 'csv', output_path=str(out))
        assert result['success'] is True
        assert out.exists()

    def test_excel_to_json(self, tmp_path: Path, xlsx_with_data: Path) -> None:
        out = tmp_path / 'out.json'
        result = convert_document(str(xlsx_with_data), 'json', output_path=str(out))
        assert result['success'] is True
        assert out.exists()

    def test_excel_to_html(self, tmp_path: Path, xlsx_with_data: Path) -> None:
        out = tmp_path / 'out.html'
        result = convert_document(str(xlsx_with_data), 'html', output_path=str(out))
        assert result['success'] is True
        assert out.exists()

    def test_word_to_markdown(self, tmp_path: Path, docx_with_text: Path) -> None:
        out = tmp_path / 'out.md'
        result = convert_document(str(docx_with_text), 'md', output_path=str(out))
        assert result['success'] is True
        assert out.exists()

    def test_missing_document(self) -> None:
        result = convert_document('/nonexistent/x.docx', 'csv')
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.DOCUMENT_NOT_FOUND

    def test_unsupported_format(self, tmp_path: Path, docx_with_text: Path) -> None:
        result = convert_document(str(docx_with_text), 'pdfx')
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.UNSUPPORTED_FORMAT

    def test_default_output_path(self, tmp_path: Path, xlsx_with_data: Path, monkeypatch) -> None:
        monkeypatch.chdir(tmp_path)
        result = convert_document(str(xlsx_with_data), 'csv')
        assert result['success'] is True
        assert result['data']['output_path'].endswith('src.csv')
        assert (tmp_path / 'src.csv').exists()

    def test_dry_run(self, tmp_path: Path, docx_with_text: Path) -> None:
        result = convert_document(
            str(docx_with_text),
            'md',
            options=ToolOptions(dry_run=True),
        )
        assert result['success'] is True
        assert result['data']['dry_run'] is True


class TestExtractDocumentData:
    def test_metadata(self, docx_with_text: Path) -> None:
        result = extract_document_data(str(docx_with_text), 'metadata')
        assert result['success'] is True
        assert 'author' in result['data']['metadata']

    def test_text(self, docx_with_text: Path) -> None:
        result = extract_document_data(str(docx_with_text), 'text')
        assert result['success'] is True
        assert 'hello world' in result['data']['text']

    def test_text_excel(self, xlsx_with_data: Path) -> None:
        result = extract_document_data(str(xlsx_with_data), 'text')
        assert result['success'] is True
        assert 'Alice' in result['data']['text']

    def test_structure(self, docx_with_text: Path) -> None:
        result = extract_document_data(str(docx_with_text), 'structure')
        assert result['success'] is True
        assert result['data']['paragraphs'] >= 1

    def test_missing_document(self) -> None:
        result = extract_document_data('/nonexistent/x.docx', 'metadata')
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.DOCUMENT_NOT_FOUND

    def test_invalid_mode(self, docx_with_text: Path) -> None:
        result = extract_document_data(str(docx_with_text), 'bogus')
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.INVALID_PARAMETER


class TestFillTemplate:
    def test_fill_docx(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('Hi {{name}}')
        tpl = tmp_path / 'tpl.docx'
        e.save(str(tpl))
        out = tmp_path / 'filled.docx'

        result = fill_template(
            str(tpl),
            {'name': 'Alice'},
            output_path=str(out),
        )
        assert result['success'] is True
        assert result['data']['placeholders_filled'] >= 1
        reopened = WordEngine()
        reopened.open(str(out))
        assert 'Alice' in reopened.extract_text()

    def test_missing_template(self) -> None:
        result = fill_template('/nonexistent/tpl.docx', {'name': 'x'})
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.DOCUMENT_NOT_FOUND

    def test_dry_run(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('Hi {{name}}')
        tpl = tmp_path / 'tpl.docx'
        e.save(str(tpl))
        result = fill_template(str(tpl), {'name': 'x'}, options=ToolOptions(dry_run=True))
        assert result['success'] is True
        assert result['data']['dry_run'] is True


class TestValidateTemplate:
    def test_all_keys_present(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('Hi {{name}} from {{city}}')
        tpl = tmp_path / 'tpl.docx'
        e.save(str(tpl))
        result = validate_template(str(tpl), {'name': 'a', 'city': 'b'})
        assert result['success'] is True
        assert result['data']['valid'] is True

    def test_missing_key(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('Hi {{name}}')
        tpl = tmp_path / 'tpl.docx'
        e.save(str(tpl))
        result = validate_template(str(tpl), {})
        assert result['success'] is True
        assert result['data']['valid'] is False
        assert any('name' in m for m in result['data']['missing'])

    def test_missing_template(self) -> None:
        result = validate_template('/nonexistent/tpl.docx', {'name': 'x'})
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.DOCUMENT_NOT_FOUND

    def test_nested_data_flattening(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('Hi {{user.name}}')
        tpl = tmp_path / 'tpl.docx'
        e.save(str(tpl))
        result = validate_template(str(tpl), {'user': {'name': 'a'}})
        assert result['success'] is True
        assert result['data']['valid'] is True


class TestCompareDocuments:
    def test_identical(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('same')
        a = tmp_path / 'a.docx'
        b = tmp_path / 'b.docx'
        e.save(str(a))
        e.save(str(b))
        result = compare_documents(str(a), str(b))
        assert result['success'] is True
        assert result['data']['identical'] is True

    def test_differences(self, tmp_path: Path, docx_with_text: Path) -> None:
        e = WordEngine()
        e.create()
        other = tmp_path / 'other.docx'
        e.save(str(other))
        result = compare_documents(str(docx_with_text), str(other))
        assert result['success'] is True
        assert result['data']['identical'] is False
        assert result['data']['removed'] or result['data']['changed']

    def test_missing_document(self, docx_with_text: Path) -> None:
        result = compare_documents(str(docx_with_text), '/nonexistent/b.docx')
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.DOCUMENT_NOT_FOUND

    def test_non_word_unsupported(self, tmp_path: Path, xlsx_with_data: Path) -> None:
        e = WordEngine()
        e.create()
        w = tmp_path / 'w.docx'
        e.save(str(w))
        result = compare_documents(str(w), str(xlsx_with_data))
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.UNSUPPORTED_FORMAT


class TestErrorsHelpers:
    def test_error_response(self) -> None:
        resp = error_response(McpErrorCode.DOCUMENT_NOT_FOUND, 'detail')
        assert resp['success'] is False
        assert resp['error_code'] == McpErrorCode.DOCUMENT_NOT_FOUND
        assert 'detail' in resp['error_message']
        assert resp['retryable'] is True

    def test_error_response_unknown_code(self) -> None:
        resp = error_response(12345)
        assert resp['success'] is False
        assert 'unexpected error' in resp['error_message'].lower()
        assert resp['suggested_fix']

    def test_success_response(self) -> None:
        resp = success_response({'k': 1}, content=[{'type': 'text', 'text': 'ok'}])
        assert resp['success'] is True
        assert resp['data'] == {'k': 1}
        assert resp['content']
        assert success_response() == {'success': True}

    def test_make_content_with_existing_file(self, tmp_path: Path) -> None:
        f = tmp_path / 'out.docx'
        f.write_bytes(b'x')
        content = _make_content(str(f), 'done')
        assert content[0] == {'type': 'text', 'text': 'done'}
        resource = content[1]['resource']
        assert resource['mimeType'].startswith('application/')
        assert resource['title'] == 'out.docx'

    def test_make_content_no_file(self, tmp_path: Path) -> None:
        content = _make_content(str(tmp_path / 'missing.docx'), 'done')
        assert len(content) == 1

    def test_send_progress_with_writer(self) -> None:
        written: list[str] = []
        _set_notify_writer(written.append)
        try:
            import json

            send_progress(1, 10, 'step')
            assert len(written) == 1
            payload = json.loads(written[0])
            assert payload['method'] == 'notifications/progress'
            assert payload['params']['progress'] == 1
        finally:
            _set_notify_writer(None)

    def test_send_progress_no_writer(self) -> None:
        _set_notify_writer(None)
        send_progress(1, 10)  # must be a no-op
