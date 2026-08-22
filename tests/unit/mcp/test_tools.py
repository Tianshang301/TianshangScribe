"""Direct-drive tests for the 7 MCP tool implementations.

The tools are plain callables whose parameters are validated by the MCP SDK
at call time; here we invoke them directly with pydantic models so the full
response schemas (success/error/structured content) are exercised without a
running server.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tianshang_scribe.core.excel_engine import ExcelEngine
from tianshang_scribe.core.word_engine import WordEngine
from tianshang_scribe.mcp.errors import (
    McpErrorCode,
    _make_content,
    _set_notify_writer,
    error_response,
    send_progress,
    success_response,
)
from tianshang_scribe.mcp.schemas import ContentBlock, EditOperation, ToolOptions
from tianshang_scribe.mcp.tools.compare import compare_documents
from tianshang_scribe.mcp.tools.convert import convert_document, extract_document_data
from tianshang_scribe.mcp.tools.create import create_office_document
from tianshang_scribe.mcp.tools.edit import edit_office_document
from tianshang_scribe.mcp.tools.template import fill_template
from tianshang_scribe.mcp.tools.validate import validate_template


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


class TestCompareSnapshots:
    def test_snapshot_and_list(self, tmp_path: Path, docx_with_text: Path) -> None:
        store = tmp_path / 'store'
        result = compare_documents(
            str(docx_with_text),
            '',
            ToolOptions(action='snapshot', snapshot_dir=str(store)),
        )
        assert result['success'] is True
        snapshot_id = result['data']['snapshot_id']
        assert result['data']['paragraph_count'] == 1

        listed = compare_documents(
            str(docx_with_text),
            '',
            ToolOptions(action='list_snapshots', snapshot_dir=str(store)),
        )
        assert listed['success'] is True
        assert snapshot_id in [s['snapshot_id'] for s in listed['data']['snapshots']]

    def test_snapshot_missing_document(self, tmp_path: Path) -> None:
        result = compare_documents(
            str(tmp_path / 'nope.docx'),
            '',
            ToolOptions(action='snapshot', snapshot_dir=str(tmp_path / 'store')),
        )
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.DOCUMENT_NOT_FOUND

    def test_snapshot_restore_roundtrip(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('line one')
        e.add_text('line two')
        src = tmp_path / 'src.docx'
        e.save(str(src))

        store = tmp_path / 'store'
        snap = compare_documents(
            str(src), '', ToolOptions(action='snapshot', snapshot_dir=str(store))
        )
        snapshot_id = snap['data']['snapshot_id']

        out = tmp_path / 'restored.docx'
        result = compare_documents(
            str(src),
            str(out),
            ToolOptions(action='restore', snapshot_dir=str(store), snapshot_id=snapshot_id),
        )
        assert result['success'] is True
        assert result['data']['paragraph_count'] == 2
        restored = WordEngine()
        restored.open(str(out))
        texts = [p.text for p in restored.doc.paragraphs]
        assert texts == ['line one', 'line two']

    def test_restore_missing_snapshot(self, tmp_path: Path, docx_with_text: Path) -> None:
        store = tmp_path / 'store'
        result = compare_documents(
            str(docx_with_text),
            str(tmp_path / 'out.docx'),
            ToolOptions(action='restore', snapshot_dir=str(store), snapshot_id='deadbeef'),
        )
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.DOCUMENT_NOT_FOUND

    def test_restore_without_snapshot_id(self, tmp_path: Path, docx_with_text: Path) -> None:
        result = compare_documents(
            str(docx_with_text),
            str(tmp_path / 'out.docx'),
            ToolOptions(action='restore'),
        )
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.INVALID_PARAMETER

    def test_list_empty_store(self, tmp_path: Path, docx_with_text: Path) -> None:
        result = compare_documents(
            str(docx_with_text),
            '',
            ToolOptions(action='list_snapshots', snapshot_dir=str(tmp_path / 'store')),
        )
        assert result['success'] is True
        assert result['data']['snapshots'] == []


class TestCreateCrossFormatFallbacks:
    """Data-driven: content blocks must degrade gracefully on non-Word engines."""

    @pytest.mark.parametrize('fmt,ext', [('xlsx', '.xlsx'), ('pptx', '.pptx')])
    def test_heading_falls_back_to_latex(self, fmt: str, ext: str, tmp_path: Path) -> None:
        out = tmp_path / f'out{ext}'
        result = create_office_document(
            fmt,
            [ContentBlock(type='heading', text='Title', level=1)],
            output_path=str(out),
        )
        assert result['success'] is True
        assert Path(result['data']['output_path']).exists()

    @pytest.mark.parametrize('fmt,ext', [('xlsx', '.xlsx'), ('pptx', '.pptx')])
    def test_formula_falls_back_to_latex(self, fmt: str, ext: str, tmp_path: Path) -> None:
        out = tmp_path / f'out{ext}'
        result = create_office_document(
            fmt,
            [ContentBlock(type='formula', text=r'\frac{a}{b}')],
            output_path=str(out),
        )
        assert result['success'] is True

    @pytest.mark.parametrize('fmt,ext', [('xlsx', '.xlsx'), ('pptx', '.pptx')])
    def test_page_break_falls_back(self, fmt: str, ext: str, tmp_path: Path) -> None:
        out = tmp_path / f'out{ext}'
        result = create_office_document(
            fmt,
            [ContentBlock(type='page_break')],
            output_path=str(out),
        )
        assert result['success'] is True

    def test_table_plain_text_fallback(self, tmp_path: Path) -> None:
        out = tmp_path / 'out.xlsx'
        result = create_office_document(
            'xlsx',
            [ContentBlock(type='table', rows=[['A', 'B'], ['1', '2']])],
            output_path=str(out),
        )
        assert result['success'] is True

    def test_latex_detection_word(self, tmp_path: Path) -> None:
        out = tmp_path / 'out.docx'
        result = create_office_document(
            'docx',
            [ContentBlock(type='paragraph', text=r'\bfseries{Bold} and $x^2$')],
            output_path=str(out),
        )
        assert result['success'] is True

    def test_per_block_style(self, tmp_path: Path) -> None:
        out = tmp_path / 'out.docx'
        result = create_office_document(
            'docx',
            [ContentBlock(type='paragraph', text='x', style='font=Arial,size=14')],
            output_path=str(out),
        )
        assert result['success'] is True

    def test_backup_existing_file(self, tmp_path: Path) -> None:
        out = tmp_path / 'out.docx'
        out.write_bytes(b'placeholder')
        result = create_office_document(
            'docx',
            [ContentBlock(type='paragraph', text='hi')],
            output_path=str(out),
            options=ToolOptions(backup=True),
        )
        assert result['success'] is True
        assert (tmp_path / 'out.docx.bak').exists()

    def test_invalid_block_type_ignored(self, tmp_path: Path) -> None:
        from pydantic import ValidationError

        with pytest.raises(ValidationError):
            ContentBlock(type='bogus', text='x')


class TestEditCrossFormat:
    def test_delete_operation(self, tmp_path: Path, docx_with_text: Path) -> None:
        out = tmp_path / 'edited.docx'
        result = edit_office_document(
            str(docx_with_text),
            [EditOperation(action='delete', target='world')],
            output_path=str(out),
        )
        assert result['success'] is True
        assert result['data']['total_changes'] >= 1

    def test_modify_operation(self, tmp_path: Path, docx_with_text: Path) -> None:
        out = tmp_path / 'edited.docx'
        result = edit_office_document(
            str(docx_with_text),
            [EditOperation(action='modify', old_text='hello', new_text='hi')],
            output_path=str(out),
        )
        assert result['success'] is True

    def test_style_without_apply_all(self, tmp_path: Path, docx_with_text: Path) -> None:
        out = tmp_path / 'edited.docx'
        result = edit_office_document(
            str(docx_with_text),
            [EditOperation(action='style', style='font=Arial', apply_all=False)],
            output_path=str(out),
        )
        assert result['success'] is True

    def test_empty_style_string(self, tmp_path: Path, docx_with_text: Path) -> None:
        out = tmp_path / 'edited.docx'
        result = edit_office_document(
            str(docx_with_text),
            [EditOperation(action='style', style='', apply_all=True)],
            output_path=str(out),
        )
        assert result['success'] is True
        assert result['data']['total_changes'] == 1


class TestExtractExcelStructure:
    def test_structure_excel(self, xlsx_with_data: Path) -> None:
        result = extract_document_data(str(xlsx_with_data), 'structure')
        assert result['success'] is True
        assert result['data']['sheets'] == ['Sheet']
        assert result['data']['sheet_count'] == 1


class TestConvertPdfFallback:
    def test_pdf_engine_error_path(self, tmp_path: Path, docx_with_text: Path, monkeypatch) -> None:
        def _raise(*a, **k):
            raise NotImplementedError('to_pdf unavailable')

        monkeypatch.setattr(WordEngine, 'to_pdf', _raise)
        result = convert_document(str(docx_with_text), 'pdf', output_path=str(tmp_path / 'o.pdf'))
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.CONVERSION_FAILED


class TestValidateTemplateStructured:
    def test_loop_missing_key(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('{{#each items}}{{this}}{{/each}}')
        tpl = tmp_path / 'tpl.docx'
        e.save(str(tpl))
        result = validate_template(str(tpl), {'name': 'x'})
        assert result['success'] is True
        assert result['data']['valid'] is False
        assert any('each items' in m for m in result['data']['missing'])

    def test_condition_keys(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('{{#if flag}}yes{{/if}} {{#unless off}}on{{/unless}}')
        tpl = tmp_path / 'tpl.docx'
        e.save(str(tpl))
        result = validate_template(str(tpl), {'flag': True, 'off': False})
        assert result['success'] is True
        assert result['data']['conditions_detected'] == ['flag', 'off']

    def test_nested_loop_warning(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('{{#each users}}{{user.name}}{{/each}}')
        tpl = tmp_path / 'tpl.docx'
        e.save(str(tpl))
        result = validate_template(str(tpl), {'users': [{'name': 'a'}]})
        assert result['success'] is True
        assert result['data']['loops_detected'] == ['users']

    def test_excel_template(self, tmp_path: Path) -> None:
        e = ExcelEngine()
        e.create()
        e.add_text('Hi {{name}}', column=1)
        tpl = tmp_path / 'tpl.xlsx'
        e.save(str(tpl))
        result = validate_template(str(tpl), {'name': 'a'})
        assert result['success'] is True
        assert result['data']['valid'] is True

    def test_pptx_unsupported(self, tmp_path: Path) -> None:
        from tianshang_scribe.core.ppt_engine import PptEngine

        e = PptEngine()
        e.create()
        tpl = tmp_path / 'tpl.pptx'
        e.save(str(tpl))
        result = validate_template(str(tpl), {'name': 'a'})
        assert result['success'] is False
        assert result['error_code'] == McpErrorCode.UNSUPPORTED_FORMAT


class TestFillTemplateBackup:
    def test_backup_option(self, tmp_path: Path) -> None:
        e = WordEngine()
        e.create()
        e.add_text('Hi {{name}}')
        tpl = tmp_path / 'tpl.docx'
        e.save(str(tpl))
        result = fill_template(
            str(tpl),
            {'name': 'x'},
            options=ToolOptions(backup=True),
        )
        assert result['success'] is True
        assert (tmp_path / 'tpl.docx.bak').exists()


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


class TestSchemasMcpP0:
    def test_content_block_new_fields_roundtrip(self) -> None:
        b = ContentBlock(type='table', rows=[['A', 'B'], ['1', '2']], slide_index=0, chart_type='bar', number_format='A1:A10=0.00%')
        d = b.model_dump()
        assert d['slide_index'] == 0
        assert d['chart_type'] == 'bar'
        assert d['number_format'] == 'A1:A10=0.00%'

    def test_edit_operation_new_actions_accepted(self) -> None:
        op = EditOperation(action='set_formula', cell='A1', formula='=SUM(B1:B10)')
        d = op.model_dump()
        assert d['action'] == 'set_formula'
        assert d['cell'] == 'A1'
        assert d['formula'] == '=SUM(B1:B10)'


class TestCreateMcpP0:
    def test_ppt_table_no_longer_crashes(self, tmp_path: Path) -> None:
        out = tmp_path / 'deck.pptx'
        result = create_office_document('pptx', [ContentBlock(type='table', rows=[['H1', 'H2'], ['a', 'b']])], output_path=str(out))
        assert result['success'] is True, result
        assert Path(out).exists()

    def test_ppt_single_slide_stacks_text_table_chart(self, tmp_path: Path) -> None:
        out = tmp_path / 'deck.pptx'
        result = create_office_document('pptx', [
            ContentBlock(type='paragraph', text='Title'),
            ContentBlock(type='table', rows=[['H1', 'H2'], ['a', 'b']]),
            ContentBlock(type='paragraph', text='Body'),
            ContentBlock(type='paragraph', chart_type='bar', chart_data=[['', 'S1'], ['Cat1', 10], ['Cat2', 20]]),
        ], output_path=str(out))
        assert result['success'] is True, result
        from tianshang_scribe.core.document import open_document
        eng = open_document(str(out))
        shapes = list(eng.prs.slides[0].shapes)
        assert any(sh.has_table for sh in shapes)
        assert any(sh.has_chart for sh in shapes)
        assert 'Title' in eng.extract_text()

    def test_excel_capabilities_via_content_block(self, tmp_path: Path) -> None:
        out = tmp_path / 'book.xlsx'
        result = create_office_document('xlsx', [
            ContentBlock(type='paragraph', text='hi', cell='A1', formula='=1+1',
                      freeze='A2', number_format='A1:A1=0.00%',
                      conditional_format='B1:B5=color_scale', data_validation='C1:C5=list:yes,no',
                      chart_type='bar', chart_data_range='Sheet!A1:B2'),
        ], output_path=str(out))
        assert result['success'] is True, result
        from tianshang_scribe.core.document import open_document
        eng = open_document(str(out))
        ws = eng.wb.active
        assert ws['A1'].value == '=1+1'
        assert eng.wb.active.freeze_panes == 'A2'
        assert ws['A1'].number_format == '0.00%'
        assert len(list(eng.wb.active.conditional_formatting)) >= 1
        assert 'list' in {dv.type for dv in eng.wb.active.data_validations.dataValidation}
        assert len(ws._charts) >= 1


class TestEditMcpP0:
    def test_excel_edit_actions(self, tmp_path: Path) -> None:
        from tianshang_scribe.core.document import open_document
        book = tmp_path / 'b.xlsx'
        create_office_document('xlsx', [ContentBlock(type='paragraph', text='seed')], output_path=str(book))
        res = edit_office_document(str(book), [
            {'action': 'write_cell', 'cell': 'A1', 'text': 'hello'},
            {'action': 'set_formula', 'cell': 'A2', 'formula': '=1+1'},
            {'action': 'freeze_panes', 'range': 'A2'},
            {'action': 'conditional_format', 'conditional_format': 'B1:B5=color_scale'},
            {'action': 'data_validation', 'data_validation': 'C1:C5=list:yes,no'},
            {'action': 'add_chart', 'chart_type': 'bar', 'chart_data_range': 'Sheet!A1:B2'},
        ])
        assert res['success'] is True, res
        e2 = open_document(str(book))
        ws = e2.wb.active
        assert ws['A1'].value == 'hello'
        assert ws['A2'].value == '=1+1'
        assert ws.freeze_panes == 'A2'
        assert len(list(ws.conditional_formatting)) >= 1
        assert 'list' in {dv.type for dv in ws.data_validations.dataValidation}
        assert len(ws._charts) >= 1

    def test_ppt_edit_actions(self, tmp_path: Path) -> None:
        from tianshang_scribe.core.document import open_document
        d = tmp_path / 'deck.pptx'
        create_office_document('pptx', [ContentBlock(type='paragraph', text='seed')], output_path=str(d))
        res = edit_office_document(str(d), [
            {'action': 'add_table', 'rows': [['H', 'K'], ['1', '2']], 'slide_index': 0},
            {'action': 'add_notes', 'notes': 'nt', 'slide_index': 0},
            {'action': 'apply_layout', 'layout': 'Title and Content', 'slide_index': 0},
            {'action': 'set_transition', 'transition': 'fade', 'slide_index': 0},
            {'action': 'add_chart', 'chart_type': 'bar', 'chart_data': [['', 'S'], ['c', 3]], 'slide_index': 0},
        ])
        assert res['success'] is True, res
        e2 = open_document(str(d))
        shapes = list(e2.prs.slides[0].shapes)
        assert any(s.has_table for s in shapes)
        assert any(s.has_chart for s in shapes)
        assert e2.prs.slides[0].notes_slide.notes_text_frame.text == 'nt'
        assert e2.prs.slides[0].slide_layout.name == 'Title and Content'


    def test_write_cell_respects_sheet_name(self, tmp_path: Path) -> None:
        from tianshang_scribe.core.document import create_document, DocumentType
        from tianshang_scribe.core.document import open_document
        book = tmp_path / 'b.xlsx'
        e = create_document(DocumentType.EXCEL)
        e.add_sheet('Sheet2')
        e.save(str(book))
        res = edit_office_document(str(book), [
            {'action': 'write_cell', 'cell': 'A1', 'text': 'two', 'sheet_name': 'Sheet2'},
            {'action': 'write_cell', 'cell': 'A1', 'text': 'one'},
        ])
        assert res['success'] is True, res
        e2 = open_document(str(book))
        assert e2.wb['Sheet2']['A1'].value == 'two'
        assert e2.wb.active['A1'].value == 'one'
        # formula path still works on the explicitly named sheet
        res2 = edit_office_document(str(book), [
            {'action': 'write_cell', 'cell': 'B1', 'text': '=1+1', 'sheet_name': 'Sheet2'},
        ])
        assert res2['success'] is True, res2
        e3 = open_document(str(book))
        assert e3.wb['Sheet2']['B1'].value == '=1+1'


    def test_ppt_blocks_stack_single_slide_without_index(self, tmp_path: Path) -> None:
        out = tmp_path / 'deck.pptx'
        result = create_office_document('pptx', [
            ContentBlock(type='paragraph', text='A'),
            ContentBlock(type='table', rows=[['H1', 'H2'], ['x', 'y']]),
            ContentBlock(type='paragraph', text='B'),
            ContentBlock(type='paragraph', chart_type='bar', chart_data=[['', 'S'], ['Cat1', 1], ['Cat2', 2]]),
        ], output_path=str(out))
        assert result['success'] is True, result
        from tianshang_scribe.core.document import open_document
        eng = open_document(str(out))
        assert len(eng.prs.slides) == 1
        shapes = list(eng.prs.slides[0].shapes)
        assert any(s.has_table for s in shapes)
        assert any(s.has_chart for s in shapes)
        assert 'A' in eng.extract_text() and 'B' in eng.extract_text()

    def test_ppt_explicit_slide_index_overrides_stack(self, tmp_path: Path) -> None:
        out = tmp_path / 'deck.pptx'
        result = create_office_document('pptx', [
            ContentBlock(type='paragraph', text='s0'),
            ContentBlock(type='page_break'),
            ContentBlock(type='paragraph', text='s1'),
            ContentBlock(type='table', rows=[['H', 'K'], ['1', '2']], slide_index=1),
        ], output_path=str(out))
        assert result['success'] is True, result
        from tianshang_scribe.core.document import open_document
        eng = open_document(str(out))
        assert len(eng.prs.slides) == 2
        shapes0 = list(eng.prs.slides[0].shapes)
        shapes1 = list(eng.prs.slides[1].shapes)
        assert not any(s.has_table for s in shapes0)
        assert any(s.has_table for s in shapes1)


class TestCompareMcpP0:
    def test_compare_excel_reports_unsupported(self, tmp_path: Path) -> None:
        from tianshang_scribe.core.document import create_document, DocumentType
        a = tmp_path / 'a.xlsx'
        b = tmp_path / 'b.xlsx'
        for p in (a, b):
            e = create_document(DocumentType.EXCEL)
            e.save(str(p))
        res = compare_documents(str(a), str(b))
        assert res['success'] is False
        assert res['error_code'] == 1003
        assert 'Excel' in res['error_message'] and 'PowerPoint' in res['error_message']


class TestEditSchemaMcpP1:
    def test_edit_operation_per_action_builds(self) -> None:
        from tianshang_scribe.mcp.schemas import EditOperation
        samples = [
            {'action': 'replace', 'old_text': 'a', 'new_text': 'b'},
            {'action': 'delete', 'target': 'x'},
            {'action': 'modify', 'old_text': 'a', 'new_text': 'b'},
            {'action': 'style', 'style': 'bold'},
            {'action': 'add', 'text': 'hi'},
            {'action': 'clear'},
            {'action': 'write_cell', 'cell': 'A1', 'text': 'v', 'style': 'bold'},
            {'action': 'set_formula', 'cell': 'A1', 'formula': '=1'},
            {'action': 'freeze_panes', 'range': 'A2'},
            {'action': 'add_chart', 'chart_type': 'bar', 'chart_data_range': 'S!A1:B2'},
            {'action': 'conditional_format', 'conditional_format': 'B1:B5=color_scale'},
            {'action': 'data_validation', 'data_validation': 'C1:C5=list:yes,no'},
            {'action': 'add_table', 'rows': [['H', 'K'], ['1', '2']]},
            {'action': 'add_picture', 'path': 'p.png'},
            {'action': 'add_shape', 'fill': 'FF0000'},
            {'action': 'apply_layout', 'layout': 'Title'},
            {'action': 'set_transition', 'transition': 'fade'},
            {'action': 'add_notes', 'notes': 'n'},
        ]
        for s in samples:
            op = EditOperation.model_validate(s)
            assert op.action == s['action']


class TestRegistryMcpP1:
    def test_tool_descriptions_mention_new_capabilities(self) -> None:
        from tianshang_scribe.mcp.tools._registry import get_tools
        tools = {t['name']: t['description'] for t in get_tools()}
        assert 'freeze_panes' in tools['edit_office_document']
        assert 'conditional_format' in tools['edit_office_document']
        assert 'data_validation' in tools['edit_office_document']
        assert 'sheet_name' in tools['create_office_document']
        assert 'slide_index' in tools['create_office_document']
        assert 'Excel' in tools['compare_documents'] and 'PowerPoint' in tools['compare_documents']


    def test_write_cell_with_style(self, tmp_path: Path) -> None:
        from tianshang_scribe.core.document import create_document, DocumentType
        from tianshang_scribe.core.document import open_document
        book = tmp_path / 'b.xlsx'
        e = create_document(DocumentType.EXCEL)
        e.save(str(book))
        res = edit_office_document(str(book), [
            {'action': 'write_cell', 'cell': 'A1', 'text': 'x', 'style': 'bold,fill=FF0000,align=center'},
        ])
        assert res['success'] is True, res
        e2 = open_document(str(book))
        cell = e2.wb.active['A1']
        assert cell.value == 'x'
        assert cell.font.bold is True
        assert 'FF0000' in str(cell.fill.fgColor.rgb)
        assert cell.alignment.horizontal == 'center'
