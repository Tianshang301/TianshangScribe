"""Unit tests for the one-shot CLI command dispatch (src/cli/main.py).

Uses a recording fake engine injected via ``create_document`` to exercise
every operation branch without touching the real document libraries.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from typer.testing import CliRunner

import src.cli.main as cli
from src.core.document import DocumentType

runner = CliRunner()


class _Base:
    """Fake engine exposing every method the CLI may call."""

    def __init__(self) -> None:
        self.calls: list[str] = []

    def set_style(self, *a, **k) -> None:
        self.calls.append('set_style')

    def apply_style_to_all(self, *a, **k) -> None:
        self.calls.append('apply_style_to_all')

    def get_base_style(self, *a, **k) -> object:
        class _S:
            def to_cli_string(self) -> str:
                return 'font=Test'

        return _S()

    def add_text(self, *a, **k) -> None:
        self.calls.append('add_text')

    def add_latex_content(self, *a, **k) -> None:
        self.calls.append('add_latex_content')

    def add_math_formula(self, *a, **k) -> None:
        self.calls.append('add_math_formula')

    def add_table_data(self, *a, **k) -> None:
        self.calls.append('add_table_data')

    def replace_text(self, *a, **k) -> int:
        self.calls.append('replace_text')
        return 1

    def add_heading(self, *a, **k) -> None:
        self.calls.append('add_heading')

    def set_metadata(self, **k) -> None:
        self.calls.append('set_metadata')

    def get_metadata(self) -> dict:
        self.calls.append('get_metadata')
        return {'author': 't'}

    def extract_text(self) -> str:
        self.calls.append('extract_text')
        return 'text'

    def extract_structure(self) -> dict:
        self.calls.append('extract_structure')
        return {'paragraphs': 1}

    def extract_tables(self) -> list:
        self.calls.append('extract_tables')
        return [['A', 'B']]

    def extract_images(self, *a, **k) -> list[str]:
        self.calls.append('extract_images')
        return ['img1.png']

    def merge_workbooks(self, *a, **k) -> None:
        self.calls.append('merge_workbooks')

    def split_by_sheet(self, *a, **k) -> list:
        self.calls.append('split_by_sheet')
        return [1, 2]

    def add_comment(self, *a, **k) -> None:
        self.calls.append('add_comment')

    def add_sheet(self, *a, **k) -> None:
        self.calls.append('add_sheet')

    def delete_sheet(self, *a, **k) -> None:
        self.calls.append('delete_sheet')

    def rename_sheet(self, *a, **k) -> None:
        self.calls.append('rename_sheet')

    def set_column_width(self, *a, **k) -> None:
        self.calls.append('set_column_width')

    def set_row_height(self, *a, **k) -> None:
        self.calls.append('set_row_height')

    def set_formula(self, *a, **k) -> None:
        self.calls.append('set_formula')

    def import_csv(self, *a, **k) -> None:
        self.calls.append('import_csv')

    def sort(self, *a, **k) -> None:
        self.calls.append('sort')

    def add_chart(self, **k) -> None:
        self.calls.append('add_chart')

    def set_protection(self, *a, **k) -> None:
        self.calls.append('set_protection')

    def unprotect(self, *a, **k) -> None:
        self.calls.append('unprotect')

    def clear_content(self, *a, **k) -> None:
        self.calls.append('clear_content')

    def clear_formats(self, *a, **k) -> None:
        self.calls.append('clear_formats')

    def clear_links(self, *a, **k) -> None:
        self.calls.append('clear_links')

    def add_slide(self, *a, **k) -> None:
        self.calls.append('add_slide')

    def delete_slide(self, *a, **k) -> None:
        self.calls.append('delete_slide')

    def move_slide(self, *a, **k) -> None:
        self.calls.append('move_slide')

    def add_notes(self, *a, **k) -> None:
        self.calls.append('add_notes')

    def apply_layout(self, *a, **k) -> None:
        self.calls.append('apply_layout')

    def set_transition(self, *a, **k) -> None:
        self.calls.append('set_transition')

    def compress_media(self, **k) -> int:
        self.calls.append('compress_media')
        return 100

    def add_toc(self, *a, **k) -> None:
        self.calls.append('add_toc')

    def add_section_break(self, *a, **k) -> None:
        self.calls.append('add_section_break')

    def set_header(self, *a, **k) -> None:
        self.calls.append('set_header')

    def set_footer(self, *a, **k) -> None:
        self.calls.append('set_footer')

    def add_watermark(self, *a, **k) -> None:
        self.calls.append('add_watermark')

    def to_images(self, *a, **k) -> list:
        self.calls.append('to_images')
        return ['1.png']

    def to_pdf(self, *a, **k) -> None:
        self.calls.append('to_pdf')

    def export_csv(self, *a, **k) -> None:
        self.calls.append('export_csv')

    def export_json(self, *a, **k) -> None:
        self.calls.append('export_json')

    def export_html(self, *a, **k) -> None:
        self.calls.append('export_html')

    def save(self, *a, **k) -> None:
        self.calls.append('save')


@pytest.fixture
def fake_engine(monkeypatch) -> _Base:
    engine = _Base()

    def _create(doc_type: DocumentType) -> _Base:
        engine.doc_type = doc_type
        return engine

    def _open(path: str) -> _Base:
        engine.opened = path
        return engine

    monkeypatch.setattr(cli, 'create_document', _create)
    monkeypatch.setattr(cli, 'open_document', _open)
    return engine


@pytest.fixture
def tmp_in(tmp_path: Path) -> Path:
    p = tmp_path / 'in.docx'
    p.write_bytes(b'x' * 10)
    return p


def _run(*args: str) -> tuple[int, str]:
    result = runner.invoke(cli.app, list(args))
    return result.exit_code, result.stdout


class TestBootstrap:
    def test_no_args_prints_help(self) -> None:
        code, out = _run()
        assert code == 0
        assert 'Cross-platform' in out

    def test_no_input_no_create_exit2(self, tmp_path: Path, monkeypatch) -> None:
        engine = _Base()
        monkeypatch.setattr(cli, 'create_document', lambda t: engine)
        code, out = _run('--force', '--add', 'x')
        assert code == 0
        assert 'Cross-platform' in out

    def test_stdin_empty_exit2(self, monkeypatch) -> None:
        import io
        import sys

        monkeypatch.setattr(sys, 'stdin', type('S', (), {'buffer': io.BytesIO(b'')})())
        code, out = _run('--stdin', '--force')
        assert code == 2
        assert 'No data' in out

    def test_stdin_with_data(self, monkeypatch, tmp_path: Path) -> None:
        engine = _Base()
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        out = tmp_path / 'o.docx'
        result = runner.invoke(
            cli.app, ['--stdin', '--output', str(out), '--force'], input=b'12345'
        )
        assert result.exit_code == 0
        assert 'save' in engine.calls

    def test_help_flag(self) -> None:
        code, out = _run('--help')
        assert code == 0
        assert 'Usage' in out


class TestBatchValidation:
    def test_files_no_match(self, tmp_path: Path) -> None:
        code, out = _run('--files', str(tmp_path / '*.nomatch'))
        assert code == 2
        assert 'No files match' in out

    def test_input_glob_no_match(self, tmp_path: Path) -> None:
        code, out = _run(str(tmp_path / '*.nomatch'))
        assert code == 2
        assert 'No files match' in out

    def test_batch_stdout_rejected(self, tmp_path: Path) -> None:
        p = tmp_path / 'a.docx'
        p.write_bytes(b'x')
        code, out = _run('--batch', '--stdout', str(p))
        assert code == 2
        assert '--stdout' in out

    def test_batch_create_rejected(self, tmp_path: Path) -> None:
        p = tmp_path / 'a.docx'
        p.write_bytes(b'x')
        code, out = _run('--batch', '--create', str(p))
        assert code == 2
        assert 'not compatible' in out


class TestCreateDispatch:
    def test_create_word(self, fake_engine, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run('--create', '--word', '--output', str(out), '--force')
        assert code == 0
        assert fake_engine.doc_type == DocumentType.WORD
        assert 'save' in fake_engine.calls

    def test_create_excel(self, fake_engine, tmp_path: Path) -> None:
        out = tmp_path / 'o.xlsx'
        code, _ = _run('--create', '--excel', '--output', str(out), '--force')
        assert code == 0
        assert fake_engine.doc_type == DocumentType.EXCEL

    def test_create_ppt(self, fake_engine, tmp_path: Path) -> None:
        out = tmp_path / 'o.pptx'
        code, _ = _run('--create', '--ppt', '--output', str(out), '--force')
        assert code == 0
        assert fake_engine.doc_type == DocumentType.PPT

    def test_open_existing(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run('--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert fake_engine.opened == str(tmp_in)


class TestOverwriteGuard:
    def test_existing_output_requires_force(self, fake_engine, tmp_in: Path) -> None:
        code, out = _run('--output', str(tmp_in), str(tmp_in))
        assert code == 1
        assert 'already exists' in out

    def test_no_input_and_no_create(self, fake_engine) -> None:
        code, _out = _run('--force')
        assert code == 0


class TestWordOps:
    def test_style_add_replace_delete_modify(
        self, fake_engine, tmp_in: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run(
            '--style',
            'font=Test,size=12',
            '--add',
            'hello',
            '--replace',
            'a',
            '--replace-new',
            'b',
            '--delete',
            'x',
            '--modify',
            'old',
            '--modify-new',
            'new',
            '--output',
            str(out),
            '--force',
            str(tmp_in),
        )
        assert code == 0
        assert 'set_style' in fake_engine.calls
        assert 'apply_style_to_all' in fake_engine.calls
        assert 'add_text' in fake_engine.calls
        assert 'replace_text' in fake_engine.calls
        assert 'clear' not in fake_engine.calls

    def test_replace_without_new(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, out_txt = _run('--replace', 'a', '--output', str(out), '--force', str(tmp_in))
        assert code == 2
        assert 'requires --replace-new' in out_txt

    def test_heading_parsed(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run(
            '--heading', 'level:2 text:Title', '--output', str(out), '--force', str(tmp_in)
        )
        assert code == 0
        assert 'add_heading' in fake_engine.calls

    def test_heading_plain(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run('--heading', 'plain title', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'add_text' in fake_engine.calls

    def test_heading_no_match(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run('--heading', 'Title', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'add_text' in fake_engine.calls

    def test_heading_no_add_heading_method(self, tmp_in: Path, tmp_path: Path) -> None:
        class _NoHeading:
            def __init__(self) -> None:
                self.calls: list[str] = []

            def add_text(self, *a, **k) -> None:
                self.calls.append('add_text')

            def save(self, *a, **k) -> None:
                pass

        engine = _NoHeading()
        monkeypatch = pytest.MonkeyPatch()
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        out = tmp_path / 'o.docx'
        code, _ = _run('--heading', 'level:1 text:H', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'add_text' in engine.calls
        monkeypatch.undo()

    def test_style_without_input(self, tmp_path: Path, monkeypatch) -> None:
        class _StyleOnly:
            def __init__(self) -> None:
                self.calls: list[str] = []

            def set_style(self, *a, **k) -> None:
                self.calls.append('set_style')

            def get_base_style(self) -> object:
                class S:
                    def to_cli_string(self) -> str:
                        return 'font=Test'

                return S()

            def save(self, *a, **k) -> None:
                self.calls.append('save')

        engine = _StyleOnly()
        monkeypatch.setattr(cli, 'create_document', lambda t: engine)
        out = tmp_path / 'o.docx'
        code, _ = _run(
            '--create', '--word', '--style', 'font=Test', '--output', str(out), '--force'
        )
        assert code == 0
        assert 'set_style' in engine.calls
        assert 'apply_style_to_all' not in engine.calls

    def test_math_word(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run('--math', r'\frac{a}{b}', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'add_math_formula' in fake_engine.calls

    def test_add_table_word(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run('--add-table', 'H1,H2|a1,a2', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'add_table_data' in fake_engine.calls

    def test_toc_section_header_footer_watermark(
        self, fake_engine, tmp_in: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run(
            '--toc',
            '--section-break',
            '--header',
            'H',
            '--footer',
            'F',
            '--watermark',
            'W',
            '--output',
            str(out),
            '--force',
            str(tmp_in),
        )
        assert code == 0
        for name in ('add_toc', 'add_section_break', 'set_header', 'set_footer', 'add_watermark'):
            assert name in fake_engine.calls

    def test_latex_style_add(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run(
            '--latex-style',
            '--add',
            r'\bfseries{x}',
            '--output',
            str(out),
            '--force',
            str(tmp_in),
        )
        assert code == 0
        assert 'add_latex_content' in fake_engine.calls


class TestExcelOps:
    def test_excel_full_ops(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.xlsx'
        code, _ = _run(
            '--excel',
            '--sheet-add',
            'S1',
            '--sheet-delete',
            'S0',
            '--sheet-rename',
            'A B',
            '--column-width',
            '2=20',
            '--row-height',
            '3=30',
            '--formula',
            'A1 =SUM(B1:B10)',
            '--sort',
            'A1:A10 asc',
            '--protect',
            'pw',
            '--unprotect',
            '--clear',
            'content',
            '--output',
            str(out),
            '--force',
            str(tmp_in),
        )
        assert code == 0
        for name in (
            'add_sheet',
            'delete_sheet',
            'rename_sheet',
            'set_column_width',
            'set_row_height',
            'set_formula',
            'sort',
            'set_protection',
            'unprotect',
            'clear_content',
        ):
            assert name in fake_engine.calls

    def test_clear_formats(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.xlsx'
        code, _ = _run(
            '--excel', '--clear', 'formats', '--output', str(out), '--force', str(tmp_in)
        )
        assert code == 0
        assert 'clear_formats' in fake_engine.calls

    def test_clear_links(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.xlsx'
        code, _ = _run('--excel', '--clear', 'links', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'clear_links' in fake_engine.calls

    def test_export_csv_json(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.csv'
        code, _ = _run('--excel', '--to-csv', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'export_csv' in fake_engine.calls
        out2 = tmp_path / 'o.json'
        code, _ = _run('--excel', '--to-json', '--output', str(out2), '--force', str(tmp_in))
        assert code == 0
        assert 'export_json' in fake_engine.calls

    def test_merge_workbooks(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.xlsx'
        p2 = tmp_path / 'b.xlsx'
        p2.write_bytes(b'x')
        code, _ = _run(
            '--excel',
            '--merge',
            f'{p2}',
            '--output',
            str(out),
            '--force',
            str(tmp_in),
        )
        assert code == 0
        assert 'merge_workbooks' in fake_engine.calls

    def test_split_sheets(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.xlsx'
        code, _ = _run(
            '--excel', '--split', 'by-sheet', '--output', str(out), '--force', str(tmp_in)
        )
        assert code == 0
        assert 'split_by_sheet' in fake_engine.calls

    def test_chart(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.xlsx'
        code, _ = _run(
            '--excel',
            '--chart-add',
            'type=bar data=B1:C10 pos=E2',
            '--output',
            str(out),
            '--force',
            str(tmp_in),
        )
        assert code == 0
        assert 'add_chart' in fake_engine.calls

    def test_from_csv(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.xlsx'
        csvf = tmp_path / 'd.csv'
        csvf.write_text('a,b\n1,2\n', encoding='utf-8')
        code, _ = _run(
            '--excel',
            '--from-csv',
            str(csvf),
            '--output',
            str(out),
            '--force',
            str(tmp_in),
        )
        assert code == 0
        assert 'import_csv' in fake_engine.calls


class TestPptOps:
    def test_ppt_full_ops(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.pptx'
        code, _ = _run(
            '--ppt',
            '--slide-add',
            '--slide-delete',
            '2',
            '--slide-move',
            '1 2',
            '--notes',
            '1 hi',
            '--layout',
            '1 Title and Content',
            '--transition',
            'fade',
            '--compress-media',
            '1920,80',
            '--output',
            str(out),
            '--force',
            str(tmp_in),
        )
        assert code == 0
        for name in (
            'add_slide',
            'delete_slide',
            'move_slide',
            'add_notes',
            'apply_layout',
            'set_transition',
            'compress_media',
        ):
            assert name in fake_engine.calls

    def test_ppt_to_images(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        outdir = tmp_path / 'imgs'
        outdir.mkdir()
        code, _ = _run('--ppt', '--toimg', '--output', str(outdir), '--force', str(tmp_in))
        assert code == 0
        assert 'to_images' in fake_engine.calls

    def test_transition_with_index(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.pptx'
        code, _ = _run(
            '--ppt', '--transition', '1 fade', '--output', str(out), '--force', str(tmp_in)
        )
        assert code == 0
        assert 'set_transition' in fake_engine.calls


class TestExtractModes:
    def test_extract_text(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        code, out = _run('--extract', 'text', str(tmp_in))
        assert code == 0
        assert 'text' in out

    def test_extract_metadata(self, fake_engine, tmp_in: Path) -> None:
        code, out = _run('--extract', 'metadata', str(tmp_in))
        assert code == 0
        assert 'author' in out

    def test_extract_structure(self, fake_engine, tmp_in: Path) -> None:
        code, out = _run('--extract', 'structure', str(tmp_in))
        assert code == 0
        assert 'paragraphs' in out

    def test_extract_tables(self, fake_engine, tmp_in: Path) -> None:
        code, out = _run('--extract', 'tables', str(tmp_in))
        assert code == 0
        assert 'A' in out

    def test_extract_images(self, fake_engine, tmp_in: Path) -> None:
        code, out = _run('--extract', 'images', str(tmp_in))
        assert code == 0
        assert 'img1.png' in out

    def test_extract_images_no_input(self, fake_engine, tmp_path: Path) -> None:
        class _C:
            def __init__(self) -> None:
                self.calls: list[str] = []

            def extract_images(self, d: str) -> list[str]:
                self.calls.append(d)
                return ['a.png']

            def save(self, *a, **k) -> None:
                pass

        engine = _C()
        monkeypatch = pytest.MonkeyPatch()
        monkeypatch.setattr(cli, 'create_document', lambda t: engine)
        code, _ = _run('--create', '--word', '--extract', 'images')
        assert code == 0
        assert engine.calls == ['images']
        monkeypatch.undo()

    def test_extract_images_to_output(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        outdir = tmp_path / 'img'
        outdir.mkdir()
        code, out = _run('--extract', 'images', '--output', str(outdir), '--force', str(tmp_in))
        assert code == 0
        assert 'img1.png' in out

    def test_extract_unknown_mode(self, fake_engine, tmp_in: Path) -> None:
        code, out = _run('--extract', 'bogus', str(tmp_in))
        assert code == 0
        assert 'not supported' in out


class TestTemplateAndMeta:
    def test_template_fill(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        data = tmp_path / 'd.json'
        data.write_text('{"name": "x"}', encoding='utf-8')
        out = tmp_path / 'o.docx'
        code, _ = _run('--template', str(data), '--output', str(out), '--force', str(tmp_in))
        assert code == 0

    def test_meta(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run(
            '--meta', 'title="T",author="A"', '--output', str(out), '--force', str(tmp_in)
        )
        assert code == 0
        assert 'set_metadata' in fake_engine.calls


class TestComment:
    def test_comment_int(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run('--comment', '3 hi', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'add_comment' in fake_engine.calls

    def test_comment_keyword(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run('--comment', 'kw text', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'add_comment' in fake_engine.calls

    def test_comment_single(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _ = _run('--comment', 'note', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'add_comment' in fake_engine.calls


class TestUnsupportedWarnings:
    def test_latex_fallback(self, tmp_in: Path, tmp_path: Path) -> None:
        class _NoLatex:
            def __init__(self) -> None:
                self.calls: list[str] = []

            def add_text(self, *a, **k) -> None:
                self.calls.append('add_text')

            def save(self, *a, **k) -> None:
                self.calls.append('save')

        engine = _NoLatex()
        monkeypatch = pytest.MonkeyPatch()
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        out = tmp_path / 'o.docx'
        code, out_txt = _run(
            '--latex-style', '--add', 'x', '--output', str(out), '--force', str(tmp_in)
        )
        assert code == 0
        assert 'not supported' in out_txt
        assert 'add_text' in engine.calls
        monkeypatch.undo()

    def test_merge_unsupported(self, tmp_in: Path, tmp_path: Path) -> None:
        class _NoMerge:
            def save(self, *a, **k) -> None:
                pass

        engine = _NoMerge()
        monkeypatch = pytest.MonkeyPatch()
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        out = tmp_path / 'o.docx'
        code, out_txt = _run('--merge', 'x.docx', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'Merge not supported' in out_txt
        monkeypatch.undo()

    def test_split_unsupported(self, tmp_in: Path, tmp_path: Path) -> None:
        class _NoSplit:
            def save(self, *a, **k) -> None:
                pass

        engine = _NoSplit()
        monkeypatch = pytest.MonkeyPatch()
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        out = tmp_path / 'o.docx'
        code, out_txt = _run('--split', 'by-page', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'Split not supported' in out_txt
        monkeypatch.undo()

    def test_clear_formats_unsupported(self, tmp_in: Path, tmp_path: Path) -> None:
        class _NoClearFormats:
            def clear_content(self, *a, **k) -> None:
                pass

            def save(self, *a, **k) -> None:
                pass

        engine = _NoClearFormats()
        monkeypatch = pytest.MonkeyPatch()
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        out = tmp_path / 'o.docx'
        code, out_txt = _run('--clear', 'formats', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'Format clearing not supported' in out_txt
        monkeypatch.undo()

    def test_clear_links_unsupported(self, tmp_in: Path, tmp_path: Path) -> None:
        class _NoClearLinks:
            def clear_content(self, *a, **k) -> None:
                pass

            def save(self, *a, **k) -> None:
                pass

        engine = _NoClearLinks()
        monkeypatch = pytest.MonkeyPatch()
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        out = tmp_path / 'o.docx'
        code, out_txt = _run('--clear', 'links', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'Link clearing not supported' in out_txt
        monkeypatch.undo()

    def test_add_table_non_word(self, tmp_in: Path, tmp_path: Path) -> None:
        class _NoTable:
            def __init__(self) -> None:
                self.calls: list[str] = []

            def save(self, *a, **k) -> None:
                self.calls.append('save')

        engine = _NoTable()
        monkeypatch = pytest.MonkeyPatch()
        monkeypatch.setattr(cli, 'create_document', lambda t: engine)
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        out = tmp_path / 'o.docx'
        code, out_txt = _run('--add-table', 'A,B|1,2', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'only supported for Word' in out_txt
        assert 'save' in engine.calls
        monkeypatch.undo()

    def test_math_non_word(self, tmp_in: Path, tmp_path: Path) -> None:
        class _NoMath:
            def save(self, *a, **k) -> None:
                pass

        monkeypatch = pytest.MonkeyPatch()
        engine = _NoMath()
        monkeypatch.setattr(cli, 'create_document', lambda t: engine)
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        out = tmp_path / 'o.docx'
        code, out_txt = _run('--math', 'x', '--output', str(out), '--force', str(tmp_in))
        assert code == 2
        assert 'only supported for Word' in out_txt
        monkeypatch.undo()


class TestBatchProcessing:
    def test_batch_success_and_failure(self, fake_engine, tmp_path: Path, monkeypatch) -> None:
        a = tmp_path / 'a.docx'
        b = tmp_path / 'b.docx'
        a.write_bytes(b'x')
        b.write_bytes(b'x')
        calls: list[str] = []
        real_save = fake_engine.save

        def fail_once(path: str) -> None:
            nonlocal calls
            calls.append(path)
            if path.endswith('b-out.docx') and not fake_engine._failed:
                fake_engine._failed = True
                raise FileNotFoundError('boom')

        fake_engine._failed = False
        fake_engine.save = fail_once  # type: ignore[assignment]
        code, out = _run('--files', str(tmp_path / '*.docx'), '--force')
        assert 'Batch complete' in out
        assert '1 succeeded' in out
        assert '1 failed' in out
        assert code == 1
        monkeypatch.setattr(fake_engine, 'save', real_save)

    def test_batch_all_success(self, fake_engine, tmp_path: Path) -> None:
        a = tmp_path / 'a.docx'
        b = tmp_path / 'b.docx'
        a.write_bytes(b'x')
        b.write_bytes(b'x')
        code, out = _run('--files', str(tmp_path / '*.docx'), '--force')
        assert code == 0
        assert '2 succeeded' in out


class TestReverseConversion:
    def test_markdown_to_word(self, tmp_path: Path, monkeypatch) -> None:
        md = tmp_path / 'd.md'
        md.write_text('# Hi', encoding='utf-8')
        out = tmp_path / 'o.docx'
        calls = []

        def fake_md2word(src, dst) -> None:
            calls.append((src, dst))

        import src.transform.reverse as rev

        monkeypatch.setattr(rev, 'markdown_to_word', fake_md2word)
        code, _ = _run('--output', str(out), '--force', str(md))
        assert code == 0
        assert calls

    def test_html_to_word(self, tmp_path: Path, monkeypatch) -> None:
        html = tmp_path / 'd.html'
        html.write_text('<h1>Hi</h1>', encoding='utf-8')
        out = tmp_path / 'o.docx'
        calls = []
        import src.transform.reverse as rev

        monkeypatch.setattr(rev, 'html_to_word', lambda s, d: calls.append((s, d)))
        code, _ = _run('--output', str(out), '--force', str(html))
        assert code == 0
        assert calls

    def test_json_to_excel(self, tmp_path: Path, monkeypatch) -> None:
        data = tmp_path / 'd.json'
        data.write_text('[{"a":1}]', encoding='utf-8')
        out = tmp_path / 'o.xlsx'
        calls = []

        class _E:
            def __init__(self) -> None:
                pass

            def create(self) -> None:
                pass

            def import_json(self, p: str) -> None:
                calls.append(p)

            def save(self, p: str) -> None:
                calls.append(p)

        import src.core.excel_engine as ee

        monkeypatch.setattr(ee, 'ExcelEngine', _E)
        code, _ = _run('--output', str(out), '--force', str(data))
        assert code == 0
        assert calls

    def test_unsupported_source(self, tmp_path: Path) -> None:
        src = tmp_path / 'd.xyz'
        src.write_text('x')
        out = tmp_path / 'o.docx'
        code, _ = _run('--output', str(out), '--force', str(src))
        assert code != 0

    def test_markdown_to_word_topdf(self, tmp_path: Path, monkeypatch) -> None:
        md = tmp_path / 'd.md'
        md.write_text('# Hi', encoding='utf-8')
        out = tmp_path / 'o.pdf'
        calls = []
        import src.transform.pdf as pdf
        import src.transform.reverse as rev

        monkeypatch.setattr(rev, 'markdown_to_word', lambda s, d: calls.append((s, d)))
        monkeypatch.setattr(pdf, 'word_to_pdf', lambda s, d: calls.append(('pdf', s, d)))
        code, _ = _run('--topdf', '--output', str(out), '--force', str(md))
        assert code == 0
        assert any(c[0] == 'pdf' for c in calls)

    def test_json_to_excel_topdf(self, tmp_path: Path, monkeypatch) -> None:
        data = tmp_path / 'd.json'
        data.write_text('[{"a":1}]', encoding='utf-8')
        out = tmp_path / 'o.pdf'
        calls = []
        import src.core.excel_engine as ee
        import src.transform.pdf as pdf

        class _E:
            def create(self) -> None:
                pass

            def import_json(self, p: str) -> None:
                calls.append(p)

            def save(self, p: str) -> None:
                calls.append(p)

        monkeypatch.setattr(ee, 'ExcelEngine', _E)
        monkeypatch.setattr(pdf, 'word_to_pdf', lambda s, d: calls.append(('pdf', s, d)))
        code, _ = _run('--topdf', '--output', str(out), '--force', str(data))
        assert code == 0
        assert any(c[0] == 'pdf' for c in calls)

    def test_reverse_unsupported_suffix(self, tmp_path: Path) -> None:
        src = tmp_path / 'd.xyz'
        src.write_text('x')
        out = tmp_path / 'o.docx'
        code, _ = _run('--output', str(out), '--force', str(src))
        assert code != 0


class TestOutputFormats:
    def test_tomd(self, fake_engine, tmp_in: Path, tmp_path: Path, monkeypatch) -> None:
        out = tmp_path / 'o.md'
        import src.transform.pdf as pdf

        monkeypatch.setattr(pdf, 'word_to_markdown', lambda s, d: None)
        code, _ = _run('--tomd', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert out.exists() or 'save' in fake_engine.calls

    def test_tomd_conversion_failure_falls_back(
        self, fake_engine, tmp_in: Path, tmp_path: Path, monkeypatch
    ) -> None:
        out = tmp_path / 'o.md'
        import src.transform.pdf as pdf

        monkeypatch.setattr(
            pdf, 'word_to_markdown', lambda s, d: (_ for _ in ()).throw(RuntimeError('boom'))
        )
        code, _ = _run('--tomd', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'save' in fake_engine.calls

    def test_tohtml_word_export(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.html'
        code, _ = _run('--tohtml', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'export_html' in fake_engine.calls

    def test_tohtml_word_fallback(self, tmp_in: Path, tmp_path: Path, monkeypatch) -> None:
        class _NoExportHtml:
            def __init__(self) -> None:
                self.calls: list[str] = []

            def save(self, *a, **k) -> None:
                self.calls.append('save')

        engine = _NoExportHtml()
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        import src.transform.pdf as pdf

        monkeypatch.setattr(pdf, 'word_to_html', lambda s, d: None)
        out = tmp_path / 'o.html'
        code, _ = _run('--tohtml', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'save' in engine.calls

    def test_tohtml_word_fallback_failure(self, tmp_in: Path, tmp_path: Path, monkeypatch) -> None:
        class _NoExportHtml:
            def __init__(self) -> None:
                self.calls: list[str] = []

            def save(self, *a, **k) -> None:
                self.calls.append('save')

        engine = _NoExportHtml()
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        import src.transform.pdf as pdf

        monkeypatch.setattr(
            pdf, 'word_to_html', lambda s, d: (_ for _ in ()).throw(RuntimeError('boom'))
        )
        out = tmp_path / 'o.html'
        code, _ = _run('--tohtml', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'save' in engine.calls

    def test_topdf(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.pdf'
        code, _ = _run('--topdf', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'to_pdf' in fake_engine.calls

    def test_stdout(self, fake_engine, tmp_in: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.docx'
        code, _out_txt = _run('--stdout', '--output', str(out), '--force', str(tmp_in))
        assert code == 0
        assert 'save' in fake_engine.calls


class TestErrorHandling:
    def test_file_not_found(self, tmp_path: Path) -> None:
        missing = tmp_path / 'missing.docx'
        code, _out = _run('--force', str(missing))
        assert code == 1

    def test_value_error(self, tmp_in: Path, tmp_path: Path, monkeypatch) -> None:
        class _Bad:
            def save(self, *a, **k) -> None:
                raise ValueError('bad')

        monkeypatch.setattr(cli, 'open_document', lambda p: _Bad())
        out = tmp_path / 'o.docx'
        code, out_txt = _run('--output', str(out), '--force', str(tmp_in))
        assert code == 2
        assert 'bad' in out_txt

    def test_not_implemented(self, tmp_in: Path, tmp_path: Path, monkeypatch) -> None:
        class _NI:
            def save(self, *a, **k) -> None:
                raise NotImplementedError('nope')

        monkeypatch.setattr(cli, 'open_document', lambda p: _NI())
        out = tmp_path / 'o.docx'
        code, out_txt = _run('--output', str(out), '--force', str(tmp_in))
        assert code == 3
        assert 'nope' in out_txt

    def test_unexpected_exception(self, tmp_in: Path, tmp_path: Path, monkeypatch) -> None:
        class _U:
            def add_text(self, *a, **k) -> None:
                raise RuntimeError('boom')

        monkeypatch.setattr(cli, 'open_document', lambda p: _U())
        out = tmp_path / 'o.docx'
        code, out_txt = _run('--add', 'x', '--output', str(out), '--force', str(tmp_in))
        assert code == 1
        assert 'boom' in out_txt


class TestHelperFunctions:
    def test_no_expand(self) -> None:
        assert cli._no_expand(['a', 'b*'], c=1) == ['a', 'b*']

    def test_parse_chart_opts(self) -> None:
        assert cli._parse_chart_opts('type=bar data=B1 pos=E2') == {
            'type': 'bar',
            'data': 'B1',
            'pos': 'E2',
        }

    def test_parse_heading(self, monkeypatch) -> None:
        calls = []
        engine = _Base()
        engine.add_heading = lambda t, level=None: calls.append((t, level))
        monkeypatch.setattr(engine, 'add_heading', engine.add_heading)
        cli._parse_heading(engine, 'level:3 text:Sub')
        assert calls == [('Sub', 3)]

    def test_apply_meta(self, monkeypatch) -> None:
        engine = _Base()
        got: dict = {}

        def set_md(**k) -> None:
            got.update(k)

        engine.set_metadata = set_md
        cli._apply_meta(engine, 'a=1, b=2')
        assert got == {'a': '1', 'b': '2'}


class TestOpenApp:
    def test_open_dispatch_calls_session(self, tmp_path: Path, monkeypatch) -> None:
        docx = tmp_path / 't.docx'
        docx.write_bytes(b'x' * 10)
        calls = []

        class _S:
            def __init__(self, *a, **k) -> None:
                calls.append(('init', a, k))

            def run(self) -> None:
                calls.append(('run',))

        import src.cli.repl as repl

        monkeypatch.setattr(repl, 'InteractiveSession', _S)
        monkeypatch.setattr(cli, 'open_document', lambda p: _Base())
        result = runner.invoke(cli.open_app, [str(docx)])
        assert result.exit_code == 0
        assert any(c[0] == 'run' for c in calls)

    def test_open_force_type(self, tmp_path: Path, monkeypatch) -> None:
        docx = tmp_path / 't.docx'
        docx.write_bytes(b'x' * 10)
        calls = []

        class _E:
            def __init__(self) -> None:
                pass

            def open(self, p: str) -> None:
                calls.append(p)

        import src.cli.repl as repl
        import src.core.word_engine as we

        monkeypatch.setattr(cli, 'detect_document_type', lambda p: DocumentType.EXCEL)
        monkeypatch.setattr(we, 'WordEngine', _E)
        monkeypatch.setattr(
            repl,
            'InteractiveSession',
            type('S', (), {'__init__': lambda self, *a, **k: None, 'run': lambda self: None}),
        )
        result = runner.invoke(cli.open_app, ['--word', str(docx)])
        assert result.exit_code == 0
        assert calls == [str(docx)]

    def test_open_no_force_type(self, tmp_path: Path, monkeypatch) -> None:
        docx = tmp_path / 't.docx'
        docx.write_bytes(b'x' * 10)
        monkeypatch.setattr(cli, 'detect_document_type', lambda p: DocumentType.UNKNOWN)
        engine = _Base()
        monkeypatch.setattr(cli, 'open_document', lambda p: engine)
        import src.cli.repl as repl

        monkeypatch.setattr(
            repl,
            'InteractiveSession',
            type('S', (), {'__init__': lambda self, *a, **k: None, 'run': lambda self: None}),
        )
        result = runner.invoke(cli.open_app, [str(docx)])
        assert result.exit_code == 0


def test_open_engine_unsupported_explicit(monkeypatch) -> None:
    class _FakeDocType:
        pass

    fake = _FakeDocType()
    monkeypatch.setattr(cli, 'detect_document_type', lambda p: DocumentType.WORD)
    with pytest.raises(ValueError):
        cli._open_engine('x.docx', fake)  # type: ignore[arg-type]


def test_open_excel_explicit(tmp_path: Path, monkeypatch) -> None:
    docx = tmp_path / 't.xlsx'
    docx.write_bytes(b'x' * 10)
    import src.cli.repl as repl
    import src.core.excel_engine as ee

    class _E:
        def __init__(self) -> None:
            pass

        def open(self, p: str) -> None:
            pass

    monkeypatch.setattr(cli, 'detect_document_type', lambda p: DocumentType.WORD)
    monkeypatch.setattr(ee, 'ExcelEngine', _E)
    monkeypatch.setattr(
        repl,
        'InteractiveSession',
        type('S', (), {'__init__': lambda self, *a, **k: None, 'run': lambda self: None}),
    )
    result = runner.invoke(cli.open_app, ['--excel', str(docx)])
    assert result.exit_code == 0


def test_open_ppt_explicit(tmp_path: Path, monkeypatch) -> None:
    docx = tmp_path / 't.pptx'
    docx.write_bytes(b'x' * 10)
    import src.cli.repl as repl
    import src.core.ppt_engine as pe

    class _E:
        def __init__(self) -> None:
            pass

        def open(self, p: str) -> None:
            pass

    monkeypatch.setattr(cli, 'detect_document_type', lambda p: DocumentType.WORD)
    monkeypatch.setattr(pe, 'PptEngine', _E)
    monkeypatch.setattr(
        repl,
        'InteractiveSession',
        type('S', (), {'__init__': lambda self, *a, **k: None, 'run': lambda self: None}),
    )
    result = runner.invoke(cli.open_app, ['--ppt', str(docx)])
    assert result.exit_code == 0


class TestMainDispatch:
    def test_main_cli_dispatch_open(self, monkeypatch) -> None:
        import sys

        calls = []
        monkeypatch.setattr(sys, 'argv', ['tianshang-scribe', 'open', 'x.docx'])
        monkeypatch.setattr(cli, 'open_app', lambda args: calls.append(args))
        cli.main_cli()
        assert calls == [['x.docx']]

    def test_main_cli_dispatch_app(self, monkeypatch) -> None:
        import sys

        calls = []
        monkeypatch.setattr(sys, 'argv', ['tianshang-scribe', '--force'])
        monkeypatch.setattr(cli, 'app', lambda: calls.append('app'))
        cli.main_cli()
        assert calls == ['app']


class TestScheduleCli:
    def test_schedule_add_list(self, tmp_path: Path) -> None:
        db = tmp_path / 's.db'
        code, out = _run('--schedule-db', str(db), '--schedule-add', 'daily|0 9 * * *|echo hi')
        assert code == 0
        assert 'daily' in out
        code, out = _run('--schedule-db', str(db), '--schedule-list')
        assert code == 0
        assert 'daily' in out
        assert '0 9 * * *' in out

    def test_schedule_add_invalid(self, tmp_path: Path) -> None:
        db = tmp_path / 's.db'
        code, _ = _run('--schedule-db', str(db), '--schedule-add', 'onlyname')
        assert code == 2

    def test_schedule_rm(self, tmp_path: Path) -> None:
        db = tmp_path / 's.db'
        _run('--schedule-db', str(db), '--schedule-add', 'x|0 9 * * *|echo hi')
        code, out = _run('--schedule-db', str(db), '--schedule-rm', 'x')
        assert code == 0
        assert 'removed' in out
        code, out = _run('--schedule-db', str(db), '--schedule-rm', 'missing')
        assert code == 0
        assert 'not found' in out

    def test_schedule_list_empty(self, tmp_path: Path) -> None:
        db = tmp_path / 's.db'
        code, out = _run('--schedule-db', str(db), '--schedule-list')
        assert code == 0
        assert 'No schedules' in out

    def test_schedule_run_unknown(self, tmp_path: Path) -> None:
        db = tmp_path / 's.db'
        code, _ = _run('--schedule-db', str(db), '--schedule-run', 'nope')
        assert code == 1

    def test_schedule_run_success(self, tmp_path: Path) -> None:
        db = tmp_path / 's.db'
        _run('--schedule-db', str(db), '--schedule-add', 'ok|0 9 * * *|python -c pass')
        code, out = _run('--schedule-db', str(db), '--schedule-run', 'ok')
        assert code == 0
        assert 'ok:' in out

    def test_schedule_run_unsatisfied_dep(self, tmp_path: Path) -> None:
        db = tmp_path / 's.db'
        _run('--schedule-db', str(db), '--schedule-add', 'child|0 9 * * *|python -c pass')
        from src.utils.store import Schedule, ScheduleStore

        with ScheduleStore(db) as store:
            store.upsert(
                Schedule(
                    name='child',
                    cron='0 9 * * *',
                    command=['python', '-c', 'pass'],
                    depends_on=['base'],
                )
            )
        code, out = _run('--schedule-db', str(db), '--schedule-run', 'child')
        assert code == 1
        assert 'unsatisfied' in out

    def test_schedule_run_all_due(self, tmp_path: Path, monkeypatch) -> None:
        import datetime as _dt

        class _FixedDT(_dt.datetime):
            _fixed = _dt.datetime(2026, 1, 1, 10, 1, 0, tzinfo=_dt.timezone.utc)

            @classmethod
            def now(cls, tz=None):
                if tz is None:
                    return cls._fixed.replace(tzinfo=None)
                return cls._fixed.astimezone(tz)

        monkeypatch.setattr('src.core.scheduler.datetime', _FixedDT)
        db = tmp_path / 's.db'
        _run('--schedule-db', str(db), '--schedule-add', 'a|* * * * *|python -c pass')
        code, out = _run('--schedule-db', str(db), '--schedule-run-all')
        assert code == 0
        assert 'a:' in out


class TestRunScriptCli:
    def test_run_script_ok(self, tmp_path: Path) -> None:
        script = tmp_path / 's.py'
        script.write_text('x = 1\n', encoding='utf-8')
        code, out = _run('--run-script', str(script))
        assert code == 0
        assert 'successfully' in out

    def test_run_script_rejected(self, tmp_path: Path) -> None:
        script = tmp_path / 's.py'
        script.write_text('import os\n', encoding='utf-8')
        code, out = _run('--run-script', str(script))
        assert code == 2
        assert 'rejected' in out

    def test_run_script_error(self, tmp_path: Path) -> None:
        script = tmp_path / 's.py'
        script.write_text('raise RuntimeError("boom")\n', encoding='utf-8')
        code, out = _run('--run-script', str(script))
        assert code == 1
        assert 'failed' in out
