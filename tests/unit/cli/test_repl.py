"""Tests for the interactive document session (``open`` subcommand REPL)."""

from __future__ import annotations

from pathlib import Path

import pytest
from rich.console import Console

from src.cli.repl import InteractiveSession
from src.core.excel_engine import ExcelEngine
from src.core.word_engine import WordEngine


class TestInteractiveSession:
    def _session(self, engine, path: Path, latex_style: bool = False) -> InteractiveSession:
        return InteractiveSession(engine, path, Console(file=None, no_color=True), latex_style)

    def test_unknown_command(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('nosuchcmd') is True

    def test_add_and_extract(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('add "Hello World"') is True
        assert s.dirty is True
        assert 'Hello World' in w.extract_text()

    def test_heading(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        s.execute('heading 2 "Section"')
        assert 'Section' in w.extract_text()

    def test_table(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('table "A,B|1,2"') is True
        assert w.extract_tables() == [[['A', 'B'], ['1', '2']]]

    def test_table_from_csv(self, tmp_path: Path) -> None:
        csv_file = tmp_path / 'd.csv'
        csv_file.write_text('x,y\n1,2\n', encoding='utf-8')
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        s.execute(f'table @{csv_file}')
        assert w.extract_tables() == [[['x', 'y'], ['1', '2']]]

    def test_table_not_word_raises(self, tmp_path: Path) -> None:
        e = ExcelEngine()
        e.create()
        s = self._session(e, tmp_path / 't.xlsx')
        with pytest.raises(NotImplementedError):
            s.execute('table "A|B"')

    def test_replace_and_delete(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        w.add_text('foo bar')
        s = self._session(w, tmp_path / 't.docx')
        s.execute('replace foo baz')
        assert 'baz bar' in w.extract_text()
        s.execute('delete bar')
        assert 'baz' in w.extract_text()
        assert 'bar' not in w.extract_text()

    def test_extract_text_returns(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        w.add_text('alpha')
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('extract text') is True

    def test_save_clears_dirty(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        w.add_text('x')
        target = tmp_path / 't.docx'
        w.save(str(target))
        s = self._session(w, target)
        s.execute('add "y"')
        assert s.dirty is True
        s.execute('save')
        assert s.dirty is False
        assert 'y' in _read_text(target)

    def test_quit_clean_no_prompt(self, tmp_path: Path, monkeypatch) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('quit') is False

    def test_quit_dirty_prompts_and_saves(self, tmp_path: Path, monkeypatch) -> None:
        w = WordEngine()
        w.create()
        w.add_text('orig')
        target = tmp_path / 't.docx'
        w.save(str(target))
        w = WordEngine()
        w.open(str(target))
        s = self._session(w, target)
        s.execute('add "new"')
        monkeypatch.setattr('src.cli.repl.Prompt.ask', lambda *a, **k: 'y')
        assert s.execute('quit') is False
        assert 'new' in _read_text(target)

    def test_quit_dirty_eof_no_save(self, tmp_path: Path, monkeypatch) -> None:
        w = WordEngine()
        w.create()
        w.add_text('orig')
        target = tmp_path / 't.docx'
        w.save(str(target))
        w = WordEngine()
        w.open(str(target))
        s = self._session(w, target)
        s.execute('add "new"')
        monkeypatch.setattr(
            'src.cli.repl.Prompt.ask', lambda *a, **k: (_ for _ in ()).throw(EOFError())
        )
        assert s.execute('quit') is False
        assert 'new' not in _read_text(target)

    def test_quit_dirty_answer_no(self, tmp_path: Path, monkeypatch) -> None:
        w = WordEngine()
        w.create()
        w.add_text('orig')
        target = tmp_path / 't.docx'
        w.save(str(target))
        w = WordEngine()
        w.open(str(target))
        s = self._session(w, target)
        s.execute('add "new"')
        monkeypatch.setattr('src.cli.repl.Prompt.ask', lambda *a, **k: 'n')
        assert s.execute('quit') is False
        assert 'new' not in _read_text(target)

    def test_help(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('help') is True
        assert s.execute('?') is True

    def test_math_word(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('math x^2') is True
        assert s.dirty is True

    def test_math_empty(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        with pytest.raises(ValueError):
            s.execute('math')

    def test_math_non_word(self, tmp_path: Path) -> None:
        e = ExcelEngine()
        e.create()
        s = self._session(e, tmp_path / 't.xlsx')
        with pytest.raises(NotImplementedError):
            s.execute('math x^2')

    def test_style(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('style font=Arial,size=14') is True
        assert w._base_style.font_name == 'Arial'

    def test_style_empty(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        with pytest.raises(ValueError):
            s.execute('style')

    def test_latex_style_add(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx', latex_style=True)
        assert s.execute('add "\\bfseries{bold}"') is True
        assert 'bold' in w.extract_text()

    def test_extract_structure(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('extract structure') is True

    def test_extract_metadata(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('extract metadata') is True

    def test_extract_tables(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        w.add_table_data([['A', 'B'], ['1', '2']])
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('extract tables') is True

    def test_extract_tables_empty(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('extract tables') is True

    def test_extract_unknown_mode(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        with pytest.raises(ValueError):
            s.execute('extract bogus')

    def test_info(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('info') is True

    def test_path_get_set(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        assert s.execute('path') is True
        new_path = tmp_path / 'u.docx'
        assert s.execute(f'path {new_path}') is True
        assert s.path == new_path

    def test_add_empty_raises(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        with pytest.raises(ValueError):
            s.execute('add')

    def test_heading_empty_raises(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        with pytest.raises(ValueError):
            s.execute('heading')

    def test_heading_no_level(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('heading Title') is True

    def test_heading_excel(self, tmp_path: Path) -> None:
        e = ExcelEngine()
        e.create()
        s = self._session(e, tmp_path / 't.xlsx')
        assert s.execute('heading Section') is True

    def test_table_empty_raises(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        with pytest.raises(ValueError):
            s.execute('table')

    def test_replace_too_few_args(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        with pytest.raises(ValueError):
            s.execute('replace onlyone')

    def test_replace_multiword_new(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        w.add_text('foo bar')
        s = self._session(w, tmp_path / 't.docx')
        s.execute('replace foo baz qux')
        assert 'baz qux bar' in w.extract_text()

    def test_delete_empty_raises(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        with pytest.raises(ValueError):
            s.execute('delete')

    def test_save_with_path(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        w.add_text('orig')
        target = tmp_path / 't.docx'
        w.save(str(target))
        w = WordEngine()
        w.open(str(target))
        s = self._session(w, target)
        s.execute('add "more"')
        new_path = tmp_path / 'saved.docx'
        assert s.execute(f'save {new_path}') is True
        assert s.dirty is False
        assert 'more' in _read_text(new_path)

    def test_run_loop_add_then_quit(self, tmp_path: Path, monkeypatch) -> None:
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        answers = iter(['add "hello"', 'quit', 'n'])
        monkeypatch.setattr('src.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()
        assert 'hello' in w.extract_text()

    def test_run_loop_eof(self, tmp_path: Path, monkeypatch) -> None:
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        monkeypatch.setattr(
            'src.cli.repl.Prompt.ask', lambda *a, **k: (_ for _ in ()).throw(EOFError())
        )
        s.run()

    def test_run_loop_value_error_continues(self, tmp_path: Path, monkeypatch) -> None:
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        answers = iter(['add', 'quit'])
        monkeypatch.setattr('src.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()

    def test_run_loop_generic_error_continues(self, tmp_path: Path, monkeypatch) -> None:
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        answers = iter(['table "A|B"', 'quit', 'n'])
        monkeypatch.setattr('src.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()

    def test_run_loop_blank_line(self, tmp_path: Path, monkeypatch) -> None:
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        answers = iter(['', '  ', 'quit'])
        monkeypatch.setattr('src.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()

    def test_execute_empty(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('') is True

    def test_exit_alias(self, tmp_path: Path) -> None:
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx')
        assert s.execute('exit') is False
        assert s.execute('q') is False

    def test_split_tokens_quotes(self) -> None:
        from src.cli.repl import _split_tokens

        assert _split_tokens('add "hello world"') == ['add', 'hello world']
        assert _split_tokens(r'math \frac{a}{b}') == ['math', r'\frac{a}{b}']


def _read_text(path: Path) -> str:
    from src.core.word_engine import WordEngine

    e = WordEngine()
    e.open(str(path))
    return e.extract_text()
