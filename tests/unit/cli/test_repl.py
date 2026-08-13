"""Tests for the interactive document session (``open`` subcommand REPL)."""

from __future__ import annotations

from pathlib import Path

import pytest
from rich.console import Console

from src.cli.repl import InteractiveSession
from src.core.excel_engine import ExcelEngine
from src.core.word_engine import WordEngine


class TestInteractiveSession:
    def _session(self, engine, path: Path) -> InteractiveSession:
        return InteractiveSession(engine, path, Console(file=None, no_color=True))

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


def _read_text(path: Path) -> str:
    from src.core.word_engine import WordEngine

    e = WordEngine()
    e.open(str(path))
    return e.extract_text()
