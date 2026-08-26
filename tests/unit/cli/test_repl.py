"""Tests for the interactive document session (``open`` subcommand REPL)."""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from rich.console import Console

from tianshang_scribe.cli.repl import InteractiveSession
from tianshang_scribe.core.excel_engine import ExcelEngine
from tianshang_scribe.core.word_engine import WordEngine
from tianshang_scribe.utils.repl_env import ReplEnvironment


class TestInteractiveSession:
    def _session(
        self,
        engine,
        path: Path,
        latex_style: bool = False,
        env: ReplEnvironment | None = None,
        file=None,
    ) -> InteractiveSession:
        return InteractiveSession(
            engine, path, Console(file=file, no_color=True, width=400), latex_style, env=env
        )

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
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: 'y')
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
            'tianshang_scribe.cli.repl.Prompt.ask',
            lambda *a, **k: (_ for _ in ()).throw(EOFError()),
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
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: 'n')
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
        assert s.path == new_path.resolve()

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
        monkeypatch.chdir(tmp_path)
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        answers = iter(['add "hello"', 'quit', 'n'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()
        assert 'hello' in w.extract_text()

    def test_run_loop_eof(self, tmp_path: Path, monkeypatch) -> None:
        monkeypatch.chdir(tmp_path)
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        monkeypatch.setattr(
            'tianshang_scribe.cli.repl.Prompt.ask',
            lambda *a, **k: (_ for _ in ()).throw(EOFError()),
        )
        s.run()

    def test_run_loop_value_error_continues(self, tmp_path: Path, monkeypatch) -> None:
        monkeypatch.chdir(tmp_path)
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        answers = iter(['add', 'quit'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()

    def test_run_loop_generic_error_continues(self, tmp_path: Path, monkeypatch) -> None:
        monkeypatch.chdir(tmp_path)
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        answers = iter(['table "A|B"', 'quit', 'n'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()

    def test_run_loop_blank_line(self, tmp_path: Path, monkeypatch) -> None:
        monkeypatch.chdir(tmp_path)
        w = WordEngine()
        w.create()
        target = tmp_path / 't.docx'
        s = self._session(w, target)
        answers = iter(['', '  ', 'quit'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
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
        from tianshang_scribe.cli.repl import _split_tokens

        assert _split_tokens('add "hello world"') == ['add', 'hello world']
        assert _split_tokens(r'math \frac{a}{b}') == ['math', r'\frac{a}{b}']

    # -- document-directory behavior ---------------------------------------

    def _reopened(self, target: Path, seed: str = 'seed') -> WordEngine:
        w = WordEngine()
        w.create()
        w.add_text(seed)
        w.save(str(target))
        w = WordEngine()
        w.open(str(target))
        return w

    def test_init_resolves_relative_path(self, tmp_path: Path, monkeypatch) -> None:
        work = tmp_path / 'work'
        work.mkdir()
        monkeypatch.chdir(tmp_path)
        s = self._session(WordEngine(), Path('work/t.docx'))
        assert s.path == (tmp_path / 'work' / 't.docx').resolve()
        assert s._doc_dir == s.path.parent
        assert s._original_cwd == Path.cwd()  # __init__ never chdirs

    def test_relative_launch_save_lands_in_original_location(
        self, tmp_path: Path, monkeypatch
    ) -> None:
        root = tmp_path
        work = root / 'work'
        work.mkdir()
        target = work / 'report.docx'
        w = self._reopened(target)
        monkeypatch.chdir(root)
        s = self._session(w, Path('work/report.docx'))
        assert s.path == target.resolve()
        answers = iter(['add "x"', 'save', 'quit'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()
        assert 'x' in _read_text(target)
        assert not (work / 'work' / 'report.docx').exists()  # regression: nested duplicate
        assert Path.cwd() == root.resolve()

    def test_chdir_enables_doc_relative_csv(self, tmp_path: Path, monkeypatch) -> None:
        root = tmp_path
        work = root / 'work'
        work.mkdir()
        (work / 'd.csv').write_text('x,y\n1,2\n', encoding='utf-8')
        target = work / 'doc.docx'
        w = self._reopened(target)
        monkeypatch.chdir(root)
        s = self._session(w, Path('work/doc.docx'))
        answers = iter(['table @d.csv', 'quit', 'n'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()
        assert w.extract_tables() == [[['x', 'y'], ['1', '2']]]
        assert Path.cwd() == root.resolve()

    def test_run_restores_cwd_on_eof(self, tmp_path: Path, monkeypatch) -> None:
        root = tmp_path
        work = root / 'work'
        work.mkdir()
        w = self._reopened(work / 't.docx')
        monkeypatch.chdir(root)
        s = self._session(w, Path('work/t.docx'))
        monkeypatch.setattr(
            'tianshang_scribe.cli.repl.Prompt.ask',
            lambda *a, **k: (_ for _ in ()).throw(EOFError()),
        )
        s.run()
        assert Path.cwd() == root.resolve()

    def test_enter_doc_dir_oserror_warns_and_continues(self, tmp_path: Path, monkeypatch) -> None:
        work = tmp_path / 'work'
        work.mkdir()
        monkeypatch.chdir(tmp_path)
        w = WordEngine()
        w.create()
        sink = io.StringIO()
        s = self._session(w, work / 't.docx', file=sink)

        def broken_chdir(path: str) -> None:
            raise OSError('denied')

        monkeypatch.setattr('tianshang_scribe.cli.repl.os.chdir', broken_chdir)
        answers = iter(['quit'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()
        assert 'could not enter' in sink.getvalue()
        assert 'Opened' in sink.getvalue()  # banner still shown
        assert Path.cwd() == tmp_path.resolve()  # cwd untouched

    def test_restore_cwd_oserror_swallowed(self, tmp_path: Path, monkeypatch) -> None:
        work = tmp_path / 'work'
        work.mkdir()
        monkeypatch.chdir(tmp_path)
        w = WordEngine()
        w.create()
        s = self._session(w, work / 't.docx')

        import os as os_module

        real_chdir = os_module.chdir

        def selective_chdir(path) -> None:
            if Path(path) == s._doc_dir:
                real_chdir(path)  # entering succeeds
            else:
                raise OSError('cannot restore')

        monkeypatch.setattr('tianshang_scribe.cli.repl.os.chdir', selective_chdir)
        answers = iter(['quit'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()  # must not raise despite failed restore

    def test_banner_working_directory_line(self, tmp_path: Path, monkeypatch) -> None:
        work = tmp_path / 'work'
        work.mkdir()
        monkeypatch.chdir(tmp_path)
        w = WordEngine()
        w.create()
        sink = io.StringIO()
        s = self._session(w, work / 't.docx', file=sink)
        answers = iter(['quit'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()
        assert f'Working directory: {work.resolve()}' in sink.getvalue()

        same = tmp_path / 'flat.docx'
        sink2 = io.StringIO()
        s2 = self._session(WordEngine(), same, file=sink2)
        answers2 = iter(['quit'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers2))
        s2.run()
        assert 'Working directory:' not in sink2.getvalue()

    def test_path_resolves_against_doc_dir(self, tmp_path: Path, monkeypatch) -> None:
        work = tmp_path / 'work'
        work.mkdir()
        monkeypatch.chdir(tmp_path)
        s = self._session(WordEngine(), work / 't.docx')
        assert s.execute('path out.docx') is True
        assert s.path == (work / 'out.docx').resolve()

    def test_save_with_relative_path_anchors_doc_dir(self, tmp_path: Path, monkeypatch) -> None:
        work = tmp_path / 'work'
        work.mkdir()
        other = tmp_path / 'other'
        other.mkdir()
        target = work / 't.docx'
        w = self._reopened(target)
        monkeypatch.chdir(other)  # process cwd deliberately different from doc dir
        s = self._session(w, target)
        assert s.execute('save rel.docx') is True
        assert (work / 'rel.docx').exists()
        assert s.path == (work / 'rel.docx').resolve()

    # -- startup commands --------------------------------------------------

    def test_startup_commands_run_before_loop(self, tmp_path: Path, monkeypatch) -> None:
        env = ReplEnvironment(startup_commands=['heading 1 Summary', 'style font=Arial,size=14'])
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx', env=env)
        answers = iter(['quit', 'n'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()
        assert 'Summary' in w.extract_text()
        assert w._base_style.font_name == 'Arial'

    def test_startup_quit_ends_session_without_loop(self, tmp_path: Path, monkeypatch) -> None:
        env = ReplEnvironment(startup_commands=['quit', 'add should_not'])
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx', env=env)
        prompted: list[str] = []
        monkeypatch.setattr(
            'tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: prompted.append('asked')
        )
        s.run()
        assert prompted == []  # interactive loop never reached
        assert 'should_not' not in w.extract_text()

    def test_startup_failure_continues_with_next(self, tmp_path: Path, monkeypatch) -> None:
        env = ReplEnvironment(startup_commands=['nosuchcmd', 'add ok'])
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx', env=env)
        answers = iter(['quit', 'n'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()
        assert 'ok' in w.extract_text()

    def test_startup_commands_echoed(self, tmp_path: Path, monkeypatch) -> None:
        env = ReplEnvironment(startup_commands=['add hi'])
        w = WordEngine()
        w.create()
        sink = io.StringIO()
        s = self._session(w, tmp_path / 't.docx', env=env, file=sink)
        answers = iter(['quit', 'n'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()
        assert '> add hi' in sink.getvalue()

    def test_startup_keyboard_interrupt_enters_loop(self, tmp_path: Path, monkeypatch) -> None:
        env = ReplEnvironment(startup_commands=['boom', 'add skipped'])
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx', env=env)
        real_execute = s.execute

        def flaky_execute(line: str) -> bool:
            if line == 'boom':
                raise KeyboardInterrupt()
            return real_execute(line)

        monkeypatch.setattr(s, 'execute', flaky_execute)
        answers = iter(['quit'])
        monkeypatch.setattr('tianshang_scribe.cli.repl.Prompt.ask', lambda *a, **k: next(answers))
        s.run()
        assert 'skipped' not in w.extract_text()  # remaining startup dropped

    # -- aliases and env command -------------------------------------------

    def test_alias_define_and_use(self, tmp_path: Path) -> None:
        env = ReplEnvironment()
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx', env=env)
        assert s.execute('env alias st style font=Arial,size=14') is True
        assert s.execute('st') is True
        assert w._base_style.font_name == 'Arial'
        assert env.aliases['st'] == 'style font=Arial,size=14'

    def test_alias_args_appended(self, tmp_path: Path) -> None:
        env = ReplEnvironment(aliases={'h2': 'heading 2'})
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx', env=env)
        assert s.execute('h2 Deep Title') is True
        assert 'Deep Title' in w.extract_text()

    def test_alias_expansion_is_single_level(self, tmp_path: Path) -> None:
        env = ReplEnvironment(aliases={'a': 'x', 'x': 'add hi'})
        w = WordEngine()
        w.create()
        sink = io.StringIO()
        s = self._session(w, tmp_path / 't.docx', env=env, file=sink)
        assert s.execute('a') is True  # expands to 'x', which is NOT re-expanded
        assert 'Unknown command' in sink.getvalue()
        assert '"x"' in sink.getvalue()
        assert 'hi' not in w.extract_text()

    def test_alias_cannot_shadow_env(self, tmp_path: Path) -> None:
        env = ReplEnvironment()
        w = WordEngine()
        w.create()
        sink = io.StringIO()
        s = self._session(w, tmp_path / 't.docx', env=env, file=sink)
        assert s.execute('env alias env help') is True
        assert 'env' not in env.aliases
        assert 'Cannot shadow' in sink.getvalue()

    def test_alias_shadowing_builtin_warns_but_works(self, tmp_path: Path) -> None:
        env = ReplEnvironment()
        w = WordEngine()
        w.create()
        sink = io.StringIO()
        s = self._session(w, tmp_path / 't.docx', env=env, file=sink)
        assert s.execute('env alias info add tagged') is True
        assert 'shadows a built-in command' in sink.getvalue()
        assert s.execute('info') is True  # alias takes precedence over the builtin
        assert 'tagged' in w.extract_text()

    def test_alias_missing_target_raises(self, tmp_path: Path) -> None:
        s = self._session(WordEngine(), tmp_path / 't.docx')
        with pytest.raises(ValueError):
            s.execute('env alias foo')

    def test_unalias_removes_and_warns_missing(self, tmp_path: Path) -> None:
        env = ReplEnvironment(aliases={'gone': 'add x'})
        w = WordEngine()
        w.create()
        sink = io.StringIO()
        s = self._session(w, tmp_path / 't.docx', env=env, file=sink)
        assert s.execute('env unalias gone') is True
        assert 'Alias removed' in sink.getvalue()
        assert 'gone' not in env.aliases
        assert s.execute('env unalias gone') is True
        assert 'No such alias' in sink.getvalue()

    def test_env_show_lists_state(self, tmp_path: Path) -> None:
        env = ReplEnvironment(aliases={'b': 'bold'}, startup_commands=['add 1'])
        sink = io.StringIO()
        s = self._session(WordEngine(), tmp_path / 't.docx', latex_style=True, env=env, file=sink)
        assert s.execute('env') is True
        out = sink.getvalue()
        assert 'latex_style: True' in out
        assert 'startup commands: 1' in out
        assert 'b = bold' in out

    def test_alias_target_with_nested_quotes(self, tmp_path: Path) -> None:
        env = ReplEnvironment()
        w = WordEngine()
        w.create()
        s = self._session(w, tmp_path / 't.docx', env=env)
        assert s.execute('env alias g "add \'two words\'"') is True
        assert s.execute('g') is True
        assert 'two words' in w.extract_text()


def _read_text(path: Path) -> str:
    from tianshang_scribe.core.word_engine import WordEngine

    e = WordEngine()
    e.open(str(path))
    return e.extract_text()
