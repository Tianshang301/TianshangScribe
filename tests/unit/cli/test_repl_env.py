"""Tests for REPL environment (rc) file loading."""

from __future__ import annotations

from pathlib import Path

import pytest

from tianshang_scribe.utils import repl_env
from tianshang_scribe.utils.repl_env import ReplEnvironment, load_repl_env


def _write_rc(directory: Path, text: str, name: Path | None = None) -> Path:
    rc = directory / (name or Path('.scribe') / 'repl.rc')
    rc.parent.mkdir(parents=True, exist_ok=True)
    rc.write_text(text, encoding='utf-8')
    return rc


class TestLoadReplEnv:
    def test_defaults_when_no_files(self, tmp_path: Path, monkeypatch) -> None:
        monkeypatch.setattr(repl_env, 'USER_RC_PATH', tmp_path / 'nonexistent' / 'repl.rc')
        env, warnings = load_repl_env(project_dir=tmp_path / 'proj')
        assert env == ReplEnvironment()
        assert warnings == []

    def test_user_rc_fully_parsed(self, tmp_path: Path, monkeypatch) -> None:
        user_rc = _write_rc(
            tmp_path / 'home',
            '[repl]\nlatex_style = true\n\n[aliases]\nh1 = heading 1\n'
            '\n[startup]\ncommands = style font=Arial\n    heading Title\n',
            name=Path('.tianshang-scribe') / 'repl.rc',
        )
        monkeypatch.setattr(repl_env, 'USER_RC_PATH', user_rc)
        env, warnings = load_repl_env(project_dir=tmp_path / 'proj')
        assert warnings == []
        assert env.latex_style is True
        assert env.aliases == {'h1': 'heading 1'}
        assert env.startup_commands == ['style font=Arial', 'heading Title']

    @pytest.mark.parametrize('value', ['true', 'false', '1', '0', 'yes', 'no', 'on', 'off'])
    def test_boolean_forms(self, tmp_path: Path, value: str) -> None:
        rc = _write_rc(tmp_path / 'rc', f'[repl]\nlatex_style = {value}\n')
        env, _ = load_repl_env(explicit_file=rc)
        expected = value in ('true', '1', 'yes', 'on')
        assert env.latex_style is expected

    def test_project_overrides_user(self, tmp_path: Path, monkeypatch) -> None:
        user_rc = _write_rc(
            tmp_path / 'home',
            '[repl]\nlatex_style = false\n\n[aliases]\nfoo = bar\n'
            '\n[startup]\ncommands = add first\n',
            name=Path('.tianshang-scribe') / 'repl.rc',
        )
        proj_dir = tmp_path / 'proj'
        _write_rc(
            proj_dir,
            '[repl]\nlatex_style = true\n\n[aliases]\nfoo = baz\n'
            '\n[startup]\ncommands = add second\n',
        )
        monkeypatch.setattr(repl_env, 'USER_RC_PATH', user_rc)
        env, _ = load_repl_env(project_dir=proj_dir)
        assert env.aliases['foo'] == 'baz'  # project wins
        assert env.latex_style is True  # last True wins
        assert env.startup_commands == ['add first', 'add second']  # both, user order first

    def test_explicit_file_skips_default_locations(self, tmp_path: Path, monkeypatch) -> None:
        user_rc = _write_rc(
            tmp_path / 'home',
            '[aliases]\nfoo = bar\n',
            name=Path('.tianshang-scribe') / 'repl.rc',
        )
        proj_dir = tmp_path / 'proj'
        _write_rc(proj_dir, '[aliases]\nfoo = bar2\n')
        explicit = _write_rc(tmp_path / 'other', '[aliases]\nbaz = qux\n')
        monkeypatch.setattr(repl_env, 'USER_RC_PATH', user_rc)
        env, _ = load_repl_env(explicit_file=explicit, project_dir=proj_dir)
        assert env.aliases == {'baz': 'qux'}

    def test_corrupt_file_warns_and_other_still_loads(self, tmp_path: Path, monkeypatch) -> None:
        corrupt = tmp_path / 'corrupt.rc'
        corrupt.write_text('not an ini header\nkey = value\n', encoding='utf-8')
        monkeypatch.setattr(repl_env, 'USER_RC_PATH', corrupt)
        env, warnings = load_repl_env()
        assert len(warnings) == 1
        assert 'skipping' in warnings[0]
        assert env == ReplEnvironment()  # corrupt user file discarded wholesale

        # a valid project file still loads alongside the warning
        proj_dir = tmp_path / 'proj'
        _write_rc(proj_dir, '[aliases]\nfoo = baz\n')
        env, warnings = load_repl_env(project_dir=proj_dir)
        assert len(warnings) == 1
        assert env.aliases['foo'] == 'baz'

    def test_directory_as_rc_file_warns(self, tmp_path: Path) -> None:
        rc_dir = tmp_path / 'rcdir'
        rc_dir.mkdir()
        env, warnings = load_repl_env(explicit_file=rc_dir)
        assert env == ReplEnvironment()
        assert len(warnings) == 1
        assert 'skipping' in warnings[0]

    def test_invalid_boolean_warns_keeps_value(self, tmp_path: Path, monkeypatch) -> None:
        user_rc = _write_rc(
            tmp_path / 'home',
            '[repl]\nlatex_style = true\n',
            name=Path('.tianshang-scribe') / 'repl.rc',
        )
        _write_rc(tmp_path / 'proj', '[repl]\nlatex_style = maybe\n')
        monkeypatch.setattr(repl_env, 'USER_RC_PATH', user_rc)
        env, warnings = load_repl_env(project_dir=tmp_path / 'proj')
        assert any('latex_style' in w for w in warnings)
        assert env.latex_style is True  # user value preserved

    def test_unknown_repl_key_warns(self, tmp_path: Path) -> None:
        rc = _write_rc(tmp_path / 'rc', '[repl]\nfont_size = 12\n')
        env, warnings = load_repl_env(explicit_file=rc)
        assert env.latex_style is False
        assert any('font_size' in w for w in warnings)

    def test_empty_alias_skipped_and_startup_blank_lines_dropped(self, tmp_path: Path) -> None:
        rc = _write_rc(
            tmp_path / 'rc',
            '[aliases]\nred = \nblue = style color=0000FF\n'
            '\n[startup]\ncommands = add one\n\n    \n    add two\n',
        )
        env, warnings = load_repl_env(explicit_file=rc)
        assert env.aliases == {'blue': 'style color=0000FF'}
        assert any('empty alias' in w for w in warnings)
        assert env.startup_commands == ['add one', 'add two']

    def test_percent_in_values(self, tmp_path: Path) -> None:
        rc = _write_rc(tmp_path / 'rc', '[aliases]\npct = replace 50% off\n')
        env, warnings = load_repl_env(explicit_file=rc)
        assert warnings == []
        assert env.aliases['pct'] == 'replace 50% off'

    def test_alias_keys_lowercased(self, tmp_path: Path) -> None:
        rc = _write_rc(tmp_path / 'rc', '[aliases]\nGS = get structure\n')
        env, _ = load_repl_env(explicit_file=rc)
        assert env.aliases['gs'] == 'get structure'

    def test_numbered_startup_keys_in_order(self, tmp_path: Path) -> None:
        rc = _write_rc(tmp_path / 'rc', '[startup]\n1 = add one\n2 = add two\n3 = add three\n')
        env, _ = load_repl_env(explicit_file=rc)
        assert env.startup_commands == ['add one', 'add two', 'add three']

    def test_project_dir_none_skips_project_lookup(self, tmp_path: Path, monkeypatch) -> None:
        user_rc = _write_rc(
            tmp_path / 'home',
            '[aliases]\nfoo = bar\n',
            name=Path('.tianshang-scribe') / 'repl.rc',
        )
        proj_dir = tmp_path / 'proj'
        _write_rc(proj_dir, '[aliases]\nfoo = baz\n')
        monkeypatch.setattr(repl_env, 'USER_RC_PATH', user_rc)
        env, warnings = load_repl_env(project_dir=None)
        assert env.aliases == {'foo': 'bar'}
        assert warnings == []

    def test_explicit_file_tilde_expands(self, tmp_path: Path, monkeypatch) -> None:
        home = tmp_path / 'fakehome'
        _write_rc(home, '[aliases]\ntilde = works\n', name=Path('repl.rc'))
        monkeypatch.setenv('HOME', str(home))
        monkeypatch.setenv('USERPROFILE', str(home))
        env, warnings = load_repl_env(explicit_file='~/repl.rc')
        assert warnings == []
        assert env.aliases['tilde'] == 'works'
