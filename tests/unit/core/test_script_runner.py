"""Unit tests for the sandboxed script runner."""

from __future__ import annotations

from pathlib import Path

from src.core.script_runner import FORBIDDEN_BUILTINS, ScriptRunner


class TestCheckImports:
    def _runner(self) -> ScriptRunner:
        return ScriptRunner()

    def test_clean_source(self) -> None:
        violations = self._runner().check_imports('import json\nx = 1 + 2\n')
        assert violations == []

    def test_forbidden_import(self) -> None:
        violations = self._runner().check_imports('import os\n')
        assert any('os' in v for v in violations)

    def test_from_import_forbidden(self) -> None:
        violations = self._runner().check_imports('from socket import socket\n')
        assert any('socket' in v for v in violations)

    def test_allowed_submodule_import(self) -> None:
        violations = self._runner().check_imports('from pathlib import Path\n')
        assert violations == []

    def test_eval_forbidden(self) -> None:
        violations = self._runner().check_imports('eval("1+1")\n')
        assert any('eval' in v for v in violations)

    def test_exec_forbidden(self) -> None:
        violations = self._runner().check_imports('exec("x = 1")\n')
        assert any('exec' in v for v in violations)

    def test_open_attribute_forbidden(self) -> None:
        violations = self._runner().check_imports('f = open("x")\n')
        assert any('open' in v for v in violations)

    def test_syntax_error_reported(self) -> None:
        violations = self._runner().check_imports('def broken(\n')
        assert violations and 'syntax' in violations[0]

    def test_custom_whitelist(self) -> None:
        runner = ScriptRunner(allowed_imports={'json'})
        assert runner.check_imports('import csv\n')
        assert runner.check_imports('import json\n') == []


class TestRun:
    def _runner(self) -> ScriptRunner:
        return ScriptRunner(default_timeout=5.0)

    def test_ok(self) -> None:
        result = self._runner().run('x = 40 + 2\n')
        assert result.ok is True
        assert result.error is None

    def test_prints_to_stdout_allowed(self) -> None:
        result = self._runner().run('print(chr(72) + chr(105))\n')
        assert result.ok is True

    def test_raises_reports_error(self) -> None:
        result = self._runner().run('raise ValueError(chr(66) + chr(97))\n')
        assert result.ok is False
        assert result.error is not None
        assert 'ValueError' in result.error

    def test_forbidden_import_blocks_execution(self) -> None:
        result = self._runner().run('import subprocess\n')
        assert result.ok is False
        assert result.violations

    def test_timeout(self) -> None:
        result = self._runner().run('import time\nwhile True:\n    time.sleep(1)\n', timeout=0.1)
        assert result.ok is False
        assert result.timed_out is True

    def test_extra_globals_available(self) -> None:
        result = self._runner().run('out = seed * 2\n', extra_globals={'seed': 21})
        assert result.ok is True

    def test_safe_builtins_missing_dangerous(self) -> None:
        result = self._runner().run('__import__("os")\n')
        # __import__ is blocked by the static scanner before execution
        assert result.ok is False
        assert result.violations
        assert any('__import__' in v for v in result.violations)

    def test_run_file(self, tmp_path: Path) -> None:
        script = tmp_path / 's.py'
        script.write_text('a = 1\n', encoding='utf-8')
        result = self._runner().run_file(script)
        assert result.ok is True

    def test_run_file_with_bom(self, tmp_path: Path) -> None:
        script = tmp_path / 's.py'
        script.write_text('\ufeffa = 1\n', encoding='utf-8')
        result = self._runner().run_file(script)
        assert result.ok is True

    def test_whitelisted_import_usable(self) -> None:
        result = self._runner().run('import json\nout = json.dumps(chr(120) + chr(121))\n')
        assert result.ok is True

    def test_runtime_import_rejected_by_guard(self) -> None:
        result = self._runner().run('import os\n')
        assert result.ok is False
        assert result.violations


class TestConstants:
    def test_forbidden_builtins(self) -> None:
        assert 'eval' in FORBIDDEN_BUILTINS
        assert 'exec' in FORBIDDEN_BUILTINS
        assert 'open' in FORBIDDEN_BUILTINS
