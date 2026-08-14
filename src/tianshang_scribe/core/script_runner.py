"""Sandboxed execution of user-provided Python scripts.

Used by ``--run-script`` to run small automation scripts (e.g. generating
template data or driving the engines) without giving the script full access to
the host process. Enforcement is layered:

1. :func:`check_imports` statically scans the source with :mod:`ast` and
   rejects imports outside a whitelist plus any dangerous builtins
   (``eval``/``exec``/``__import__``/``open``).
2. Execution runs in a fresh restricted ``globals`` dict (no real
   ``__builtins__`` except a curated safe subset) with a wall-clock timeout.
3. Timeout is enforced via a worker thread; on expiry the worker is detached
   and the run is reported as timed out.

This is a hardening layer, not a security boundary: treat ``--run-script`` as
trusted-but-untrusted input (like CI scripts), not as protection against an
adversary. The subprocess-based scheduler (:mod:`tianshang_scribe.core.scheduler`) remains
the stronger isolation for scheduled work.
"""

from __future__ import annotations

import ast
import builtins
import importlib
import threading
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

ALLOWED_IMPORTS: frozenset[str] = frozenset(
    {
        'json',
        'csv',
        'yaml',
        're',
        'math',
        'time',
        'datetime',
        'pathlib',
        'typing',
        'dataclasses',
        'random',
        'string',
        'itertools',
        'collections',
        'tianshang_scribe.core.word_engine',
        'tianshang_scribe.core.excel_engine',
        'tianshang_scribe.core.ppt_engine',
        'tianshang_scribe.core.document',
        'tianshang_scribe.rendering.template',
        'tianshang_scribe.utils.store',
        'tianshang_scribe.core.scheduler',
    }
)

FORBIDDEN_BUILTINS: frozenset[str] = frozenset(
    {'eval', 'exec', '__import__', 'open', 'compile', 'input', 'breakpoint', 'globals', 'locals'}
)

SAFE_BUILTINS: dict[str, Any] = {
    name: getattr(builtins, name)
    for name in (
        'len',
        'range',
        'str',
        'int',
        'float',
        'bool',
        'list',
        'dict',
        'tuple',
        'set',
        'frozenset',
        'abs',
        'min',
        'max',
        'sum',
        'sorted',
        'reversed',
        'enumerate',
        'zip',
        'filter',
        'map',
        'isinstance',
        'issubclass',
        'getattr',
        'hasattr',
        'setattr',
        'type',
        'repr',
        'print',
        'Exception',
        'ValueError',
        'TypeError',
        'RuntimeError',
        'KeyError',
        'IndexError',
        'NotImplementedError',
        'True',
        'False',
        'None',
        'bytes',
        'bytearray',
        'chr',
        'ord',
        'hex',
        'oct',
        'bin',
        'round',
        'divmod',
        'pow',
        'format',
        'hash',
    )
}


@dataclass
class ScriptResult:
    """Outcome of running a sandboxed script."""

    ok: bool
    output: str = ''
    error: str | None = None
    timed_out: bool = False
    violations: list[str] = field(default_factory=list)


class ScriptRunner:
    """Run a Python source string under a restricted environment + timeout."""

    def __init__(
        self,
        allowed_imports: frozenset[str] | set[str] = ALLOWED_IMPORTS,
        default_timeout: float = 30.0,
    ) -> None:
        """Create a runner with the given import whitelist and default timeout."""
        self.allowed_imports = frozenset(allowed_imports)
        self.default_timeout = default_timeout

    def check_imports(self, source: str) -> list[str]:
        """Return a list of violated import/builtin rules (empty when clean)."""
        violations: list[str] = []
        try:
            tree = ast.parse(source)
        except SyntaxError as e:
            return [f'syntax error: {e}']
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    root = alias.name.split('.')[0]
                    if root not in self.allowed_imports:
                        violations.append(f'import {alias.name} (module {root!r} not allowed)')
            elif isinstance(node, ast.ImportFrom):
                root = (node.module or '').split('.')[0]
                if root not in self.allowed_imports:
                    violations.append(
                        f'from {node.module} import ... (module {root!r} not allowed)'
                    )
            elif isinstance(node, ast.Call):
                fn = node.func
                if isinstance(fn, ast.Name) and fn.id in FORBIDDEN_BUILTINS:
                    violations.append(f'use of forbidden builtin {fn.id!r}')
                elif isinstance(fn, ast.Attribute) and fn.attr in FORBIDDEN_BUILTINS:
                    violations.append(f'use of forbidden attribute {fn.attr!r}')
        return violations

    def _make_safe_import(self) -> Callable[..., Any]:
        """Return a guarded ``__import__`` that only resolves whitelisted roots."""

        def _safe_import(name: str, *args: Any, **kwargs: Any) -> Any:
            if not name or name.split('.')[0] not in self.allowed_imports:
                raise ImportError(f'{name!r} is not in the sandbox import whitelist')
            return importlib.import_module(name)

        return _safe_import

    def run(
        self,
        source: str,
        *,
        timeout: float | None = None,
        extra_globals: dict[str, Any] | None = None,
    ) -> ScriptResult:
        """Execute ``source`` and return a :class:`ScriptResult`."""
        violations = self.check_imports(source)
        if violations:
            return ScriptResult(ok=False, violations=violations)

        safe_globals: dict[str, Any] = {'__builtins__': dict(SAFE_BUILTINS)}
        if extra_globals:
            safe_globals.update(extra_globals)
        safe_globals['__builtins__']['__import__'] = self._make_safe_import()

        result_holder: dict[str, Any] = {}
        thread = threading.Thread(
            target=_exec_target, args=(source, safe_globals, result_holder), daemon=True
        )
        thread.start()
        thread.join(timeout if timeout is not None else self.default_timeout)

        if thread.is_alive():
            return ScriptResult(ok=False, timed_out=True)
        if 'error' in result_holder:
            return ScriptResult(ok=False, error=result_holder['error'])
        return ScriptResult(ok=True)

    def run_file(
        self,
        path: str | Path,
        *,
        timeout: float | None = None,
        extra_globals: dict[str, Any] | None = None,
    ) -> ScriptResult:
        """Read a script file and execute it sandboxed."""
        return self.run(
            Path(path).read_text(encoding='utf-8-sig'),
            timeout=timeout,
            extra_globals=extra_globals,
        )


def _exec_target(source: str, globals_dict: dict[str, Any], holder: dict[str, Any]) -> None:
    try:
        exec(compile(source, '<script>', 'exec'), globals_dict)  # noqa: S102  # guarded by ScriptRunner
    except Exception as e:  # any script error is reported to the caller
        holder['error'] = f'{type(e).__name__}: {e}'
