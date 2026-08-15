"""Doc-sync guard: every CLI option in ``main.py`` must be documented.

``AGENTS.md`` is the project's single source of truth and the root + zh-CN
READMEs are the user-facing reference. This test parses ``cli/main.py`` with
``ast`` to extract every long ``typer.Option(...)`` flag and asserts each one
appears in all three documents, so adding a CLI flag without updating the
docs fails CI.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

import tianshang_scribe.cli.main as cli

REPO_ROOT = Path(__file__).resolve().parents[3]
MAIN_PY = REPO_ROOT / 'src' / 'tianshang_scribe' / 'cli' / 'main.py'
DOCS = {
    'AGENTS.md': REPO_ROOT / 'AGENTS.md',
    'README.md': REPO_ROOT / 'README.md',
    'README.zh-CN.md': REPO_ROOT / 'readme' / 'README.zh-CN.md',
}


def _cli_long_options() -> set[str]:
    """Return the set of long options (``--foo``) declared in ``main.py``."""
    tree = ast.parse(MAIN_PY.read_text(encoding='utf-8'))
    options: set[str] = set()
    for node in ast.walk(tree):
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Attribute)
            and node.func.attr == 'Option'
        ):
            for arg in node.args:
                if (
                    isinstance(arg, ast.Constant)
                    and isinstance(arg.value, str)
                    and arg.value.startswith('--')
                ):
                    options.add(arg.value)
    return options


@pytest.mark.parametrize('doc_name', sorted(DOCS))
def test_every_cli_option_is_documented(doc_name: str) -> None:
    text = DOCS[doc_name].read_text(encoding='utf-8')
    missing = sorted(opt for opt in _cli_long_options() if opt not in text)
    assert not missing, (
        f'{doc_name} is missing documentation for CLI option(s): {missing}. '
        'Update the options table in the document.'
    )


def test_shared_app_option_parity() -> None:
    """The REPL ``open`` app reuses the same option help strings as the CLI.

    Both apps draw ``--latex-style`` and ``-w/-e/-p`` from module-level
    constants; AGENTS.md must keep its ``open`` (REPL) section in sync.
    """
    agents = REPO_ROOT / 'AGENTS.md'
    text = agents.read_text(encoding='utf-8')
    assert '--latex-style' in text
    assert cli.LATEX_STYLE_HELP == 'Enable LaTeX style markup parsing'
    assert cli.WORD_HELP == 'Process Word document'
    assert cli.open_app is not None
