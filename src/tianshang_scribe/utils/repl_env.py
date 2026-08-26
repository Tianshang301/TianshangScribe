"""REPL environment files (rc): latex flag, command aliases, startup commands.

Search order: explicit ``--env-file`` wins; otherwise the project-level
``<project>/.scribe/repl.rc`` is merged on top of the user-level
``~/.tianshang-scribe/repl.rc``.
"""

from __future__ import annotations

import configparser
from dataclasses import dataclass, field
from pathlib import Path

USER_RC_PATH = Path.home() / '.tianshang-scribe' / 'repl.rc'
PROJECT_RC_NAME = Path('.scribe') / 'repl.rc'


@dataclass
class ReplEnvironment:
    """Resolved REPL environment: latex flag, aliases and startup commands."""

    latex_style: bool = False
    aliases: dict[str, str] = field(default_factory=dict)
    startup_commands: list[str] = field(default_factory=list)


def load_repl_env(
    explicit_file: str | Path | None = None,
    project_dir: str | Path | None = None,
) -> tuple[ReplEnvironment, list[str]]:
    """Load REPL rc files; return the merged env plus human-readable warnings.

    With ``explicit_file`` only that file is loaded. Otherwise the user rc is
    loaded first and the project rc overrides it (aliases) or extends it
    (startup commands run in both files' order). Missing files are silently
    skipped; corrupt or unreadable files produce a warning entry instead.
    """
    env = ReplEnvironment()
    warnings: list[str] = []

    if explicit_file is not None:
        _merge_file(env, warnings, Path(explicit_file).expanduser(), f'--env-file {explicit_file}')
        return env, warnings

    _merge_file(env, warnings, USER_RC_PATH, str(USER_RC_PATH))
    if project_dir is not None:
        project_rc = Path(project_dir) / PROJECT_RC_NAME
        _merge_file(env, warnings, project_rc, str(project_rc))
    return env, warnings


def _merge_file(env: ReplEnvironment, warnings: list[str], path: Path, label: str) -> None:
    """Merge one rc file into ``env``, appending to ``warnings`` on failure."""
    if not path.exists():
        return
    # ConfigParser.read silently ignores files it cannot open, so guard here.
    if not path.is_file():
        warnings.append(f'skipping REPL config {label}: not a regular file')
        return

    parser = configparser.ConfigParser(interpolation=None)
    try:
        parser.read(path, encoding='utf-8')
    except (configparser.Error, UnicodeDecodeError, OSError) as e:
        warnings.append(f'skipping REPL config {label}: {e}')
        return

    if parser.has_section('repl'):
        _merge_repl_section(env, warnings, parser, label)
    if parser.has_section('aliases'):
        for name, target in parser.items('aliases'):
            target = target.strip()
            if not target:
                warnings.append(f'{label}: ignoring empty alias {name!r}')
                continue
            # optionxform lowercases keys, matching the REPL cmd.lower() dispatch.
            env.aliases[name] = target
    if parser.has_section('startup'):
        for _key, value in parser.items('startup'):
            for line in value.splitlines():
                command = line.strip()
                if command:
                    env.startup_commands.append(command)


def _merge_repl_section(
    env: ReplEnvironment,
    warnings: list[str],
    parser: configparser.ConfigParser,
    label: str,
) -> None:
    """Apply the ``[repl]`` section: only ``latex_style`` is recognised."""
    raw = parser.get('repl', 'latex_style', fallback=None)
    if raw is not None:
        try:
            env.latex_style = parser.getboolean('repl', 'latex_style')
        except ValueError:
            warnings.append(f'{label}: invalid [repl] latex_style value {raw!r}')
    for key in parser['repl']:
        if key != 'latex_style':
            warnings.append(f'{label}: ignoring unknown [repl] key {key!r}')
