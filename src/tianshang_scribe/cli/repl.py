"""Interactive document session (REPL) for the ``open`` subcommand.

``tianshang-scribe open file.docx`` loads the document once and holds it in
memory (main thread); commands mutate it in place and ``save`` persists.
"""

from __future__ import annotations

import contextlib
import csv
import io
import json
import os
import shlex
from collections.abc import Callable
from pathlib import Path
from typing import Any

from rich.console import Console
from rich.prompt import Prompt

from tianshang_scribe.cli.global_opts import parse_table_input
from tianshang_scribe.utils.repl_env import ReplEnvironment

# Keep in sync with the handlers dict in ``execute`` (and the alias hook there).
BUILTIN_COMMANDS: frozenset[str] = frozenset(
    {
        'help',
        '?',
        'quit',
        'exit',
        'q',
        'add',
        'heading',
        'table',
        'math',
        'replace',
        'delete',
        'style',
        'extract',
        'info',
        'path',
        'save',
        'env',
    }
)


class InteractiveSession:
    """REPL session bound to a single opened document engine."""

    def __init__(
        self,
        engine: Any,
        path: str | Path,
        console: Console,
        latex_style: bool = False,
        env: ReplEnvironment | None = None,
    ) -> None:
        """Initialize the session bound to a document engine and path."""
        self.engine = engine
        # Resolve eagerly so saves land in the launch location even after run()
        # chdirs into the document directory.
        self.path = Path(path).expanduser().resolve()
        self.console = console
        self.latex_style = latex_style
        self.dirty = False
        self.env = env if env is not None else ReplEnvironment()
        # Anchored at construction; ``path <new>`` never re-anchors it, keeping
        # "process cwd == relative-path anchor" true for the whole session.
        self._doc_dir = self.path.parent
        # Never chdir here: direct execute() callers (tests) must be unaffected.
        self._original_cwd = Path.cwd()
        self._entered_doc_dir = False

    # -- lifecycle ---------------------------------------------------------

    def run(self) -> None:
        """Enter the interactive loop until the user quits."""
        try:
            # Must precede startup commands (@file.csv resolves against the
            # process cwd); the engine was already opened by main against the
            # original cwd — this ordering is load-bearing.
            self._enter_doc_dir()
            self._print_banner()
            if not self._run_startup_commands():
                return  # e.g. literal 'quit' in rc ends the session
            while True:
                try:
                    line = Prompt.ask(f'[bold cyan]{self.path.name}[/bold cyan]')
                except (EOFError, KeyboardInterrupt):
                    self.console.print()
                    self._quit([])
                    return
                if not line or not line.strip():
                    continue
                try:
                    keep = self.execute(line)
                except KeyboardInterrupt:
                    self.console.print('[dim]interrupted[/dim]')
                    continue
                except (ValueError, NotImplementedError) as e:
                    self.console.print(f'[red]Error:[/red] {e}')
                    continue
                except Exception as e:  # interactive robustness, never crash the session
                    self.console.print(f'[red]Unexpected error:[/red] {e}')
                    continue
                if not keep:
                    return
        finally:
            self._restore_cwd()

    def _print_banner(self) -> None:
        """Print the opening banner and the effective working directory."""
        self.console.print(
            f'[bold]Opened[/bold] {self.path} — '
            'type [cyan]help[/cyan] for commands, [cyan]quit[/cyan] to exit.'
        )
        if self._entered_doc_dir and self._doc_dir != self._original_cwd:
            self.console.print(f'[dim]Working directory: {self._doc_dir}[/dim]')

    def _enter_doc_dir(self) -> None:
        """Change the process working directory to the document's directory."""
        try:
            os.chdir(self._doc_dir)
            self._entered_doc_dir = True
        except OSError as e:
            self.console.print(f'[yellow]Warning:[/yellow] could not enter {self._doc_dir}: {e}')

    def _restore_cwd(self) -> None:
        """Restore the original working directory (best effort)."""
        if not self._entered_doc_dir:
            return
        self._entered_doc_dir = False
        with contextlib.suppress(OSError):
            os.chdir(self._original_cwd)

    def _resolve_doc_path(self, raw: str) -> Path:
        """Resolve a user path relative to the document directory."""
        resolved = Path(raw).expanduser()
        if not resolved.is_absolute():
            resolved = self._doc_dir / resolved
        return resolved.resolve()

    def _run_startup_commands(self) -> bool:
        """Execute rc startup commands; ``False`` when the session should end."""
        for command in self.env.startup_commands:
            if not command.strip():
                continue
            self.console.print(f'[dim]> {command}[/dim]')
            try:
                keep = self.execute(command)
            except KeyboardInterrupt:
                self.console.print('[dim]interrupted[/dim]')
                break
            except Exception as e:  # one bad rc line must not kill the session
                self.console.print(f'[yellow]Startup command failed:[/yellow] {e}')
                continue
            if not keep:
                return False
        return True

    # -- dispatch ----------------------------------------------------------

    def execute(self, line: str) -> bool:
        """Run one command line; return ``False`` when the session should exit."""
        parts = _split_tokens(line)
        if not parts:
            return True
        cmd = parts[0].lower()
        args = parts[1:]
        # Single-level alias expansion: the expanded head is never re-checked
        # against aliases, so cycles are impossible. ``env`` stays exempt so the
        # environment can never be locked out.
        if cmd != 'env' and cmd in self.env.aliases:
            parts = _split_tokens(self.env.aliases[cmd]) + args
            if not parts:
                return True
            cmd = parts[0].lower()
            args = parts[1:]
        handlers: dict[str, Callable[[list[str]], bool]] = {
            'help': self._help,
            '?': self._help,
            'quit': self._quit,
            'exit': self._quit,
            'q': self._quit,
            'add': self._add,
            'heading': self._heading,
            'table': self._table,
            'math': self._math,
            'replace': self._replace,
            'delete': self._delete,
            'style': self._style,
            'extract': self._extract,
            'info': self._info,
            'path': self._path,
            'save': self._save,
            'env': self._env,
        }
        handler = handlers.get(cmd)
        if handler is None:
            self.console.print(f'[red]Unknown command[/red] "{cmd}". Type [cyan]help[/cyan].')
            return True
        return handler(args)

    # -- commands ----------------------------------------------------------

    def _help(self, args: list[str]) -> bool:
        self.console.print(
            'Commands: add <text> | heading [level] <text> | table <inline|@file.csv> | '
            'math <latex> | replace <old> <new> | delete <text> | style key=value,... | '
            'extract <text|tables|structure|metadata> | info | path [path] | save [path] | '
            'env [alias <name> <cmd...> | unalias <name>] | help | quit'
        )
        return True

    def _quit(self, args: list[str]) -> bool:
        if self.dirty:
            try:
                answer = Prompt.ask(
                    'Unsaved changes — save before exit?',
                    choices=['y', 'n'],
                    default='y',
                )
            except (EOFError, KeyboardInterrupt):
                self.console.print('[dim]Exiting without saving.[/dim]')
                return False
            if answer.strip().lower().startswith('y'):
                self._save([])
        return False

    def _add(self, args: list[str]) -> bool:
        text = ' '.join(args)
        if not text:
            raise ValueError('add requires text')
        if self.latex_style and hasattr(self.engine, 'add_latex_content'):
            self.engine.add_latex_content(text)
        else:
            self.engine.add_text(text)
        self.dirty = True
        self.console.print('[green]Added.[/green]')
        return True

    def _heading(self, args: list[str]) -> bool:
        if not args:
            raise ValueError('heading requires text')
        if args[0].isdigit() and len(args) > 1:
            level = int(args[0])
            text = ' '.join(args[1:])
        else:
            level = 1
            text = ' '.join(args)
        if hasattr(self.engine, 'add_heading'):
            self.engine.add_heading(text, level=level)
        else:
            self.engine.add_text(text)
        self.dirty = True
        self.console.print(f'[green]Heading {level}:[/green] {text}')
        return True

    def _table(self, args: list[str]) -> bool:
        if not args:
            raise ValueError('table requires "H1,H2|a1,a2" or @file.csv')
        if not hasattr(self.engine, 'add_table_data'):
            raise NotImplementedError('Tables are only supported for Word documents')
        rows = parse_table_input(args[0])
        self.engine.add_table_data(rows)
        self.dirty = True
        self.console.print(f'[green]Table added:[/green] {len(rows)}x{len(rows[0])}.')
        return True

    def _math(self, args: list[str]) -> bool:
        latex = ' '.join(args)
        if not latex:
            raise ValueError('math requires a LaTeX formula')
        if not hasattr(self.engine, 'add_math_formula'):
            raise NotImplementedError('Math formulas are only supported for Word documents')
        self.engine.add_math_formula(latex)
        self.dirty = True
        self.console.print('[green]Formula added.[/green]')
        return True

    def _replace(self, args: list[str]) -> bool:
        if len(args) < 2:
            raise ValueError('replace requires "old" "new"')
        count = self.engine.replace_text(args[0], ' '.join(args[1:]))
        self.dirty = True
        self.console.print(f'[green]Replaced[/green] {count} occurrence(s).')
        return True

    def _delete(self, args: list[str]) -> bool:
        if not args:
            raise ValueError('delete requires text')
        count = self.engine.replace_text(' '.join(args), '')
        self.dirty = True
        self.console.print(f'[green]Deleted[/green] {count} occurrence(s).')
        return True

    def _style(self, args: list[str]) -> bool:
        style_str = ' '.join(args)
        if not style_str:
            raise ValueError('style requires key=value,...')
        self.engine.set_style(style_str)
        self.dirty = True
        self.console.print(f'[green]Style set:[/green] {style_str}')
        return True

    def _extract(self, args: list[str]) -> bool:
        mode = args[0].lower() if args else 'text'
        if mode == 'text':
            self.console.print(self.engine.extract_text())
        elif mode == 'structure':
            self.console.print(
                json.dumps(self.engine.extract_structure(), ensure_ascii=False, default=str)
            )
        elif mode == 'metadata':
            self.console.print(
                json.dumps(self.engine.get_metadata(), ensure_ascii=False, default=str)
            )
        elif mode == 'tables':
            tables = self.engine.extract_tables()
            if not tables:
                self.console.print('[yellow]No tables found.[/yellow]')
                return True
            buf = io.StringIO()
            writer = csv.writer(buf)
            for table in tables:
                for row in table:
                    writer.writerow(row)
                writer.writerow([])
            self.console.print(buf.getvalue())
        else:
            raise ValueError('extract modes: text, tables, structure, metadata')
        return True

    def _info(self, args: list[str]) -> bool:
        self.console.print(
            json.dumps(self.engine.extract_structure(), ensure_ascii=False, default=str)
        )
        return True

    def _path(self, args: list[str]) -> bool:
        if args:
            self.path = self._resolve_doc_path(args[0])
            self.console.print(f'[green]Path set:[/green] {self.path}')
        else:
            self.console.print(str(self.path))
        return True

    def _save(self, args: list[str]) -> bool:
        path = self._resolve_doc_path(args[0]) if args else self.path
        self.engine.save(path)
        self.path = Path(path)
        self.dirty = False
        self.console.print(f'[green]Saved:[/green] {path}')
        return True

    # -- environment -------------------------------------------------------

    def _env(self, args: list[str]) -> bool:
        """Show or modify the REPL environment (aliases, latex flag)."""
        sub = args[0].lower() if args else 'show'
        if sub in ('show', ''):
            self.console.print('[bold]REPL Environment[/bold]')
            self.console.print(f'  latex_style: {self.latex_style}')
            self.console.print(f'  startup commands: {len(self.env.startup_commands)}')
            if self.env.aliases:
                self.console.print('  aliases:')
                for name in sorted(self.env.aliases):
                    self.console.print(f'    {name} = {self.env.aliases[name]}')
            else:
                self.console.print('  aliases: (none)')
            return True
        if sub == 'alias':
            if len(args) < 3:
                raise ValueError('env alias requires <name> <command...>')
            name = args[1].lower()
            target = ' '.join(args[2:])
            if name == 'env':
                self.console.print('[red]Cannot shadow the env command.[/red]')
                return True
            if name in BUILTIN_COMMANDS:
                self.console.print(
                    f'[yellow]Warning:[/yellow] "{name}" shadows a built-in command.'
                )
            self.env.aliases[name] = target
            self.console.print(f'[green]Alias set:[/green] {name} = {target}')
            return True
        if sub == 'unalias':
            if len(args) < 2:
                raise ValueError('env unalias requires <name>')
            name = args[1].lower()
            if name not in self.env.aliases:
                self.console.print(f'[yellow]No such alias:[/yellow] {name}')
                return True
            del self.env.aliases[name]
            self.console.print(f'[green]Alias removed:[/green] {name}')
            return True
        self.console.print('Usage: env | env alias <name> <command...> | env unalias <name>')
        return True


def _split_tokens(line: str) -> list[str]:
    r"""Split a REPL line, preserving backslashes (Windows paths, LaTeX).

    ``shlex.split(posix=True)`` would treat ``\\`` as an escape and corrupt
    ``@C:\\Users\\...`` and ``\\frac{a}{b}``, so ``posix=False`` is used and
    surrounding quotes are stripped here.
    """
    tokens = []
    for token in shlex.split(line, posix=False):
        if len(token) >= 2 and token[0] == token[-1] and token[0] in ('"', "'"):
            token = token[1:-1]
        tokens.append(token)
    return tokens
