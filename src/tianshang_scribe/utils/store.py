"""SQLite-backed persistence for scheduled tasks and run history.

The :class:`ScheduleStore` persists named schedules (cron expression + command)
and their execution history in a single SQLite database. It is used by the
scheduler (:mod:`tianshang_scribe.core.scheduler`) and the ``--schedule-*`` CLI options, and
kept deliberately engine-agnostic so agents can share one store.

Concurrency: a single writer is assumed (the CLI is single-process). A
``WAL`` journal is used so a long-lived scheduler and short CLI invocations can
coexist on the same file without write-timeout failures.
"""

from __future__ import annotations

import json
import sqlite3
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, List


@dataclass
class Schedule:
    """A named, cron-triggered command."""

    name: str
    cron: str
    command: list[str] = field(default_factory=list)
    enabled: bool = True
    depends_on: list[str] = field(default_factory=list)
    timeout: float = 300.0
    last_run: str | None = None
    last_status: str | None = None
    last_exit_code: int | None = None

    def to_dict(self) -> dict[str, Any]:
        """Return the schedule as a plain dictionary."""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Schedule:
        """Build a schedule from a dictionary produced by :meth:`to_dict`."""
        return cls(
            name=data['name'],
            cron=data['cron'],
            command=list(data.get('command') or []),
            enabled=bool(data.get('enabled', True)),
            depends_on=list(data.get('depends_on') or []),
            timeout=float(data.get('timeout', 300.0)),
            last_run=data.get('last_run'),
            last_status=data.get('last_status'),
            last_exit_code=data.get('last_exit_code'),
        )


@dataclass
class RunRecord:
    """One execution attempt of a schedule."""

    schedule_name: str
    started_at: str
    finished_at: str | None = None
    exit_code: int | None = None
    output: str = ''


class ScheduleStore:
    """SQLite persistence for schedules and their run history."""

    def __init__(self, path: str | Path) -> None:
        """Open (or create) the SQLite database at ``path``."""
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._conn = sqlite3.connect(str(self.path))
        self._conn.row_factory = sqlite3.Row
        self._conn.execute('PRAGMA journal_mode=WAL')
        self._init_schema()

    def _init_schema(self) -> None:
        with self._conn:
            self._conn.execute(
                """
                CREATE TABLE IF NOT EXISTS schedules (
                    name        TEXT PRIMARY KEY,
                    cron        TEXT NOT NULL,
                    command     TEXT NOT NULL DEFAULT '[]',
                    enabled     INTEGER NOT NULL DEFAULT 1,
                    depends_on  TEXT NOT NULL DEFAULT '[]',
                    timeout     REAL NOT NULL DEFAULT 300.0,
                    last_run    TEXT,
                    last_status TEXT,
                    last_exit_code INTEGER
                )
                """
            )
            self._conn.execute(
                """
                CREATE TABLE IF NOT EXISTS runs (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    schedule    TEXT NOT NULL,
                    started_at  TEXT NOT NULL,
                    finished_at TEXT,
                    exit_code   INTEGER,
                    output      TEXT NOT NULL DEFAULT ''
                )
                """
            )

    # -- schedules -----------------------------------------------------------

    def upsert(self, schedule: Schedule) -> None:
        """Insert or replace a schedule (keyed by name)."""
        with self._conn:
            self._conn.execute(
                """
                INSERT INTO schedules (
                    name, cron, command, enabled, depends_on, timeout,
                    last_run, last_status, last_exit_code
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(name) DO UPDATE SET
                    cron=excluded.cron,
                    command=excluded.command,
                    enabled=excluded.enabled,
                    depends_on=excluded.depends_on,
                    timeout=excluded.timeout
                """,
                (
                    schedule.name,
                    schedule.cron,
                    json.dumps(schedule.command),
                    int(schedule.enabled),
                    json.dumps(schedule.depends_on),
                    schedule.timeout,
                    schedule.last_run,
                    schedule.last_status,
                    schedule.last_exit_code,
                ),
            )

    def get(self, name: str) -> Schedule | None:
        """Return the named schedule, or None when it does not exist."""
        row = self._conn.execute('SELECT * FROM schedules WHERE name = ?', (name,)).fetchone()
        return self._row_to_schedule(row) if row is not None else None

    def list(self, enabled_only: bool = False) -> list[Schedule]:
        """Return all schedules, optionally only enabled ones (sorted by name)."""
        query = 'SELECT * FROM schedules'
        if enabled_only:
            query += ' WHERE enabled = 1'
        query += ' ORDER BY name'
        rows = self._conn.execute(query).fetchall()
        return [self._row_to_schedule(r) for r in rows]

    def delete(self, name: str) -> bool:
        """Delete a schedule by name; return True when it existed."""
        cur = self._conn.execute('DELETE FROM schedules WHERE name = ?', (name,))
        return cur.rowcount > 0

    def set_enabled(self, name: str, enabled: bool) -> bool:
        """Enable or disable a schedule; return True when it existed."""
        cur = self._conn.execute(
            'UPDATE schedules SET enabled = ? WHERE name = ?', (int(enabled), name)
        )
        return cur.rowcount > 0

    def mark_result(
        self,
        name: str,
        *,
        started_at: str,
        finished_at: str | None = None,
        exit_code: int | None = None,
        output: str = '',
    ) -> None:
        """Record a run and update the schedule's ``last_*`` columns."""
        status = 'ok' if exit_code == 0 else 'failed'
        with self._conn:
            self._conn.execute(
                """
                UPDATE schedules
                SET last_run = ?, last_status = ?, last_exit_code = ?
                WHERE name = ?
                """,
                (started_at, status, exit_code, name),
            )
            self._conn.execute(
                """
                INSERT INTO runs (schedule, started_at, finished_at, exit_code, output)
                VALUES (?, ?, ?, ?, ?)
                """,
                (name, started_at, finished_at, exit_code, output),
            )

    def history(self, name: str | None = None, limit: int = 100) -> List[RunRecord]:
        """Return run history, optionally filtered by schedule, newest first."""
        query = 'SELECT * FROM runs'
        params: tuple[Any, ...] = ()
        if name:
            query += ' WHERE schedule = ?'
            params = (name,)
        query += ' ORDER BY id DESC LIMIT ?'
        rows = self._conn.execute(query, (*params, limit)).fetchall()
        return [
            RunRecord(
                schedule_name=r['schedule'],
                started_at=r['started_at'],
                finished_at=r['finished_at'],
                exit_code=r['exit_code'],
                output=r['output'],
            )
            for r in rows
        ]

    # -- helpers -------------------------------------------------------------

    def _row_to_schedule(self, row: sqlite3.Row) -> Schedule:
        return Schedule.from_dict(
            {
                'name': row['name'],
                'cron': row['cron'],
                'command': json.loads(row['command'] or '[]'),
                'enabled': bool(row['enabled']),
                'depends_on': json.loads(row['depends_on'] or '[]'),
                'timeout': row['timeout'],
                'last_run': row['last_run'],
                'last_status': row['last_status'],
                'last_exit_code': row['last_exit_code'],
            }
        )

    def close(self) -> None:
        """Close the underlying database connection."""
        self._conn.close()

    def __enter__(self) -> ScheduleStore:
        """Support ``with`` statements for scoped connections."""
        return self

    def __exit__(self, *exc: object) -> None:
        """Close the connection when leaving the ``with`` block."""
        self.close()


def now_iso() -> str:
    """Return the current UTC time in ISO 8601 format (used for run timestamps)."""
    return datetime.now(timezone.utc).isoformat()
