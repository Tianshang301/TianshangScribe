"""Cron-based scheduler with dependency chains.

Schedules are stored in a :class:`~src.utils.store.ScheduleStore` (SQLite).
A schedule fires when its cron expression matches and every dependency listed
in ``depends_on`` has completed successfully at least once (``last_status ==
'ok'``). Commands run as subprocesses so a runaway job cannot corrupt the
caller, and each run is recorded with exit code and captured output.

The CLI exposes ``--schedule-add/--schedule-rm/--schedule-list/--schedule-run``
(see :mod:`src.cli.main`).
"""

from __future__ import annotations

import shlex
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime, timezone

from apscheduler.triggers.cron import CronTrigger

from src.utils.store import Schedule, ScheduleStore, now_iso

UTC = timezone.utc


@dataclass
class RunResult:
    """Outcome of executing one schedule."""

    name: str
    started_at: str
    exit_code: int
    output: str
    status: str = field(init=False)
    timed_out: bool = False

    def __post_init__(self) -> None:
        """Derive the ``status`` field from the exit code."""
        self.status = 'ok' if self.exit_code == 0 else 'failed'


class SchedulerError(ValueError):
    """Raised for invalid cron expressions or unsatisfiable schedules."""


def _cron_trigger(expr: str) -> CronTrigger:
    try:
        return CronTrigger.from_crontab(expr)
    except (ValueError, TypeError) as e:
        raise SchedulerError(f'invalid cron expression {expr!r}: {e}') from None


class TaskScheduler:
    """Execute due schedules respecting cron expressions and dependencies."""

    def __init__(self, store: ScheduleStore) -> None:
        """Create a scheduler backed by ``store`` for schedule persistence."""
        self.store = store

    def next_fire_time(
        self,
        schedule: Schedule,
        now: datetime | None = None,
        previous: datetime | None = None,
    ) -> datetime | None:
        """Compute the next fire time for a schedule (UTC)."""
        trigger = _cron_trigger(schedule.cron)
        nxt = trigger.get_next_fire_time(previous, now or datetime.now(UTC))
        return nxt  # type: ignore[no-any-return]  # apscheduler returns datetime | None

    def is_due(self, schedule: Schedule, now: datetime | None = None) -> bool:
        """Return whether a schedule should fire now.

        A schedule is due when it is enabled and its cron expression has a
        next fire time at or before ``now``. ``last_run`` is used as the
        reference point so a schedule never re-fires within the same cron
        window.
        """
        if not schedule.enabled:
            return False
        now = now or datetime.now(UTC)
        if now.tzinfo is None:
            now = now.replace(tzinfo=UTC)
        previous: datetime | None = None
        if schedule.last_run:
            try:
                previous = datetime.fromisoformat(schedule.last_run)
                if previous.tzinfo is None:
                    previous = previous.replace(tzinfo=UTC)
            except ValueError:
                previous = None
        nxt = self.next_fire_time(schedule, now, previous=previous)
        return nxt is not None and nxt <= now

    def dependencies_satisfied(self, schedule: Schedule) -> bool:
        """Return True when every dependency has completed successfully."""
        if not schedule.depends_on:
            return True
        for dep in schedule.depends_on:
            d = self.store.get(dep)
            if d is None or d.last_status != 'ok':
                return False
        return True

    def run(self, schedule: Schedule, timeout: float | None = None) -> RunResult:
        """Execute a schedule's command as a subprocess and record the run."""
        started = now_iso()
        cmd = schedule.command
        if not cmd:
            raise SchedulerError(f'schedule {schedule.name!r} has an empty command')
        self.store.upsert(schedule)
        effective_timeout = timeout if timeout is not None else schedule.timeout
        try:
            completed = subprocess.run(  # noqa: S603  # commands are user-defined schedules (equivalent to crontab)
                cmd,
                capture_output=True,
                text=True,
                timeout=effective_timeout,
                check=False,
            )
            output = (completed.stdout or '') + (completed.stderr or '')
            result = RunResult(
                name=schedule.name,
                started_at=started,
                exit_code=completed.returncode,
                output=output,
            )
        except subprocess.TimeoutExpired as e:
            result = RunResult(
                name=schedule.name,
                started_at=started,
                exit_code=124,
                output=(e.stdout or '') if isinstance(e.stdout, str) else '',
                timed_out=True,
            )
        except (FileNotFoundError, OSError) as e:
            result = RunResult(
                name=schedule.name,
                started_at=started,
                exit_code=127,
                output=f'command not found: {cmd!r} ({e})',
            )
        self.store.mark_result(
            schedule.name,
            started_at=started,
            finished_at=now_iso(),
            exit_code=result.exit_code,
            output=result.output,
        )
        return result

    def run_if_due(
        self, schedule: Schedule, now: datetime | None = None, timeout: float | None = None
    ) -> RunResult | None:
        """Run a schedule now, but only when its cron window and deps allow it."""
        if not self.is_due(schedule, now):
            return None
        if not self.dependencies_satisfied(schedule):
            return None
        return self.run(schedule, timeout)

    def run_all_due(self, now: datetime | None = None) -> list[RunResult]:
        """Run every due schedule in dependency order; return the results."""
        schedules = self.store.list(enabled_only=True)
        results: list[RunResult] = []
        remaining = list(schedules)
        while remaining:
            progressed = False
            for schedule in list(remaining):
                if self.is_due(schedule, now) and self.dependencies_satisfied(schedule):
                    results.append(self.run(schedule))
                    remaining.remove(schedule)
                    progressed = True
            if not progressed:
                break
        return results


def build_command(args: list[str]) -> list[str]:
    """Return the shell command to re-invoke the CLI with ``args``."""
    return [sys.executable, '-m', 'src.cli.main', *args]


def parse_command(text: str) -> list[str]:
    """Split a command string into argv, stripping surrounding quotes."""
    tokens = shlex.split(text, posix=False)
    cleaned: list[str] = []
    for token in tokens:
        if len(token) >= 2 and token[0] == token[-1] and token[0] in ('"', "'"):
            token = token[1:-1]
        cleaned.append(token)
    return cleaned
