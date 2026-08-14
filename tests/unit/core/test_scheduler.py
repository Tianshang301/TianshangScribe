"""Unit tests for the cron scheduler."""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import pytest

from src.core.scheduler import (
    RunResult,
    SchedulerError,
    TaskScheduler,
    build_command,
    parse_command,
)
from src.utils.store import Schedule, ScheduleStore


def _make_store(tmp_path: Path) -> ScheduleStore:
    return ScheduleStore(tmp_path / 's.db')


def _py(code: str) -> list[str]:
    """Return an argv list that runs ``code`` in the current interpreter."""
    return [sys.executable, '-c', code]


ECHO_OK = _py('print(chr(104) + chr(105))')  # prints "hi", exit 0
FAIL = _py('raise SystemExit(3)')
SLEEP5 = _py('import time; time.sleep(5)')


class TestCronTrigger:
    def test_invalid_cron_raises(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='not a cron', command=ECHO_OK)
            with pytest.raises(SchedulerError):
                TaskScheduler(store).next_fire_time(sched)

    def test_next_fire_time(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='0 9 * * *', command=ECHO_OK)
            nxt = TaskScheduler(store).next_fire_time(sched, datetime(2026, 1, 1, 0, 0))
            assert nxt is not None
            assert nxt.hour == 9
            assert nxt.minute == 0


class TestIsDue:
    def test_disabled_not_due(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=ECHO_OK, enabled=False)
            assert TaskScheduler(store).is_due(sched) is False

    def test_never_run_is_due_if_window_open(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=ECHO_OK)
            now = datetime(2026, 1, 1, 10, 1, 0)
            assert TaskScheduler(store).is_due(sched, now) is True

    def test_future_cron_not_due(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='0 0 1 1 *', command=ECHO_OK)
            now = datetime(2026, 2, 1, 0, 0)
            assert TaskScheduler(store).is_due(sched, now) is False

    def test_recently_run_not_due_again(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(
                name='a',
                cron='* * * * *',
                command=ECHO_OK,
                last_run=datetime(2026, 1, 1, 10, 0, 40).isoformat(),
            )
            now = datetime(2026, 1, 1, 10, 0, 45)
            assert TaskScheduler(store).is_due(sched, now) is False

    def test_run_after_next_minute_due(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(
                name='a',
                cron='* * * * *',
                command=ECHO_OK,
                last_run=datetime(2026, 1, 1, 10, 0, 40).isoformat(),
            )
            now = datetime(2026, 1, 1, 10, 1, 0)
            assert TaskScheduler(store).is_due(sched, now) is True


class TestDependencies:
    def test_no_deps_satisfied(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=ECHO_OK)
            assert TaskScheduler(store).dependencies_satisfied(sched) is True

    def test_missing_dep_fails(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=ECHO_OK, depends_on=['base'])
            assert TaskScheduler(store).dependencies_satisfied(sched) is False

    def test_failed_dep_fails(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            store.upsert(
                Schedule(
                    name='base',
                    cron='* * * * *',
                    command=ECHO_OK,
                    last_status='failed',
                    last_exit_code=1,
                )
            )
            sched = Schedule(name='a', cron='* * * * *', command=ECHO_OK, depends_on=['base'])
            assert TaskScheduler(store).dependencies_satisfied(sched) is False

    def test_ok_dep_satisfied(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            store.upsert(
                Schedule(
                    name='base',
                    cron='* * * * *',
                    command=ECHO_OK,
                    last_status='ok',
                    last_exit_code=0,
                )
            )
            sched = Schedule(name='a', cron='* * * * *', command=ECHO_OK, depends_on=['base'])
            assert TaskScheduler(store).dependencies_satisfied(sched) is True


class TestRun:
    def test_run_success(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=ECHO_OK)
            result = TaskScheduler(store).run(sched)
            assert result.exit_code == 0
            assert result.status == 'ok'
            assert 'hi' in result.output
            got = store.get('a')
            assert got is not None
            assert got.last_status == 'ok'

    def test_run_failure_records(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=FAIL)
            result = TaskScheduler(store).run(sched)
            assert result.exit_code == 3
            assert result.status == 'failed'
            got = store.get('a')
            assert got is not None
            assert got.last_status == 'failed'

    def test_run_empty_command_raises(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=[])
            with pytest.raises(SchedulerError):
                TaskScheduler(store).run(sched)

    def test_run_timeout(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=SLEEP5)
            result = TaskScheduler(store).run(sched, timeout=0.2)
            assert result.timed_out is True
            assert result.exit_code == 124
            got = store.get('a')
            assert got is not None
            assert got.last_status == 'failed'

    def test_run_command_not_found(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=['_no_such_binary_xyz'])
            result = TaskScheduler(store).run(sched)
            assert result.exit_code == 127
            assert 'not found' in result.output


class TestRunIfDue:
    def test_not_due_returns_none(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='0 0 1 1 *', command=ECHO_OK)
            assert TaskScheduler(store).run_if_due(sched) is None

    def test_dep_unsatisfied_returns_none(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=ECHO_OK, depends_on=['nope'])
            assert TaskScheduler(store).run_if_due(sched) is None

    def test_due_and_satisfied_runs(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            sched = Schedule(name='a', cron='* * * * *', command=ECHO_OK)
            result = TaskScheduler(store).run_if_due(sched, now=datetime(2026, 1, 1, 10, 1, 0))
            assert result is not None
            assert result.exit_code == 0


class TestRunAllDue:
    def test_runs_in_dependency_order(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            store.upsert(Schedule(name='base', cron='* * * * *', command=ECHO_OK))
            store.upsert(
                Schedule(
                    name='child',
                    cron='* * * * *',
                    command=ECHO_OK,
                    depends_on=['base'],
                )
            )
            results = TaskScheduler(store).run_all_due(now=datetime(2026, 1, 1, 10, 1, 0))
            assert len(results) == 2
            assert {r.name for r in results} == {'base', 'child'}

    def test_blocked_chain_skips_child(self, tmp_path: Path) -> None:
        with _make_store(tmp_path) as store:
            store.upsert(Schedule(name='base', cron='* * * * *', command=FAIL))
            store.upsert(
                Schedule(
                    name='child',
                    cron='* * * * *',
                    command=ECHO_OK,
                    depends_on=['base'],
                )
            )
            results = TaskScheduler(store).run_all_due(now=datetime(2026, 1, 1, 10, 1, 0))
            assert len(results) == 1
            assert results[0].name == 'base'


class TestHelpers:
    def test_parse_command(self) -> None:
        assert parse_command('echo "a b" c') == ['echo', 'a b', 'c']

    def test_build_command(self) -> None:
        cmd = build_command(['--run-script', 'x.py'])
        assert cmd[-2:] == ['--run-script', 'x.py']

    def test_run_result_status(self) -> None:
        assert RunResult('a', 't', 0, 'out').status == 'ok'
        assert RunResult('a', 't', 1, 'out').status == 'failed'
