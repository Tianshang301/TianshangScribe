"""Unit tests for the SQLite schedule store."""

from __future__ import annotations

from pathlib import Path

from src.utils.store import Schedule, ScheduleStore, now_iso


class TestScheduleStore:
    def test_upsert_and_get(self, tmp_path: Path) -> None:
        with ScheduleStore(tmp_path / 's.db') as store:
            sched = Schedule(name='daily', cron='0 9 * * *', command=['echo', 'hi'])
            store.upsert(sched)
            got = store.get('daily')
            assert got is not None
            assert got.name == 'daily'
            assert got.cron == '0 9 * * *'
            assert got.command == ['echo', 'hi']
            assert got.enabled is True

    def test_upsert_replaces_existing(self, tmp_path: Path) -> None:
        with ScheduleStore(tmp_path / 's.db') as store:
            store.upsert(Schedule(name='a', cron='* * * * *', command=['a']))
            store.upsert(Schedule(name='a', cron='0 0 * * *', command=['b']))
            got = store.get('a')
            assert got is not None
            assert got.cron == '0 0 * * *'
            assert got.command == ['b']

    def test_list_and_enabled_filter(self, tmp_path: Path) -> None:
        with ScheduleStore(tmp_path / 's.db') as store:
            store.upsert(Schedule(name='a', cron='* * * * *', command=['a']))
            store.upsert(Schedule(name='b', cron='* * * * *', command=['b'], enabled=False))
            assert {s.name for s in store.list()} == {'a', 'b'}
            assert {s.name for s in store.list(enabled_only=True)} == {'a'}

    def test_delete(self, tmp_path: Path) -> None:
        with ScheduleStore(tmp_path / 's.db') as store:
            store.upsert(Schedule(name='a', cron='* * * * *', command=['a']))
            assert store.delete('a') is True
            assert store.get('a') is None
            assert store.delete('a') is False

    def test_set_enabled(self, tmp_path: Path) -> None:
        with ScheduleStore(tmp_path / 's.db') as store:
            store.upsert(Schedule(name='a', cron='* * * * *', command=['a']))
            assert store.set_enabled('a', False) is True
            assert store.get('a').enabled is False
            assert store.set_enabled('missing', False) is False

    def test_mark_result_updates_schedule_and_history(self, tmp_path: Path) -> None:
        with ScheduleStore(tmp_path / 's.db') as store:
            store.upsert(Schedule(name='a', cron='* * * * *', command=['a']))
            store.mark_result('a', started_at='t1', finished_at='t2', exit_code=0, output='out')
            sched = store.get('a')
            assert sched is not None
            assert sched.last_run == 't1'
            assert sched.last_status == 'ok'
            assert sched.last_exit_code == 0
            hist = store.history('a')
            assert len(hist) == 1
            assert hist[0].exit_code == 0
            assert hist[0].output == 'out'

    def test_history_failed_and_limit(self, tmp_path: Path) -> None:
        with ScheduleStore(tmp_path / 's.db') as store:
            store.upsert(Schedule(name='a', cron='* * * * *', command=['a']))
            store.mark_result('a', started_at='1', exit_code=1)
            store.mark_result('a', started_at='2', exit_code=0)
            hist = store.history('a')
            assert len(hist) == 2
            assert hist[0].started_at == '2'
            assert hist[0].exit_code == 0
            limited = store.history('a', limit=1)
            assert len(limited) == 1

    def test_history_all_schedules(self, tmp_path: Path) -> None:
        with ScheduleStore(tmp_path / 's.db') as store:
            store.upsert(Schedule(name='a', cron='* * * * *', command=['a']))
            store.upsert(Schedule(name='b', cron='* * * * *', command=['b']))
            store.mark_result('a', started_at='1', exit_code=0)
            store.mark_result('b', started_at='2', exit_code=0)
            assert len(store.history()) == 2

    def test_persists_across_connections(self, tmp_path: Path) -> None:
        db = tmp_path / 's.db'
        with ScheduleStore(db) as store:
            store.upsert(Schedule(name='a', cron='* * * * *', command=['a']))
        with ScheduleStore(db) as store:
            assert store.get('a') is not None

    def test_round_trip_optional_fields(self, tmp_path: Path) -> None:
        with ScheduleStore(tmp_path / 's.db') as store:
            sched = Schedule(
                name='a',
                cron='*/5 * * * *',
                command=['x', 'y'],
                depends_on=['base'],
                timeout=10.5,
                last_run='2026-01-01T00:00:00+00:00',
                last_status='failed',
                last_exit_code=1,
            )
            store.upsert(sched)
            got = store.get('a')
            assert got is not None
            assert got.depends_on == ['base']
            assert got.timeout == 10.5
            assert got.last_status == 'failed'
            assert got.last_exit_code == 1


class TestNowIso:
    def test_returns_iso_string(self) -> None:
        from datetime import datetime

        val = now_iso()
        parsed = datetime.fromisoformat(val)
        assert parsed.tzinfo is not None
