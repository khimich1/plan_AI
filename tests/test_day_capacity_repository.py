"""Day capacity override schema + DayCapacityRepository (TASK-001)."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

from app.repositories.day_capacity_repository import DayCapacityRepository
from core import kp_db_schema
from core.kp_db_common import _connect


def _fresh_db(tmp_path: Path, name: str = "day_capacity.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _repo(tmp_path: Path) -> DayCapacityRepository:
    return DayCapacityRepository(db_path=_fresh_db(tmp_path))


def test_ensure_schema_creates_day_capacity_override_table(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)

    with _connect(db_path) as conn:
        cols = {row[1] for row in conn.execute("PRAGMA table_info(day_capacity_override)")}
    assert cols == {"id", "date", "max_tracks", "updated_at", "updated_by"}


def test_ensure_schema_idempotent(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._init_schema_impl(db_path)

    with _connect(db_path) as conn:
        cols = {row[1] for row in conn.execute("PRAGMA table_info(day_capacity_override)")}
    assert "date" in cols
    assert "max_tracks" in cols


def test_get_max_tracks_default_without_override(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    got_date = repo.get_max_tracks(date(2026, 9, 1))
    got_str = repo.get_max_tracks("2026-09-01")
    assert got_date == 5
    assert got_str == 5
    assert isinstance(got_date, int)
    assert isinstance(got_str, int)
    assert type(got_date) is int  # not bool/float


def test_set_override_then_get_max_tracks(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    repo.set_override(date(2026, 9, 2), 4, updated_by="alice")
    got = repo.get_max_tracks(date(2026, 9, 2))
    assert got == 4
    assert type(got) is int
    assert repo.get_max_tracks("2026-09-02") == 4


def test_set_override_upserts_same_date(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    day = date(2026, 9, 3)
    repo.set_override(day, 5, updated_by="alice")
    repo.set_override(day, 3, updated_by="bob")
    assert repo.get_max_tracks(day) == 3

    with _connect(repo.db_path) as conn:
        count = conn.execute(
            "SELECT COUNT(*) FROM day_capacity_override WHERE date = ?",
            ("2026-09-03",),
        ).fetchone()[0]
        updated_by = conn.execute(
            "SELECT updated_by FROM day_capacity_override WHERE date = ?",
            ("2026-09-03",),
        ).fetchone()[0]
    assert count == 1
    assert updated_by == "bob"


def test_list_overrides_empty(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    assert repo.list_overrides() == {}


def test_list_overrides_returns_date_keys(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    repo.set_override("2026-09-10", 4)
    repo.set_override(date(2026, 9, 11), 5)
    overrides = repo.list_overrides()
    assert overrides == {
        date(2026, 9, 10): 4,
        date(2026, 9, 11): 5,
    }
    assert all(isinstance(k, date) for k in overrides)
    assert all(type(v) is int for v in overrides.values())


def test_set_override_negative_raises(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    with pytest.raises(ValueError, match="max_tracks"):
        repo.set_override(date(2026, 9, 1), -1)


def test_set_override_above_hard_cap_raises(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    with pytest.raises(ValueError, match="hard cap"):
        repo.set_override(date(2026, 9, 1), 6)


def test_get_max_tracks_clamps_stale_override_above_cap(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    with _connect(repo.db_path) as conn:
        conn.execute(
            "INSERT INTO day_capacity_override (date, max_tracks, updated_at, updated_by) "
            "VALUES (?, ?, ?, ?)",
            ("2026-09-06", 9, "2026-09-01T00:00:00", "legacy"),
        )
        conn.commit()
    assert repo.get_max_tracks(date(2026, 9, 6)) == 5
    assert repo.list_overrides()[date(2026, 9, 6)] == 5


def test_set_override_zero_allowed(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    repo.set_override(date(2026, 9, 5), 0)
    assert repo.get_max_tracks(date(2026, 9, 5)) == 0
