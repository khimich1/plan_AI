"""ProductionCapacityService (TASK-005)."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

from app.services.production_capacity_service import (
    ProductionCapacityError,
    ProductionCapacityService,
)
from core import kp_db_schema
from core.kp_db_common import _connect


def _fresh_db(tmp_path: Path, name: str = "capacity_service.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _service(tmp_path: Path) -> ProductionCapacityService:
    return ProductionCapacityService(db_path=_fresh_db(tmp_path))


def test_get_capacity_map_defaults_to_five(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    days = [date(2026, 9, 1), "2026-09-02"]
    got = svc.get_capacity_map(days)
    assert got == {
        date(2026, 9, 1): 5,
        date(2026, 9, 2): 5,
    }
    assert all(isinstance(k, date) for k in got)
    assert all(type(v) is int for v in got.values())


def test_set_then_get_reflects_override(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    day = date(2026, 9, 3)
    svc.set_day_capacity(day, 4, user={"email": "alice@example.com"})
    got = svc.get_capacity_map([day, "2026-09-04"])
    assert got[day] == 4
    assert got[date(2026, 9, 4)] == 5


def test_validate_fill_targets_ok(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    svc.set_day_capacity("2026-09-10", 4)
    svc.validate_fill_targets([{"date": "2026-09-10", "tracks": 3}])


def test_validate_fill_targets_respects_override(tmp_path: Path) -> None:
    """Override below default must shrink day max; tracks==max ok, max+1 fails."""
    svc = _service(tmp_path)
    svc.set_day_capacity("2026-09-10", 3)
    svc.validate_fill_targets([{"date": "2026-09-10", "tracks": 3}])
    with pytest.raises(ProductionCapacityError, match="свободно 3"):
        svc.validate_fill_targets([{"date": "2026-09-10", "tracks": 4}])


def test_set_day_capacity_rejects_above_hard_cap(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    with pytest.raises(ProductionCapacityError, match="hard cap"):
        svc.set_day_capacity(date(2026, 9, 1), 6)

def test_validate_fill_targets_empty_raises(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    with pytest.raises(ProductionCapacityError, match="fill_targets пуст"):
        svc.validate_fill_targets([])


def test_validate_fill_targets_over_capacity_raises(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    with pytest.raises(ProductionCapacityError, match="свободно 5"):
        svc.validate_fill_targets([{"date": "2026-09-11", "tracks": 6}])


def test_negative_max_tracks_raises(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    with pytest.raises(ProductionCapacityError, match="max_tracks"):
        svc.set_day_capacity(date(2026, 9, 1), -1)


def test_updated_by_extracted_from_user_dict(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    day = date(2026, 9, 12)
    svc.set_day_capacity(day, 4, user={"email": "bob@example.com", "login": "bob"})

    with _connect(svc.db_path) as conn:
        updated_by = conn.execute(
            "SELECT updated_by FROM day_capacity_override WHERE date = ?",
            ("2026-09-12",),
        ).fetchone()[0]
    assert updated_by == "bob@example.com"


def test_updated_by_falls_back_to_login_then_user_id(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    svc.set_day_capacity("2026-09-13", 3, user={"login": "carol"})
    with _connect(svc.db_path) as conn:
        updated_by = conn.execute(
            "SELECT updated_by FROM day_capacity_override WHERE date = ?",
            ("2026-09-13",),
        ).fetchone()[0]
    assert updated_by == "carol"

    svc.set_day_capacity("2026-09-14", 2, user={"user_id": 42})
    with _connect(svc.db_path) as conn:
        updated_by = conn.execute(
            "SELECT updated_by FROM day_capacity_override WHERE date = ?",
            ("2026-09-14",),
        ).fetchone()[0]
    assert updated_by == "42"
