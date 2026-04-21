"""Unit-тесты для :mod:`core.plan_commit`.

Покрывают три публичные функции модуля:

- :func:`count_assigned_plates` — корректный подсчёт по источникам, разделение
  сопоставленных и несопоставленных назначений.
- :func:`distribute_assigned_plates_to_orders` — распределение по строкам
  ``orders_2d`` в правильном порядке источников (``primary → secondary → rescue``)
  и корректный расчёт ``leftovers``.
- :func:`commit_plan_plates` — поведение happy-path, выброс ``PlanCommitError``
  при несопоставленных плитах и при остатках, откат пометки при проваленной
  ``mark_plates_as_planned``.
"""
from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core import kp_db
from core.plan_commit import (
    PlanCommitError,
    commit_plan_plates,
    count_assigned_plates,
    distribute_assigned_plates_to_orders,
)


@pytest.fixture
def tmp_db(tmp_path) -> str:
    db_path = str(tmp_path / "plita_test.db")
    kp_db.init_schema(db_path)
    return db_path


def _seed_kp_plate(
    db_path: str,
    *,
    kp_id: int = 1,
    plate_name: str = "ПБ 60-12-8п",
    qty: int = 3,
    length_m: float = 6.0,
    width_m: float = 1.2,
    load_class: int = 800,
) -> None:
    """Создаёт одну строку ``kp_plates`` со статусом ``'в производстве'``."""
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (?, ?) "
            "ON CONFLICT(kp_id) DO NOTHING",
            (kp_id, "2026-01-01"),
        )
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name,
                length_m, width_m, load_class, qty, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 'в производстве')
            """,
            (kp_id, 1, plate_name, length_m, width_m, load_class, qty),
        )
        conn.commit()


def _fetch_status(db_path: str, kp_id: int, plate_name: str) -> list[tuple[str, int, str | None]]:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT status, qty, plan_id FROM kp_plates WHERE kp_id = ? AND plate_name = ? ORDER BY id",
            (kp_id, plate_name),
        )
        return cur.fetchall()


def test_count_assigned_plates_groups_by_source():
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "A", "length": 6.0},
            {"source": "primary", "kp_id": 1, "plate_name": "A", "length": 6.0},
            {"source": "secondary", "kp_id": 2, "plate_name": "B", "length": 5.0},
            {"source": "primary", "kp_id": None, "plate_name": "", "length": 3.0},
        ],
    }
    rescue_tracks = [
        {
            "label": "РЕСКЬЮ",
            "items": [
                {"kp_id": 3, "plate_name": "C", "length": 2.0},
                {"kp_id": None, "plate_name": "", "length": 1.0},
            ],
        },
        {"label": "ОСНОВНАЯ", "items": [{"kp_id": 9, "plate_name": "X"}]},
    ]

    counts, unmapped = count_assigned_plates(optimization_result, rescue_tracks)

    assert counts["primary"] == {(1, "A"): 2}
    assert counts["secondary"] == {(2, "B"): 1}
    assert counts["rescue"] == {(3, "C"): 1}
    assert len(unmapped["primary"]) == 1
    assert len(unmapped["rescue"]) == 1


def test_distribute_uses_primary_before_rescue():
    orders = [
        {"kp_id": 1, "plate_name": "A", "qty": 3, "load_code": 8},
        {"kp_id": 2, "plate_name": "B", "qty": 2, "load_code": 8},
    ]
    assigned = {
        "primary": {(1, "A"): 2},
        "secondary": {(2, "B"): 1},
        "rescue": {(1, "A"): 1, (2, "B"): 2},
    }

    lost, orders_with_qty, leftovers = distribute_assigned_plates_to_orders(orders, assigned)

    assert orders_with_qty[0][1] == 3  # A: 2 primary + 1 rescue
    assert orders_with_qty[1][1] == 2  # B: 1 secondary + 1 rescue
    assert lost == []
    # rescue имел лишнюю B, она остаётся в leftovers
    assert leftovers["rescue"] == {(2, "B"): 1}


def test_distribute_reports_lost_plates_when_source_short():
    orders = [{"kp_id": 1, "plate_name": "A", "qty": 3, "load_code": 8}]
    assigned = {
        "primary": {(1, "A"): 1},
        "secondary": {},
        "rescue": {},
    }

    lost, orders_with_qty, leftovers = distribute_assigned_plates_to_orders(orders, assigned)

    assert orders_with_qty[0][1] == 1
    assert lost == [
        {"kp_id": 1, "plate_name": "A", "qty_lost": 2, "load_code": 8},
    ]
    assert leftovers["primary"] == {}


def test_commit_plan_plates_happy_path(tmp_db):
    _seed_kp_plate(tmp_db, kp_id=1, plate_name="ПБ 60-12-8п", qty=3)
    orders = [{"kp_id": 1, "plate_name": "ПБ 60-12-8п", "qty": 3, "load_code": 8}]
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "ПБ 60-12-8п"} for _ in range(3)
        ],
    }

    result = commit_plan_plates(
        plan_id="plan_test",
        orders_2d=orders,
        optimization_result=optimization_result,
        all_tracks_list=[],
        db_path=tmp_db,
    )

    assert result.plates_marked == 3
    rows = _fetch_status(tmp_db, 1, "ПБ 60-12-8п")
    assert rows == [("в плане", 3, "plan_test")]


def test_commit_plan_plates_raises_on_unmapped(tmp_db):
    _seed_kp_plate(tmp_db, kp_id=1, plate_name="A", qty=1)
    orders = [{"kp_id": 1, "plate_name": "A", "qty": 1, "load_code": 8}]
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": None, "plate_name": "", "length": 1.0},
        ],
    }

    with pytest.raises(PlanCommitError):
        commit_plan_plates(
            plan_id="plan_unmapped",
            orders_2d=orders,
            optimization_result=optimization_result,
            all_tracks_list=[],
            db_path=tmp_db,
        )

    # Плиты остаются в производстве
    rows = _fetch_status(tmp_db, 1, "A")
    assert rows == [("в производстве", 1, None)]


def test_commit_plan_plates_raises_on_leftovers(tmp_db):
    _seed_kp_plate(tmp_db, kp_id=1, plate_name="A", qty=1)
    orders = [{"kp_id": 1, "plate_name": "A", "qty": 1, "load_code": 8}]
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "A"},
            {"source": "primary", "kp_id": 99, "plate_name": "NO_SUCH_ORDER"},
        ],
    }

    with pytest.raises(PlanCommitError):
        commit_plan_plates(
            plan_id="plan_leftover",
            orders_2d=orders,
            optimization_result=optimization_result,
            all_tracks_list=[],
            db_path=tmp_db,
        )


def test_commit_plan_plates_rolls_back_on_mark_failure(tmp_db, monkeypatch):
    _seed_kp_plate(tmp_db, kp_id=1, plate_name="A", qty=2)
    orders = [{"kp_id": 1, "plate_name": "A", "qty": 2, "load_code": 8}]
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "A"} for _ in range(2)
        ],
    }

    calls: list[dict] = []

    def fake_mark(*, kp_id, plate_name, qty_to_plan, plan_id, db_path):
        calls.append({"plan_id": plan_id})
        return {"success": False, "processed_count": 0}

    rollback_calls: list[str] = []
    real_return = kp_db.return_plan_plates_to_production

    def spy_return(plan_id, db_path):
        rollback_calls.append(plan_id)
        return real_return(plan_id, db_path)

    monkeypatch.setattr("core.plan_commit.kp_db.mark_plates_as_planned", fake_mark)
    monkeypatch.setattr("core.plan_commit.kp_db.return_plan_plates_to_production", spy_return)

    with pytest.raises(PlanCommitError):
        commit_plan_plates(
            plan_id="plan_fail",
            orders_2d=orders,
            optimization_result=optimization_result,
            all_tracks_list=[],
            db_path=tmp_db,
        )

    assert calls and calls[0]["plan_id"] == "plan_fail"
    assert rollback_calls == ["plan_fail"]
