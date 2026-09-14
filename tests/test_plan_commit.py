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
    """Phase 3: счёт идёт ТОЛЬКО из plate_assignments. RESCUE-плиты тоже там
    (см. core.rescue_tracks.build_rescue_tracks возвращает rescue_assignments)."""
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "A", "length": 6.0},
            {"source": "primary", "kp_id": 1, "plate_name": "A", "length": 6.0},
            {"source": "secondary", "kp_id": 2, "plate_name": "B", "length": 5.0},
            {"source": "primary", "kp_id": None, "plate_name": "", "length": 3.0},
            {"source": "rescue", "kp_id": 3, "plate_name": "C", "length": 2.0},
            {"source": "rescue", "kp_id": None, "plate_name": "", "length": 1.0},
        ],
    }

    counts, unmapped = count_assigned_plates(optimization_result, [])

    assert counts["primary"] == {(1, "A"): 2}
    assert counts["secondary"] == {(2, "B"): 1}
    assert counts["rescue"] == {(3, "C"): 1}
    assert len(unmapped["primary"]) == 1
    assert len(unmapped["rescue"]) == 1


def test_count_assigned_plates_no_double_counting_when_tracks_passed():
    """Регрессия: даже если в all_tracks_list лежат РЕСКЬЮ-треки с теми же
    плитами, что в plate_assignments — двойного учёта НЕТ."""
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "A", "length": 6.0},
            {"source": "rescue", "kp_id": 2, "plate_name": "B", "length": 5.0},
        ],
    }
    all_tracks_list = [
        {
            "label": "РЕСКЬЮ",
            "items": [
                {"kp_id": 2, "plate_name": "B", "length": 5.0},
            ],
        },
    ]

    counts, _ = count_assigned_plates(optimization_result, all_tracks_list)

    assert counts["primary"] == {(1, "A"): 1}
    assert counts["rescue"] == {(2, "B"): 1}


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


def test_commit_plan_plates_rescue_leftovers_only_warns(tmp_db, caplog):
    """Phase 5: rescue_leftovers больше не блокируют commit.

    После Phase 1-4 фантомные rescue не возникают, но если всё же
    появились — это safety-net, не ошибка.
    """
    import logging

    _seed_kp_plate(tmp_db, kp_id=1, plate_name="A", qty=1)
    orders = [{"kp_id": 1, "plate_name": "A", "qty": 1, "load_code": 8}]
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "A"},
            {"source": "rescue", "kp_id": 1, "plate_name": "A"},
        ],
    }

    with caplog.at_level(logging.WARNING, logger="core.plan_commit"):
        result = commit_plan_plates(
            plan_id="plan_rescue_leftover",
            orders_2d=orders,
            optimization_result=optimization_result,
            all_tracks_list=[],
            db_path=tmp_db,
        )

    assert result.plates_marked == 1
    assert any(
        "RESCUE-плиты" in record.getMessage()
        for record in caplog.records
        if record.levelno == logging.WARNING
    )


def test_commit_plan_plates_all_rows_have_day_number_when_tracks_by_day_passed(tmp_db):
    """P9: с непустым ``tracks_by_day`` ВСЕ помеченные ``kp_plates`` строки
    должны иметь ``day_number != NULL``. Это и есть инвариант, который
    предотвращает «зависшие» плиты вне day_view / complete_day.
    """
    _seed_kp_plate(tmp_db, kp_id=1, plate_name="ПБ 60-12-8п", qty=4)
    orders = [{"kp_id": 1, "plate_name": "ПБ 60-12-8п", "qty": 4, "load_code": 8}]
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "ПБ 60-12-8п"}
            for _ in range(4)
        ],
    }
    tracks_by_day = {
        "2026-01-10": [
            {
                "production_day": 1,
                "items": [
                    {"length": 6.0, "mode": "solid", "width": 1.2,
                     "load_code": 8, "kp_id": 1, "plate_name": "ПБ 60-12-8п"},
                    {"length": 6.0, "mode": "solid", "width": 1.2,
                     "load_code": 8, "kp_id": 1, "plate_name": "ПБ 60-12-8п"},
                ],
            },
        ],
        "2026-01-11": [
            {
                "production_day": 2,
                "items": [
                    {"length": 6.0, "mode": "solid", "width": 1.2,
                     "load_code": 8, "kp_id": 1, "plate_name": "ПБ 60-12-8п"},
                    {"length": 6.0, "mode": "solid", "width": 1.2,
                     "load_code": 8, "kp_id": 1, "plate_name": "ПБ 60-12-8п"},
                ],
            },
        ],
    }

    commit_plan_plates(
        plan_id="plan_invariant_day",
        orders_2d=orders,
        optimization_result=optimization_result,
        all_tracks_list=[],
        db_path=tmp_db,
        tracks_by_day=tracks_by_day,
    )

    with sqlite3.connect(tmp_db) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT day_number, qty FROM kp_plates "
            "WHERE plan_id = ? AND status = 'в плане'",
            ("plan_invariant_day",),
        )
        rows = cur.fetchall()

    assert rows, "ожидаем хотя бы одну строку kp_plates со status='в плане'"
    assert all(r[0] is not None for r in rows), (
        f"найдены строки с day_number IS NULL: {rows}"
    )
    qty_by_day: dict[int, int] = {}
    for day_number, qty in rows:
        qty_by_day[day_number] = qty_by_day.get(day_number, 0) + qty
    assert qty_by_day == {1: 2, 2: 2}, (
        f"ожидаем 2 плиты на день 1 и 2 на день 2, получили {qty_by_day}"
    )


def test_commit_plan_plates_all_items_get_kp_plate_id(tmp_db):
    """P9: после commit каждый item в ``tracks_by_day`` (включая
    ``secondary_cuts``) получает ``kp_plate_id``."""
    _seed_kp_plate(tmp_db, kp_id=1, plate_name="ПБ 60-12-8п", qty=2,
                   length_m=6.0, width_m=1.2, load_class=800)
    _seed_kp_plate(tmp_db, kp_id=1, plate_name="ПБ 30-3,2-8п", qty=1,
                   length_m=3.0, width_m=0.32, load_class=800)
    orders = [
        {"kp_id": 1, "plate_name": "ПБ 60-12-8п", "qty": 2, "load_code": 8,
         "length": 6.0, "width": 1200},
        {"kp_id": 1, "plate_name": "ПБ 30-3,2-8п", "qty": 1, "load_code": 8,
         "length": 3.0, "width": 320},
    ]
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "ПБ 60-12-8п"},
            {"source": "primary", "kp_id": 1, "plate_name": "ПБ 60-12-8п"},
            {"source": "secondary", "kp_id": 1, "plate_name": "ПБ 30-3,2-8п"},
        ],
    }
    secondary_item = {
        "width": 0.32,
        "target_length": 3.0,
        "load_code": 8,
        "kp_id": 1,
        "plate_name": "ПБ 30-3,2-8п",
    }
    parent_item = {
        "length": 6.0,
        "mode": "split",
        "main_w": 1.2,
        "rest_w": 0.32,
        "load_code": 8,
        "kp_id": 1,
        "plate_name": "ПБ 60-12-8п",
        "secondary_cuts": [secondary_item],
    }
    second_solid = {
        "length": 6.0,
        "mode": "solid",
        "width": 1.2,
        "load_code": 8,
        "kp_id": 1,
        "plate_name": "ПБ 60-12-8п",
    }
    tracks_by_day = {
        "2026-02-10": [
            {"production_day": 1, "items": [parent_item, second_solid]},
        ],
    }

    commit_plan_plates(
        plan_id="plan_kp_plate_id",
        orders_2d=orders,
        optimization_result=optimization_result,
        all_tracks_list=[],
        db_path=tmp_db,
        tracks_by_day=tracks_by_day,
    )

    assert parent_item.get("kp_plate_id") is not None
    assert second_solid.get("kp_plate_id") is not None
    assert secondary_item.get("kp_plate_id") is not None
    assert parent_item["kp_plate_id"] != secondary_item["kp_plate_id"]


def test_commit_warns_when_legacy_branch_used_with_tracks_by_day(tmp_db, caplog):
    """P9: если ``tracks_by_day`` передан, но у item нет identity и
    backfill не справился — commit пишет WARNING с identity."""
    import logging

    _seed_kp_plate(tmp_db, kp_id=1, plate_name="ПБ 60-12-8п", qty=2)
    orders = [{"kp_id": 1, "plate_name": "ПБ 60-12-8п", "qty": 2, "load_code": 8}]
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "ПБ 60-12-8п"},
            {"source": "primary", "kp_id": 1, "plate_name": "ПБ 60-12-8п"},
        ],
    }
    tracks_by_day = {
        "2026-03-01": [
            {
                "production_day": 1,
                "items": [
                    {"length": 6.0, "mode": "solid", "width": 1.2,
                     "load_code": 8, "kp_id": None, "plate_name": None},
                    {"length": 6.0, "mode": "solid", "width": 1.2,
                     "load_code": 8, "kp_id": None, "plate_name": None},
                ],
            }
        ],
    }

    with caplog.at_level(logging.WARNING, logger="core.plan_commit"):
        commit_plan_plates(
            plan_id="plan_legacy_warn",
            orders_2d=orders,
            optimization_result=optimization_result,
            all_tracks_list=[],
            db_path=tmp_db,
            tracks_by_day=tracks_by_day,
        )

    legacy_warnings = [
        r for r in caplog.records
        if r.levelno == logging.WARNING and "tracks_by_day" in r.getMessage()
    ]
    assert legacy_warnings, (
        "ожидался WARNING о том, что у items нет identity, "
        f"но в логах: {[r.getMessage() for r in caplog.records]}"
    )


def test_commit_plan_plates_duplicate_order_lines_same_identity_no_orphan(tmp_db):
    """Две строки orders_2d с одной identity не должны порождать orphan kp_plates.

    Регрессия H1: day-allocation без mutual budget повторно помечала ранние дни.
    """
    plate = "ПБ 60-12-8п"
    _seed_kp_plate(tmp_db, kp_id=1, plate_name=plate, qty=4)
    orders = [
        {"kp_id": 1, "plate_name": plate, "qty": 2, "load_code": 8},
        {"kp_id": 1, "plate_name": plate, "qty": 2, "load_code": 8},
    ]
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": plate}
            for _ in range(4)
        ],
    }
    items_day1 = [
        {
            "length": 6.0, "mode": "solid", "width": 1.2, "load_code": 8,
            "kp_id": 1, "plate_name": plate,
        },
        {
            "length": 6.0, "mode": "solid", "width": 1.2, "load_code": 8,
            "kp_id": 1, "plate_name": plate,
        },
    ]
    items_day2 = [
        {
            "length": 6.0, "mode": "solid", "width": 1.2, "load_code": 8,
            "kp_id": 1, "plate_name": plate,
        },
        {
            "length": 6.0, "mode": "solid", "width": 1.2, "load_code": 8,
            "kp_id": 1, "plate_name": plate,
        },
    ]
    tracks_by_day = {
        "2026-04-01": [{"production_day": 1, "items": items_day1}],
        "2026-04-02": [{"production_day": 2, "items": items_day2}],
    }

    commit_plan_plates(
        plan_id="plan_dup_identity",
        orders_2d=orders,
        optimization_result=optimization_result,
        all_tracks_list=[],
        db_path=tmp_db,
        tracks_by_day=tracks_by_day,
    )

    all_items = items_day1 + items_day2
    assert all(it.get("kp_plate_id") is not None for it in all_items), (
        f"все items должны получить kp_plate_id, получили: "
        f"{[it.get('kp_plate_id') for it in all_items]}"
    )

    with sqlite3.connect(tmp_db) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, qty, day_number FROM kp_plates "
            "WHERE plan_id = ? AND status = 'в плане' ORDER BY id",
            ("plan_dup_identity",),
        )
        rows = cur.fetchall()
        cur.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM kp_plates "
            "WHERE plan_id = ? AND status = 'в плане'",
            ("plan_dup_identity",),
        )
        total_qty = int(cur.fetchone()[0])

    assert total_qty == 4, f"SUM(qty) в плане должен быть 4, получили {total_qty}"
    assert rows, "ожидаем строки kp_plates со status='в плане'"

    refs_by_id: dict[int, int] = {}
    for it in all_items:
        pid = int(it["kp_plate_id"])
        refs_by_id[pid] = refs_by_id.get(pid, 0) + 1

    for plate_id, qty, day_number in rows:
        assert day_number is not None, f"orphan/undated id={plate_id}"
        refs = refs_by_id.get(int(plate_id), 0)
        assert refs >= 1, (
            f"unreferenced kp_plates id={plate_id} qty={qty} day={day_number}"
        )
        assert refs == int(qty), (
            f"qty_mismatch id={plate_id}: refs={refs} qty={qty} day={day_number}"
        )


def test_commit_plan_plates_leftover_pool_raises_and_rolls_back(tmp_db, monkeypatch):
    """Leftover в пуле после link → PlanCommitError и откат «в плане»."""
    plate = "ПБ 60-12-8п"
    _seed_kp_plate(tmp_db, kp_id=1, plate_name=plate, qty=2)
    orders = [{"kp_id": 1, "plate_name": plate, "qty": 2, "load_code": 8}]
    optimization_result = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": plate},
            {"source": "primary", "kp_id": 1, "plate_name": plate},
        ],
    }
    tracks_by_day = {
        "2026-05-01": [
            {
                "production_day": 1,
                "items": [
                    {
                        "length": 6.0, "mode": "solid", "width": 1.2,
                        "load_code": 8, "kp_id": 1, "plate_name": plate,
                    },
                    {
                        "length": 6.0, "mode": "solid", "width": 1.2,
                        "load_code": 8, "kp_id": 1, "plate_name": plate,
                    },
                ],
            },
        ],
    }

    real_mark = kp_db.mark_plates_as_planned
    rollback_calls: list[str] = []
    real_return = kp_db.return_plan_plates_to_production

    def inflate_mark(*, kp_id, plate_name, qty_to_plan, plan_id, db_path, day_number=None):
        result = real_mark(
            kp_id=kp_id,
            plate_name=plate_name,
            qty_to_plan=qty_to_plan,
            plan_id=plan_id,
            db_path=db_path,
            day_number=day_number,
        )
        pairs = [list(p) for p in (result.get("id_qty_pairs") or [])]
        if pairs:
            pairs[0][1] = int(pairs[0][1]) + 1
        out = dict(result)
        out["id_qty_pairs"] = pairs
        return out

    def spy_return(plan_id, db_path):
        rollback_calls.append(plan_id)
        return real_return(plan_id, db_path)

    monkeypatch.setattr("core.kp_db_plates.mark_plates_as_planned", inflate_mark)
    monkeypatch.setattr("core.kp_db_plates.return_plan_plates_to_production", spy_return)

    with pytest.raises(PlanCommitError, match="непокрытые слоты пула"):
        commit_plan_plates(
            plan_id="plan_leftover_pool",
            orders_2d=orders,
            optimization_result=optimization_result,
            all_tracks_list=[],
            db_path=tmp_db,
            tracks_by_day=tracks_by_day,
        )

    assert "plan_leftover_pool" in rollback_calls
    statuses = _fetch_status(tmp_db, 1, plate)
    assert statuses, "ожидаем строки kp_plates после отката"
    assert all(status == "в производстве" for status, _qty, _plan in statuses)
    assert all(plan_id is None for _status, _qty, plan_id in statuses)


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

    monkeypatch.setattr("core.kp_db_plates.mark_plates_as_planned", fake_mark)
    monkeypatch.setattr("core.kp_db_plates.return_plan_plates_to_production", spy_return)

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


def _insert_active_promise(db_path: str, *, kp_id: int, week_start: str, tracks: int) -> int:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO kp_promise (
                kp_id, tracks_total, promised_date, kind, status,
                created_by, created_at, expires_at
            ) VALUES (?, ?, '2026-09-25', 'promise', 'active', 'alice',
                      '2026-09-03T12:00:00', NULL)
            """,
            (kp_id, tracks),
        )
        promise_id = int(cur.lastrowid)
        cur.execute(
            """
            INSERT INTO kp_promise_alloc (promise_id, week_start, tracks, status)
            VALUES (?, ?, ?, 'active')
            """,
            (promise_id, week_start, tracks),
        )
        conn.commit()
    return promise_id


def test_commit_plan_plates_consumes_entered_and_overdues_missed(tmp_db):
    """T8: коммит гасит аллокацию вошедшего КП; невошедшее → overdue, не блок."""
    from datetime import date

    from app.repositories.promise_repository import PromiseRepository
    from app.services.promise_service import PromiseService

    week = "2026-09-07"
    _seed_kp_plate(tmp_db, kp_id=1, plate_name="A", qty=1)
    _seed_kp_plate(tmp_db, kp_id=2, plate_name="B", qty=1)
    entered_id = _insert_active_promise(tmp_db, kp_id=1, week_start=week, tracks=4)
    missed_id = _insert_active_promise(tmp_db, kp_id=2, week_start=week, tracks=3)

    tracks_by_day = {
        week: [
            {
                "production_day": 1,
                "items": [
                    {"kp_id": 1, "plate_name": "A", "length": 6.0},
                ],
            }
        ]
    }
    commit_plan_plates(
        plan_id="plan_promise_settle",
        orders_2d=[{"kp_id": 1, "plate_name": "A", "qty": 1, "load_code": 8}],
        optimization_result={
            "plate_assignments": [
                {"source": "primary", "kp_id": 1, "plate_name": "A"},
            ],
        },
        all_tracks_list=[],
        db_path=tmp_db,
        tracks_by_day=tracks_by_day,
        settle_fn=PromiseService(db_path=tmp_db).settle_plan_commit,
    )

    with sqlite3.connect(tmp_db) as conn:
        entered = conn.execute(
            "SELECT p.status, a.status FROM kp_promise p "
            "JOIN kp_promise_alloc a ON a.promise_id = p.id WHERE p.id = ?",
            (entered_id,),
        ).fetchone()
        missed = conn.execute(
            "SELECT p.status, a.status FROM kp_promise p "
            "JOIN kp_promise_alloc a ON a.promise_id = p.id WHERE p.id = ?",
            (missed_id,),
        ).fetchone()

    assert entered == ("consumed", "consumed")
    assert missed == ("active", "overdue")
    assert PromiseRepository(db_path=tmp_db).sum_promised_by_week() == {}
    plates = _fetch_status(tmp_db, 1, "A")
    assert plates == [("в плане", 1, "plan_promise_settle")]
    leftover = _fetch_status(tmp_db, 2, "B")
    assert leftover == [("в производстве", 1, None)]
    assert date.fromisoformat(week).weekday() == 0
