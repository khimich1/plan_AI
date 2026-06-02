"""Тесты удаления дорожки из плана (remove_track_from_plan).

Сценарии:
1. happy path — возврат плит в «в производстве», день удаляется из JSON;
2. re-plan smoke — после удаления build_plan снова берёт те же плиты;
3. completed day — TrackRemovalError day_already_completed, без изменений;
4. wrong plan_id — строка kp_plates с чужим plan_id не трогается;
5. secondary_cuts — возвращаются обе физические единицы;
6. два плана на одну дату — изоляция;
7. saved_tracks_count синхронизирован с len(tracks).

Дополнительно: unit-тесты collect_plate_returns_from_track и return_plate_rows_for_plan.
"""
from __future__ import annotations

import sqlite3
import sys
from collections import Counter
from copy import deepcopy
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from app.planning import plan_manager
from core import kp_db
from core.plan_track_removal import (
    TrackRemovalError,
    collect_plate_returns_from_track,
)

PLATE_NAME = "ПБ 60-12-8п"
SECONDARY_PLATE_NAME = "ПБ 60-3,2-8п"
KP_ID = 1
DATE_KEY = "2026-04-21"
OTHER_PLAN_ID = "plan_other_test"


# ---------------------------------------------------------------------------
# Fixtures (reuse patterns from tests/test_plan_consistency.py)
# ---------------------------------------------------------------------------


@pytest.fixture
def tmp_plita(tmp_path) -> str:
    db_path = str(tmp_path / "plita.db")
    kp_db.init_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date, execution_terms, customer_name) "
            "VALUES (?, '2026-01-01', '21.04.2026', 'ТестКлиент')",
            (KP_ID,),
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status) VALUES (?, 'в работе')",
            (KP_ID,),
        )
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (?, 1, ?, 6.0, 1.2, 800, 3, 'в производстве')
            """,
            (KP_ID, PLATE_NAME),
        )
        conn.commit()
    return db_path


@pytest.fixture(autouse=True)
def _isolate_plans(tmp_path, monkeypatch):
    plans_dir = tmp_path / "plans"
    plans_dir.mkdir()
    metadata_path = tmp_path / "plans_metadata.json"
    monkeypatch.setattr(plan_manager, "PLANS_DIR", plans_dir)
    monkeypatch.setattr(plan_manager, "PLANS_METADATA_PATH", metadata_path)
    monkeypatch.setattr(plan_manager, "load_holidays", lambda: set())
    monkeypatch.setattr(plan_manager, "load_extra_workdays", lambda: set())


@pytest.fixture
def planning_service(tmp_plita, monkeypatch):
    """Сервис с мокированным оптимизатором: один трек с N плитами."""
    from app.services.production_planning_service import ProductionPlanningService

    service = ProductionPlanningService(
        plita_db_path=tmp_plita,
        pb_db_path=tmp_plita,
    )

    def fake_optimize(self, *, orders_2d, **kwargs):
        if not orders_2d:
            return [], {}
        order = orders_2d[0]
        items = [
            {
                "kp_id": order["kp_id"],
                "plate_name": order["plate_name"],
                "length": order["length"],
                "width": order["width"],
                "load_code": order["load_code"],
            }
            for _ in range(int(order.get("qty") or 0))
        ]
        tracks = [{"label": "ОСНОВНАЯ", "items": items}]
        optimization_result = {
            "total_plates": len(items),
            "plate_assignments": [
                {
                    "source": "primary",
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"],
                    "load_code": order["load_code"],
                }
                for _ in items
            ],
        }
        return tracks, optimization_result

    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        fake_optimize,
    )
    monkeypatch.setattr(
        "app.services.production_planning_service.get_reinforcement",
        lambda **kwargs: 999.0,
    )
    return service


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _build_single_track_plan(planning_service, *, tracks_count: int = 1) -> dict:
    return planning_service.build_plan(
        start_date=DATE_KEY,
        tracks_count=tracks_count,
        filter_method="all",
    )


def _kp_plate_rows(
    db_path: str,
    *,
    plan_id: str | None = None,
    status: str | None = None,
) -> list[tuple]:
    query = "SELECT id, kp_id, plate_name, qty, status, plan_id, day_number FROM kp_plates WHERE 1=1"
    params: list = []
    if plan_id is not None:
        query += " AND plan_id = ?"
        params.append(plan_id)
    if status is not None:
        query += " AND status = ?"
        params.append(status)
    query += " ORDER BY id"
    with sqlite3.connect(db_path) as conn:
        return conn.execute(query, params).fetchall()


def _total_qty(rows: list[tuple]) -> int:
    return sum(int(r[3]) for r in rows)


def _extract_kp_plate_ids(plan: dict, date_key: str, track_index: int = 0) -> list[int]:
    from core.plan_commit import _iter_physical_items

    track = plan["days"][date_key]["tracks"][track_index]
    ids: list[int] = []
    for physical in _iter_physical_items(track.get("items")):
        pid = physical.get("kp_plate_id")
        if pid is not None:
            ids.append(int(pid))
    return ids


def _snapshot_kp_plates(db_path: str) -> list[tuple]:
    with sqlite3.connect(db_path) as conn:
        return conn.execute(
            "SELECT id, qty, status, plan_id, day_number FROM kp_plates ORDER BY id"
        ).fetchall()


def _snapshot_plan_file(plan_id: str) -> str:
    path = plan_manager.get_plan_path(plan_id)
    return path.read_text(encoding="utf-8")


def _planning_service_with_secondary(tmp_plita, monkeypatch):
    """Planning service с треком primary + secondary_cuts (как в test_plan_consistency)."""
    from app.services.production_planning_service import ProductionPlanningService

    with sqlite3.connect(tmp_plita) as conn:
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (?, 2, ?, 6.0, 0.32, 800, 2, 'в производстве')
            """,
            (KP_ID, SECONDARY_PLATE_NAME),
        )
        conn.commit()

    service = ProductionPlanningService(
        plita_db_path=tmp_plita,
        pb_db_path=tmp_plita,
    )

    def fake_optimize_with_secondary(self, *, orders_2d, **kwargs):
        if not orders_2d:
            return [], {}
        primary_order = next(
            (o for o in orders_2d if int(round(float(o["width"]))) == 1200),
            None,
        )
        secondary_order = next(
            (o for o in orders_2d if int(round(float(o["width"]))) == 320),
            None,
        )
        assert primary_order and secondary_order

        items: list[dict] = []
        secondary_slots = int(secondary_order["qty"])
        for idx in range(int(primary_order["qty"])):
            item = {
                "length": primary_order["length"],
                "mode": "split",
                "main_w": 1.2,
                "rest_w": 0.32,
                "load_code": primary_order["load_code"],
                "kp_id": primary_order["kp_id"],
                "plate_name": primary_order["plate_name"],
            }
            if idx < secondary_slots:
                item["secondary_cuts"] = [
                    {
                        "width": 0.32,
                        "label": "[2] secondary без target_length",
                        "load_code": primary_order["load_code"],
                    }
                ]
            items.append(item)
        if secondary_order["qty"] > primary_order["qty"]:
            for _ in range(secondary_order["qty"] - primary_order["qty"]):
                items.append({
                    "length": primary_order["length"],
                    "mode": "split",
                    "main_w": 1.2,
                    "rest_w": 0.32,
                    "load_code": primary_order["load_code"],
                    "kp_id": primary_order["kp_id"],
                    "plate_name": primary_order["plate_name"],
                    "secondary_cuts": [
                        {
                            "width": 0.32,
                            "label": "[2] secondary без target_length",
                            "load_code": primary_order["load_code"],
                        }
                    ],
                })

        tracks = [{"label": "ОСНОВНАЯ", "items": items}]
        plate_assignments: list[dict] = []
        for _ in range(int(primary_order["qty"])):
            plate_assignments.append({
                "source": "primary",
                "kp_id": primary_order["kp_id"],
                "plate_name": primary_order["plate_name"],
                "length": primary_order["length"],
                "width": primary_order["width"],
                "load_code": primary_order["load_code"],
            })
        for _ in range(int(secondary_order["qty"])):
            plate_assignments.append({
                "source": "secondary",
                "kp_id": secondary_order["kp_id"],
                "plate_name": secondary_order["plate_name"],
                "length": secondary_order["length"],
                "width": secondary_order["width"],
                "load_code": secondary_order["load_code"],
            })

        optimization_result = {
            "total_plates": len(plate_assignments),
            "plate_assignments": plate_assignments,
        }

        from core.plate_attribution import backfill_track_items_identity
        backfill_track_items_identity(tracks, orders_2d)

        return tracks, optimization_result

    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        fake_optimize_with_secondary,
    )
    monkeypatch.setattr(
        "app.services.production_planning_service.get_reinforcement",
        lambda **kwargs: 999.0,
    )
    return service


def _planning_service_two_tracks(tmp_plita, monkeypatch):
    """Planning service: два трека на одном дне (2 + 1 плита)."""
    from app.services.production_planning_service import ProductionPlanningService

    service = ProductionPlanningService(
        plita_db_path=tmp_plita,
        pb_db_path=tmp_plita,
    )

    def fake_optimize_two_tracks(self, *, orders_2d, **kwargs):
        if not orders_2d:
            return [], {}
        order = orders_2d[0]
        items = [
            {
                "kp_id": order["kp_id"],
                "plate_name": order["plate_name"],
                "length": order["length"],
                "width": order["width"],
                "load_code": order["load_code"],
            }
            for _ in range(int(order.get("qty") or 0))
        ]
        tracks = [
            {"label": "T1", "items": items[:2]},
            {"label": "T2", "items": items[2:]},
        ]
        optimization_result = {
            "total_plates": len(items),
            "plate_assignments": [
                {
                    "source": "primary",
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"],
                    "load_code": order["load_code"],
                }
                for _ in items
            ],
        }
        return tracks, optimization_result

    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        fake_optimize_two_tracks,
    )
    monkeypatch.setattr(
        "app.services.production_planning_service.get_reinforcement",
        lambda **kwargs: 999.0,
    )
    return service


# ---------------------------------------------------------------------------
# 1. Happy path
# ---------------------------------------------------------------------------


def test_remove_track_happy_path(planning_service, tmp_plita):
    """1 трек, 3 плиты → удаление → JSON без трека, БД «в производстве», qty сохранён."""
    built = _build_single_track_plan(planning_service, tracks_count=1)
    plan_id = built["plan"]["id"]

    plan = plan_manager.load_plan(plan_id)
    assert len(plan["days"][DATE_KEY]["tracks"]) == 1
    plate_ids = _extract_kp_plate_ids(plan, DATE_KEY, 0)
    assert len(plate_ids) == 3

    in_plan_before = _kp_plate_rows(tmp_plita, plan_id=plan_id, status="в плане")
    assert _total_qty(in_plan_before) == 3

    result = plan_manager.remove_track_from_plan(
        plan_id,
        DATE_KEY,
        0,
        db_path=tmp_plita,
    )

    assert result["plates_returned"] == 3
    assert result["saved_tracks_count"] == 0

    plan_after = plan_manager.load_plan(plan_id)
    assert DATE_KEY not in plan_after.get("days", {})

    production_rows = _kp_plate_rows(tmp_plita, status="в производстве")
    assert _total_qty(production_rows) == 3
    assert not _kp_plate_rows(tmp_plita, plan_id=plan_id, status="в плане")


# ---------------------------------------------------------------------------
# 2. Re-plan smoke
# ---------------------------------------------------------------------------


def test_replan_after_track_removal(planning_service, tmp_plita):
    """После удаления дорожки build_plan снова подбирает те же 3 плиты."""
    built = _build_single_track_plan(planning_service, tracks_count=1)
    plan_id = built["plan"]["id"]

    plan_manager.remove_track_from_plan(
        plan_id, DATE_KEY, 0, db_path=tmp_plita
    )

    rebuilt = _build_single_track_plan(planning_service, tracks_count=1)
    new_plan_id = rebuilt["plan"]["id"]

    in_plan = _kp_plate_rows(tmp_plita, plan_id=new_plan_id, status="в плане")
    assert _total_qty(in_plan) == 3

    new_plan = plan_manager.load_plan(new_plan_id)
    assert len(new_plan["days"][DATE_KEY]["tracks"]) == 1
    assert len(_extract_kp_plate_ids(new_plan, DATE_KEY, 0)) == 3


# ---------------------------------------------------------------------------
# 3. Completed day
# ---------------------------------------------------------------------------


def test_remove_track_completed_day_raises(planning_service, tmp_plita):
    """day.completed=True → TrackRemovalError day_already_completed, без изменений."""
    built = _build_single_track_plan(planning_service, tracks_count=1)
    plan_id = built["plan"]["id"]

    plan = plan_manager.load_plan(plan_id)
    plan["days"][DATE_KEY]["completed"] = True
    plan_manager.save_plan(plan)

    snap_db = _snapshot_kp_plates(tmp_plita)
    snap_json = _snapshot_plan_file(plan_id)

    with pytest.raises(TrackRemovalError) as exc_info:
        plan_manager.remove_track_from_plan(
            plan_id, DATE_KEY, 0, db_path=tmp_plita
        )

    assert exc_info.value.code == "day_already_completed"
    assert _snapshot_kp_plates(tmp_plita) == snap_db
    assert _snapshot_plan_file(plan_id) == snap_json


# ---------------------------------------------------------------------------
# 4. Wrong plan_id on row
# ---------------------------------------------------------------------------


def test_wrong_plan_id_row_not_touched(planning_service, tmp_plita):
    """Чужой kp_plate_id → incomplete_return, JSON и БД плана без изменений."""
    built = _build_single_track_plan(planning_service, tracks_count=1)
    plan_id = built["plan"]["id"]

    # Строка, принадлежащая «чужому» плану.
    with sqlite3.connect(tmp_plita) as conn:
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status, plan_id, day_number
            ) VALUES (?, 99, ?, 6.0, 1.2, 800, 1, 'в плане', ?, 1)
            """,
            (KP_ID, PLATE_NAME, OTHER_PLAN_ID),
        )
        foreign_plate_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.commit()

    foreign_before = _kp_plate_rows(tmp_plita, plan_id=OTHER_PLAN_ID, status="в плане")
    assert len(foreign_before) == 1

    # Подменяем kp_plate_id первого item на чужой id.
    plan = plan_manager.load_plan(plan_id)
    track = plan["days"][DATE_KEY]["tracks"][0]
    track["items"][0]["kp_plate_id"] = foreign_plate_id
    plan_manager.save_plan(plan)

    snap_db = _snapshot_kp_plates(tmp_plita)
    snap_json = _snapshot_plan_file(plan_id)

    with pytest.raises(TrackRemovalError) as exc_info:
        plan_manager.remove_track_from_plan(
            plan_id, DATE_KEY, 0, db_path=tmp_plita
        )

    assert exc_info.value.code == "incomplete_return"
    assert _snapshot_kp_plates(tmp_plita) == snap_db
    assert _snapshot_plan_file(plan_id) == snap_json

    foreign_after = _kp_plate_rows(tmp_plita, plan_id=OTHER_PLAN_ID, status="в плане")
    assert foreign_after == foreign_before


# ---------------------------------------------------------------------------
# 5. Secondary cuts
# ---------------------------------------------------------------------------


def test_secondary_cuts_both_units_returned(tmp_plita, monkeypatch):
    """Трек с secondary_cuts возвращает обе физические единицы (primary + secondary)."""
    service = _planning_service_with_secondary(tmp_plita, monkeypatch)
    built = service.build_plan(
        start_date=DATE_KEY, tracks_count=1, filter_method="all"
    )
    plan_id = built["plan"]["id"]

    plan = plan_manager.load_plan(plan_id)
    plate_ids = _extract_kp_plate_ids(plan, DATE_KEY, 0)
    assert len(plate_ids) >= 2, "ожидаются kp_plate_id и у primary, и у secondary"

    in_plan_before = _kp_plate_rows(tmp_plita, plan_id=plan_id, status="в плане")
    expected_qty = _total_qty(in_plan_before)
    assert expected_qty == 5  # 3 primary + 2 secondary (как в test_plan_consistency)

    result = plan_manager.remove_track_from_plan(
        plan_id, DATE_KEY, 0, db_path=tmp_plita
    )

    assert result["plates_returned"] == expected_qty
    production_rows = _kp_plate_rows(tmp_plita, status="в производстве")
    assert _total_qty(production_rows) == 5


# ---------------------------------------------------------------------------
# 6. Two plans same date
# ---------------------------------------------------------------------------


def test_two_plans_same_date_isolation(planning_service, tmp_plita, monkeypatch):
    """Удаление трека в плане A не затрагивает треки плана B на ту же дату."""
    plan_ids = iter(["plan_a_isolated", "plan_b_isolated"])
    monkeypatch.setattr(plan_manager, "create_plan_id", lambda: next(plan_ids))

    built_a = _build_single_track_plan(planning_service, tracks_count=1)
    plan_a_id = built_a["plan"]["id"]
    assert plan_a_id == "plan_a_isolated"

    # Плиты первого плана уже «в плане» — добавляем ещё одну партию для плана B.
    with sqlite3.connect(tmp_plita) as conn:
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (?, 4, ?, 6.0, 1.2, 800, 3, 'в производстве')
            """,
            (KP_ID, PLATE_NAME),
        )
        conn.commit()

    built_b = _build_single_track_plan(planning_service, tracks_count=1)
    plan_b_id = built_b["plan"]["id"]
    assert plan_b_id == "plan_b_isolated"

    plan_b_before = plan_manager.load_plan(plan_b_id)
    tracks_b_before = deepcopy(plan_b_before["days"][DATE_KEY]["tracks"])
    snap_b_db = _snapshot_kp_plates(tmp_plita)

    plan_manager.remove_track_from_plan(
        plan_a_id, DATE_KEY, 0, db_path=tmp_plita
    )

    plan_a_after = plan_manager.load_plan(plan_a_id)
    assert DATE_KEY not in plan_a_after.get("days", {})

    plan_b_after = plan_manager.load_plan(plan_b_id)
    assert plan_b_after["days"][DATE_KEY]["tracks"] == tracks_b_before

    # Строки плана B в БД не изменились.
    rows_b = _kp_plate_rows(tmp_plita, plan_id=plan_b_id, status="в плане")
    assert _total_qty(rows_b) == 3

    rows_a = _kp_plate_rows(tmp_plita, plan_id=plan_a_id, status="в плане")
    assert rows_a == []


# ---------------------------------------------------------------------------
# 7. saved_tracks_count sync
# ---------------------------------------------------------------------------


def test_saved_tracks_count_sync_after_removal(tmp_plita, monkeypatch):
    """После удаления одного из двух треков count_day_tracks == saved_tracks_count."""
    service = _planning_service_two_tracks(tmp_plita, monkeypatch)
    built = service.build_plan(
        start_date=DATE_KEY, tracks_count=2, filter_method="all"
    )
    plan_id = built["plan"]["id"]

    plan = plan_manager.load_plan(plan_id)
    day = plan["days"][DATE_KEY]
    assert len(day["tracks"]) == 2
    assert day["saved_tracks_count"] == 2
    assert plan_manager.count_day_tracks(day) == 2

    plan_manager.remove_track_from_plan(
        plan_id, DATE_KEY, 0, db_path=tmp_plita
    )

    plan_after = plan_manager.load_plan(plan_id)
    day_after = plan_after["days"][DATE_KEY]
    assert len(day_after["tracks"]) == 1
    assert day_after["saved_tracks_count"] == 1
    assert plan_manager.count_day_tracks(day_after) == day_after["saved_tracks_count"]


# ---------------------------------------------------------------------------
# Unit: collect_plate_returns_from_track
# ---------------------------------------------------------------------------


def test_collect_plate_returns_from_track_by_kp_plate_id():
    """Считает qty по kp_plate_id, включая secondary_cuts."""
    track = {
        "items": [
            {"kp_plate_id": 10},
            {"kp_plate_id": 10},
            {
                "kp_plate_id": 20,
                "secondary_cuts": [{"kp_plate_id": 30}],
            },
        ]
    }
    id_qty, legacy_qty = collect_plate_returns_from_track(track)

    assert id_qty == Counter({10: 2, 20: 1, 30: 1})
    assert legacy_qty == Counter()


def test_collect_plate_returns_from_track_legacy_identity():
    """Без kp_plate_id — legacy-счётчик (kp_id, canonical plate_name)."""
    track = {
        "items": [
            {"kp_id": KP_ID, "plate_name": "Плиты ПБ 60-12-8п"},
            {
                "kp_id": KP_ID,
                "plate_name": PLATE_NAME,
                "secondary_cuts": [
                    {"kp_id": KP_ID, "plate_name": SECONDARY_PLATE_NAME},
                ],
            },
        ]
    }
    id_qty, legacy_qty = collect_plate_returns_from_track(track)

    assert id_qty == Counter()
    assert legacy_qty[(KP_ID, PLATE_NAME)] == 2
    assert legacy_qty[(KP_ID, SECONDARY_PLATE_NAME)] == 1


def test_collect_plate_returns_from_track_empty_track():
    """Пустая дорожка — пустые счётчики."""
    id_qty, legacy_qty = collect_plate_returns_from_track({"items": []})
    assert id_qty == Counter()
    assert legacy_qty == Counter()


def test_remove_track_no_plate_identity(planning_service, tmp_plita):
    """Дорожка без идентичностей → no_plate_identity, без изменений БД/JSON."""
    built = _build_single_track_plan(planning_service, tracks_count=1)
    plan_id = built["plan"]["id"]

    plan = plan_manager.load_plan(plan_id)
    plan["days"][DATE_KEY]["tracks"][0] = {
        "label": "пустая",
        "items": [{"length": 6.0, "width": 1.2}],
    }
    plan_manager.save_plan(plan)

    snap_db = _snapshot_kp_plates(tmp_plita)
    snap_json = _snapshot_plan_file(plan_id)

    with pytest.raises(TrackRemovalError) as exc_info:
        plan_manager.remove_track_from_plan(
            plan_id, DATE_KEY, 0, db_path=tmp_plita
        )

    assert exc_info.value.code == "no_plate_identity"
    assert _snapshot_kp_plates(tmp_plita) == snap_db
    assert _snapshot_plan_file(plan_id) == snap_json


# ---------------------------------------------------------------------------
# Unit: return_plate_rows_for_plan basics
# ---------------------------------------------------------------------------


def test_return_plate_rows_for_plan_happy_path(tmp_plita):
    """Прямой вызов: возвращает плиты в «в производстве», qty сохраняется."""
    with sqlite3.connect(tmp_plita) as conn:
        conn.execute(
            """
            UPDATE kp_plates
            SET status = 'в плане', plan_id = ?, day_number = 1
            WHERE kp_id = ?
            """,
            ("plan_direct", KP_ID),
        )
        row_id = conn.execute(
            "SELECT id FROM kp_plates WHERE kp_id = ?", (KP_ID,)
        ).fetchone()[0]
        conn.commit()

    result = kp_db.return_plate_rows_for_plan(
        "plan_direct",
        Counter({row_id: 3}),
        tmp_plita,
    )

    assert result["plates_returned"] == 3
    assert not result["warnings"]

    rows = _kp_plate_rows(tmp_plita, status="в производстве")
    assert _total_qty(rows) == 3
    assert rows[0][5] is None  # plan_id NULL


def test_return_plate_rows_for_plan_wrong_plan_id_not_touched(tmp_plita):
    """Строка с другим plan_id не возвращается."""
    with sqlite3.connect(tmp_plita) as conn:
        conn.execute(
            """
            UPDATE kp_plates
            SET status = 'в плане', plan_id = ?, day_number = 1
            WHERE kp_id = ?
            """,
            ("plan_owner", KP_ID),
        )
        row_id = conn.execute(
            "SELECT id FROM kp_plates WHERE kp_id = ?", (KP_ID,)
        ).fetchone()[0]
        conn.commit()

    snap = _snapshot_kp_plates(tmp_plita)

    result = kp_db.return_plate_rows_for_plan(
        "plan_intruder",
        Counter({row_id: 1}),
        tmp_plita,
    )

    assert result["plates_returned"] == 0
    assert any("не найдена" in w for w in result["warnings"])
    assert _snapshot_kp_plates(tmp_plita) == snap


def test_return_plate_rows_for_plan_partial_qty(tmp_plita):
    """Частичный возврат: qty в плане уменьшается, остаток — «в производстве»."""
    with sqlite3.connect(tmp_plita) as conn:
        conn.execute(
            """
            UPDATE kp_plates
            SET status = 'в плане', plan_id = ?, day_number = 1
            WHERE kp_id = ?
            """,
            ("plan_partial", KP_ID),
        )
        row_id = conn.execute(
            "SELECT id FROM kp_plates WHERE kp_id = ?", (KP_ID,)
        ).fetchone()[0]
        conn.commit()

    result = kp_db.return_plate_rows_for_plan(
        "plan_partial",
        Counter({row_id: 1}),
        tmp_plita,
    )

    assert result["plates_returned"] == 1

    in_plan = _kp_plate_rows(tmp_plita, plan_id="plan_partial", status="в плане")
    in_production = _kp_plate_rows(tmp_plita, status="в производстве")
    assert _total_qty(in_plan) == 2
    assert _total_qty(in_production) == 1
