"""Тесты для журнала переходов статусов плит ``plate_status_log``.

Проверяют:
1. ``mark_plates_as_planned`` пишет запись ``в производстве → в плане``.
2. После завершения дня с браком в логе видны и ``completed``, и ``rejected``.
3. Audit-вставка атомарна с основным UPDATE: если INSERT в журнал падает —
   откатывается и сама пометка плиты.
"""
from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from app.planning import plan_manager
from core import kp_db


PLATE_NAME = "ПБ 60-12-8п"
KP_ID = 1


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
        "core.production.planning.get_reinforcement",
        lambda **kwargs: 999.0,
    )
    return service


def _audit_rows(db_path: str) -> list[dict]:
    """Возвращает список записей plate_status_log (для KP_ID) в порядке вставки."""
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT plate_id, kp_id, plate_name, plan_id, day_number,
                   from_status, to_status, qty, reason, actor
            FROM plate_status_log
            WHERE kp_id = ?
            ORDER BY id
            """,
            (KP_ID,),
        ).fetchall()
    return [dict(r) for r in rows]


def test_audit_log_records_planning(planning_service, tmp_plita):
    """build_plan → одна запись 'в производстве' → 'в плане' с qty=3."""
    planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )

    rows = _audit_rows(tmp_plita)
    assert len(rows) == 1, f"Ожидалась 1 запись, получили {len(rows)}: {rows}"

    entry = rows[0]
    assert entry["from_status"] == "в производстве"
    assert entry["to_status"] == "в плане"
    assert entry["qty"] == 3
    assert entry["reason"] == "planned"
    assert entry["plate_name"] == PLATE_NAME
    assert entry["plan_id"] is not None


def test_audit_log_records_completion_and_rejection(planning_service, tmp_plita):
    """complete_day(reject=1, complete=2) → две записи в логе."""
    from app.repositories.kp_repository import KpRepository
    from app.repositories.plan_repository import PlanRepository
    from app.services.production_service import ProductionService

    built = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    service = ProductionService(
        kp_repository=KpRepository(db_path=tmp_plita),
        plan_repository=PlanRepository(db_path=tmp_plita),
        planning_service=planning_service,
    )
    service.complete_day(
        plan_id=plan_id,
        target_date="2026-04-21",
        rejected_plates=[{"track_number": 1, "plate_index": 0, "qty": 1}],
        actor="test:user@example.com",
    )

    rows = _audit_rows(tmp_plita)
    # 1× planned (build_plan), 1× completed (move_plates_to_completed),
    # 1× rejected (return_plates_to_production)
    assert len(rows) == 3, f"Ожидались 3 записи, получили {len(rows)}: {rows}"

    by_reason = {r["reason"]: r for r in rows}
    assert "planned" in by_reason
    assert "completed" in by_reason
    assert "rejected" in by_reason

    completed_entry = by_reason["completed"]
    assert completed_entry["from_status"] == "в плане"
    assert completed_entry["to_status"] == "completed"
    assert completed_entry["qty"] == 2
    assert completed_entry["actor"] == "test:user@example.com"

    rejected_entry = by_reason["rejected"]
    assert rejected_entry["from_status"] == "в плане"
    assert rejected_entry["to_status"] == "в производстве"
    assert rejected_entry["qty"] == 1
    assert rejected_entry["actor"] == "test:user@example.com"


def test_audit_atomic_with_kp_plates_update(tmp_plita, monkeypatch):
    """Если INSERT в plate_status_log падает — UPDATE kp_plates тоже откатывается."""
    import core.kp_db_plates_planning as planning_mod

    original_audit = planning_mod.audit_append

    def broken_audit(*args, **kwargs):
        raise sqlite3.IntegrityError("simulated audit failure")

    monkeypatch.setattr(planning_mod, "audit_append", broken_audit)

    result = kp_db.mark_plates_as_planned(
        kp_id=KP_ID,
        plate_name=PLATE_NAME,
        qty_to_plan=3,
        plan_id="plan_test_audit_rollback",
        db_path=tmp_plita,
        actor="test:atomic",
    )

    monkeypatch.setattr(planning_mod, "audit_append", original_audit)

    # Сама функция вернула success=False (исключение поймано)
    assert result["success"] is False

    # Главное: статус плит остался 'в производстве' — UPDATE откатился
    with sqlite3.connect(tmp_plita) as conn:
        rows = conn.execute(
            "SELECT status, qty FROM kp_plates WHERE kp_id = ?",
            (KP_ID,),
        ).fetchall()
    assert rows == [("в производстве", 3)], (
        f"После сбоя audit-лога плиты не должны быть помечены 'в плане', "
        f"но имеем: {rows}"
    )

    # И в журнале ничего не записалось
    audit = _audit_rows(tmp_plita)
    assert audit == [], f"Audit-лог должен быть пуст, но содержит: {audit}"
