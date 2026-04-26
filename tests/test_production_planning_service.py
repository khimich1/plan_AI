"""Интеграционный тест для :class:`ProductionPlanningService.build_plan`.

Проверяет критичный инвариант: если первый вызов ``build_plan`` успешно
построил план и пометил плиты в БД, второй идентичный вызов должен
выбрасывать :class:`ProductionPlanBuildError` — плит в статусе
``'в производстве'`` больше нет, значит новый план создавать не из чего.

Оптимизатор и фактическое построение дорожек подменяются моком, чтобы тест
не зависел от heavy-CP оптимизации. Папка с JSON-файлами планов и
``plans_metadata.json`` также изолируются во временной директории.
"""
from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from bot.handlers import plan_manager
from core import kp_db


PLATE_NAME = "ПБ 60-12-8п"


@pytest.fixture
def tmp_plita(tmp_path) -> str:
    """Готовит ``plita.db`` с одним КП и тремя плитами в производстве."""
    db_path = str(tmp_path / "plita.db")
    kp_db.init_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date, execution_terms, customer_name) "
            "VALUES (1, '2026-01-01', '21.04.2026', 'ТестКлиент')"
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status) VALUES (1, 'в работе')"
        )
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (1, 1, ?, 6.0, 1.2, 800, 3, 'в производстве')
            """,
            (PLATE_NAME,),
        )
        conn.commit()
    return db_path


@pytest.fixture(autouse=True)
def _isolate_plans(tmp_path, monkeypatch):
    """Изолируем planning-файлы в ``tmp_path`` и отключаем праздники."""
    plans_dir = tmp_path / "plans"
    plans_dir.mkdir()
    metadata_path = tmp_path / "plans_metadata.json"

    monkeypatch.setattr(plan_manager, "PLANS_DIR", plans_dir)
    monkeypatch.setattr(plan_manager, "PLANS_METADATA_PATH", metadata_path)
    monkeypatch.setattr(plan_manager, "load_holidays", lambda: set())
    monkeypatch.setattr(plan_manager, "load_extra_workdays", lambda: set())


@pytest.fixture
def planning_service(tmp_plita, monkeypatch):
    """Собирает :class:`ProductionPlanningService` с lean-оптимизатором."""
    from app.services.production_planning_service import ProductionPlanningService

    service = ProductionPlanningService(
        plita_db_path=tmp_plita,
        pb_db_path=tmp_plita,  # reinforcement через series — не используется в тесте
    )

    def fake_optimize(self, *, orders_2d):
        """Симулирует успешную оптимизацию: 3 плиты → 1 дорожка."""
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
            for _ in range(3)
        ]
        tracks = [{"label": "ОСНОВНАЯ", "items": items}]
        optimization_result = {
            "total_plates": 3,
            "plate_assignments": [
                {
                    "source": "primary",
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"],
                    "load_code": order["load_code"],
                }
                for _ in range(3)
            ],
        }
        return tracks, optimization_result

    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        fake_optimize,
    )
    # get_reinforcement возвращает float, чтобы _load_plates_for_kps не падал
    monkeypatch.setattr(
        "app.services.production_planning_service.get_reinforcement",
        lambda **kwargs: 999.0,
    )
    return service


def test_optimization_service_preserves_identity_orders(monkeypatch):
    from app.domain.models.plate_order import PlateOrder
    from app.services.optimization_service import OptimizationService
    from core.plan_commit import count_assigned_plates

    full_orders = [
        {
            "length": 6.0,
            "width": 1200,
            "qty": 2,
            "load_code": 8,
            "reinforcement": 999.0,
            "kp_date": "21.04.2026",
            "customer": "ТестКлиент",
            "plate_name": PLATE_NAME,
            "kp_id": 1,
            "length_dm_raw": "60",
        }
    ]
    captured: dict[str, list[dict]] = {}

    def fake_optimize(*, orders_2d):
        captured["orders_2d"] = orders_2d
        order = orders_2d[0]
        return {
            "total_plates": order["qty"],
            "plate_assignments": [
                {
                    "source": "primary",
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"],
                    "load_code": order["load_code"],
                }
                for _ in range(order["qty"])
            ],
        }

    monkeypatch.setattr(
        "app.services.optimization_service.optimize_with_cascading_longitudinal_cuts",
        fake_optimize,
    )

    plate_order = PlateOrder.from_orders_2d(full_orders)
    context = OptimizationService().optimize(plate_order, orders_2d=full_orders)

    assert captured["orders_2d"] == full_orders
    counts, unmapped = count_assigned_plates(context.optimization_result, [])
    assert counts["primary"] == {(1, PLATE_NAME): 2}
    assert unmapped["primary"] == []
    assert unmapped["secondary"] == []


def test_build_plan_marks_plates_and_second_call_fails(planning_service, tmp_plita):
    from app.services.production_planning_service import ProductionPlanBuildError

    first = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    assert first["plan"]["id"]

    # После успешного билда плиты должны быть 'в плане'
    with sqlite3.connect(tmp_plita) as conn:
        rows = conn.execute(
            "SELECT status, qty FROM kp_plates WHERE kp_id = 1 AND plate_name = ?",
            (PLATE_NAME,),
        ).fetchall()
    assert rows == [("в плане", 3)]

    # Повторный запрос не должен создать второй план
    with pytest.raises(ProductionPlanBuildError):
        planning_service.build_plan(
            start_date="2026-04-21",
            tracks_count=3,
            filter_method="all",
        )


def test_complete_day_moves_plates_and_marks_kp_completed(planning_service, tmp_plita):
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
        plan_repository=PlanRepository(),
        planning_service=planning_service,
    )
    result = service.complete_day(plan_id=plan_id, target_date="2026-04-21")

    assert result["completed"] is True
    assert result["moved_plates"] == 3
    assert result["completed_kps"] == [1]

    completion = kp_db.get_kp_completion_percentage(1, tmp_plita)
    assert completion["percentage"] == 100.0
    assert completion["completed_plates"] == 3
    assert completion["in_production"] == 0

    with sqlite3.connect(tmp_plita) as conn:
        status = conn.execute(
            "SELECT status FROM kp_meta WHERE kp_id = 1",
        ).fetchone()[0]
        remaining = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = 1",
        ).fetchone()[0]
        completed = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id = 1",
        ).fetchone()[0]

    assert status == "выполнено"
    assert remaining == 0
    assert completed == 3


def test_complete_day_with_partial_rejection_moves_only_accepted_qty(
    planning_service,
    tmp_plita,
):
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
        plan_repository=PlanRepository(),
        planning_service=planning_service,
    )
    result = service.complete_day(
        plan_id=plan_id,
        target_date="2026-04-21",
        rejected_plates=[
            {"track_number": 1, "plate_index": 0, "qty": 1},
        ],
    )

    assert result["completed"] is True
    assert result["moved_plates"] == 2
    assert result["completed_kps"] == []
    assert result["rejected_plates"] == 1
    assert result["rejected_positions"] == 1

    with sqlite3.connect(tmp_plita) as conn:
        status = conn.execute(
            "SELECT status FROM kp_meta WHERE kp_id = 1",
        ).fetchone()[0]
        remaining = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = 1",
        ).fetchone()[0]
        completed = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id = 1",
        ).fetchone()[0]

    assert status == "в работе"
    assert remaining == 1
    assert completed == 2


def test_complete_day_with_full_rejection_does_not_move_plate(
    planning_service,
    tmp_plita,
):
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
        plan_repository=PlanRepository(),
        planning_service=planning_service,
    )
    result = service.complete_day(
        plan_id=plan_id,
        target_date="2026-04-21",
        rejected_plates=[
            {"track_number": 1, "plate_index": 0, "qty": 3},
        ],
    )

    assert result["completed"] is True
    assert result["moved_plates"] == 0
    assert result["completed_kps"] == []
    assert result["rejected_plates"] == 3
    assert result["rejected_positions"] == 1

    with sqlite3.connect(tmp_plita) as conn:
        status = conn.execute(
            "SELECT status FROM kp_meta WHERE kp_id = 1",
        ).fetchone()[0]
        remaining = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = 1",
        ).fetchone()[0]
        completed = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id = 1",
        ).fetchone()[0]

    assert status == "в работе"
    assert remaining == 3
    assert completed == 0


def test_complete_day_rejects_qty_greater_than_plate_qty(
    planning_service,
    tmp_plita,
):
    from app.repositories.kp_repository import KpRepository
    from app.repositories.plan_repository import PlanRepository
    from app.services.production_completion_service import ProductionCompletionError
    from app.services.production_service import ProductionService

    built = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    service = ProductionService(
        kp_repository=KpRepository(db_path=tmp_plita),
        plan_repository=PlanRepository(),
        planning_service=planning_service,
    )

    with pytest.raises(ProductionCompletionError):
        service.complete_day(
            plan_id=plan_id,
            target_date="2026-04-21",
            rejected_plates=[
                {"track_number": 1, "plate_index": 0, "qty": 4},
            ],
        )
