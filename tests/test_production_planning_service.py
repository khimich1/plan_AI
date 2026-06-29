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

from app.repositories.plan_repository import PlanRepository
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
    """Отключаем праздники при распределении дорожек по дням."""
    import core.work_calendar as work_calendar

    monkeypatch.setattr(work_calendar, "load_holidays", lambda: set())
    monkeypatch.setattr(work_calendar, "load_extra_workdays", lambda: set())


@pytest.fixture
def planning_service(tmp_plita, monkeypatch):
    """Собирает :class:`ProductionPlanningService` с lean-оптимизатором."""
    from app.services.production_planning_service import ProductionPlanningService

    plan_repo = PlanRepository(db_path=tmp_plita)
    service = ProductionPlanningService(
        plita_db_path=tmp_plita,
        pb_db_path=tmp_plita,  # reinforcement через series — не используется в тесте
        plan_repository=plan_repo,
    )

    def fake_optimize(self, *, orders_2d, layout_reinforcement_order="asc", **kwargs):
        """Симулирует успешную оптимизацию: 3 плиты → 1 дорожка."""
        if not orders_2d:
            return [], {}
        order = orders_2d[0]
        order_qty = int(order.get("qty") or 1)
        items = [
            {
                "kp_id": order["kp_id"],
                "plate_name": order["plate_name"],
                "length": order["length"],
                "width": order["width"],
                "load_code": order["load_code"],
            }
            for _ in range(order_qty)
        ]
        tracks = [{"label": "ОСНОВНАЯ", "items": items}]
        optimization_result = {
            "total_plates": order_qty,
            "plate_assignments": [
                {
                    "source": "primary",
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"],
                    "load_code": order["load_code"],
                }
                for _ in range(order_qty)
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
        "core.production.planning.get_reinforcement",
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

    def fake_optimize(*, orders_2d, **kwargs):
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


@pytest.fixture
def tmp_plita_qty7(tmp_path):
    """КП с одной строкой kp_plates qty=7 для теста частичного планирования."""
    db_path = str(tmp_path / "plita_qty7.db")
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
            ) VALUES (1, 1, ?, 6.0, 1.2, 800, 7, 'в производстве')
            """,
            (PLATE_NAME,),
        )
        plate_id = conn.execute(
            "SELECT id FROM kp_plates WHERE kp_id = 1"
        ).fetchone()[0]
        conn.commit()
    return db_path, int(plate_id)


@pytest.fixture
def planning_service_qty7(tmp_plita_qty7, tmp_path, monkeypatch):
    """ProductionPlanningService на БД с qty=7."""
    from app.services.production_planning_service import ProductionPlanningService

    db_path, _plate_id = tmp_plita_qty7
    plan_repo = PlanRepository(db_path=db_path)
    service = ProductionPlanningService(
        plita_db_path=db_path,
        pb_db_path=db_path,
        plan_repository=plan_repo,
    )

    def fake_optimize(self, *, orders_2d, **kwargs):
        if not orders_2d:
            return [], {}
        order = orders_2d[0]
        order_qty = int(order.get("qty") or 1)
        items = [
            {
                "kp_id": order["kp_id"],
                "plate_name": order["plate_name"],
                "length": order["length"],
                "width": order["width"],
                "load_code": order["load_code"],
            }
            for _ in range(order_qty)
        ]
        tracks = [{"label": "ОСНОВНАЯ", "items": items}]
        optimization_result = {
            "total_plates": order_qty,
            "plate_assignments": [
                {
                    "source": "primary",
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"],
                    "load_code": order["load_code"],
                }
                for _ in range(order_qty)
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


def _sum_qty_by_status(db_path: str, status: str) -> int:
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = 1 AND plate_name = ? AND status = ?",
            (PLATE_NAME, status),
        ).fetchone()
    return int(row[0])


def test_build_plan_partial_plate_qty(planning_service_qty7, tmp_plita_qty7):
    """Урезание qty (7→3): в плане 3, остаток 4 остаётся «в производстве»."""
    db_path, plate_id = tmp_plita_qty7

    first = planning_service_qty7.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="kp",
        selected_kp_ids=[1],
        selected_plate_ids={1: [plate_id]},
        selected_plate_qty={1: {plate_id: 3}},
    )
    assert first["plan"]["id"]
    assert first["summary"]["selected_plates_count"] == 3

    assert _sum_qty_by_status(db_path, "в плане") == 3
    assert _sum_qty_by_status(db_path, "в производстве") == 4

    # После split остаток — новая строка kp_plates с другим id (как в UI после reload).
    with sqlite3.connect(db_path) as conn:
        remainder_id = conn.execute(
            "SELECT id FROM kp_plates WHERE kp_id = 1 AND status = 'в производстве'"
        ).fetchone()[0]

    second = planning_service_qty7.build_plan(
        start_date="2026-04-22",
        tracks_count=3,
        filter_method="kp",
        selected_kp_ids=[1],
        selected_plate_ids={1: [remainder_id]},
        selected_plate_qty={1: {remainder_id: 4}},
    )
    assert second["plan"]["id"]
    assert _sum_qty_by_status(db_path, "в плане") == 7
    assert _sum_qty_by_status(db_path, "в производстве") == 0


def test_build_plan_rejects_qty_above_available(planning_service_qty7, tmp_plita_qty7):
    from app.services.production_planning_service import ProductionPlanBuildError

    _, plate_id = tmp_plita_qty7
    with pytest.raises(ProductionPlanBuildError, match="доступно 7"):
        planning_service_qty7.build_plan(
            start_date="2026-04-21",
            tracks_count=3,
            filter_method="kp",
            selected_kp_ids=[1],
            selected_plate_ids={1: [plate_id]},
            selected_plate_qty={1: {plate_id: 10}},
        )


def test_build_plan_marks_plates_and_second_call_fails(planning_service, tmp_plita):
    from app.services.production_planning_service import ProductionPlanBuildError

    first = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    assert first["plan"]["id"]
    assert first["plan"].get("version") == 1

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


def test_build_plan_desc_reinforcement_order_smoke(planning_service, tmp_plita):
    """build_plan с layout_reinforcement_order=desc не падает и сохраняет режим в плане."""
    result = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
        layout_reinforcement_order="desc",
    )
    assert result["plan"]["id"]
    assert result["plan"].get("layout_reinforcement_order") == "desc"


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
        plan_repository=PlanRepository(db_path=tmp_plita),
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
        plan_repository=PlanRepository(db_path=tmp_plita),
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
        plan_repository=PlanRepository(db_path=tmp_plita),
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
        plan_repository=PlanRepository(db_path=tmp_plita),
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


def test_full_cycle_no_stuck_plates(tmp_plita, monkeypatch):
    """P9 e2e: воспроизведение кейса пользователя в миниатюре.

    Реалистичный набор: primary 6x1.2 (3 шт) + secondary 6x0.32 (6 шт),
    разложенный на несколько дорожек/дней. После полного прохождения
    цикла build_plan → complete_day по всем дням ожидаем:
    - ``kp_plates`` для этого плана пуст (всё списано);
    - KP помечен ``'выполнено'``.

    Этот тест подтверждает, что secondary с backfilled identity не
    «зависают» с ``day_number=NULL`` и реально списываются complete_day.
    """
    from app.repositories.kp_repository import KpRepository
    from app.repositories.plan_repository import PlanRepository
    from app.services.production_planning_service import ProductionPlanningService
    from app.services.production_service import ProductionService

    secondary_plate_name = "ПБ 60-3,2-8п"

    with sqlite3.connect(tmp_plita) as conn:
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (1, 2, ?, 6.0, 0.32, 800, 6, 'в производстве')
            """,
            (secondary_plate_name,),
        )
        conn.commit()

    def fake_optimize_with_secondary(self, *, orders_2d, **kwargs):
        if not orders_2d:
            return [], {}
        primary_order = next(o for o in orders_2d if int(round(float(o["width"]))) == 1200)
        secondary_order = next(o for o in orders_2d if int(round(float(o["width"]))) == 320)

        items: list[dict] = []
        for _ in range(int(primary_order["qty"])):
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
                        "label": "[2] secondary",
                        "load_code": primary_order["load_code"],
                    },
                    {
                        "width": 0.32,
                        "label": "[2] secondary",
                        "load_code": primary_order["load_code"],
                    },
                ],
            })

        tracks_per_day = 1
        tracks: list[dict] = []
        for chunk_start in range(0, len(items), tracks_per_day):
            chunk = items[chunk_start:chunk_start + tracks_per_day]
            tracks.append({"label": "ОСНОВНАЯ", "items": chunk})

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

        from core.plate_attribution import backfill_track_items_identity
        backfill_track_items_identity(tracks, orders_2d)

        return tracks, {
            "total_plates": len(plate_assignments),
            "plate_assignments": plate_assignments,
        }

    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        fake_optimize_with_secondary,
    )
    monkeypatch.setattr(
        "core.production.planning.get_reinforcement",
        lambda **kwargs: 999.0,
    )

    service = ProductionPlanningService(
        plita_db_path=tmp_plita,
        pb_db_path=tmp_plita,
        plan_repository=PlanRepository(db_path=tmp_plita),
    )

    built = service.build_plan(
        start_date="2026-04-21",
        tracks_count=1,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    with sqlite3.connect(tmp_plita) as conn:
        rows = conn.execute(
            "SELECT day_number, status FROM kp_plates WHERE plan_id = ?",
            (plan_id,),
        ).fetchall()
    assert rows, "ожидается, что план пометил плиты"
    null_day_in_plan = [r for r in rows if r[1] == "в плане" and r[0] is None]
    assert null_day_in_plan == [], (
        f"после P9 fix не должно быть плит с day_number IS NULL: {null_day_in_plan}"
    )

    days = sorted({r[0] for r in rows if r[0] is not None})
    assert days, "ожидается хотя бы один день с плитами"

    production = ProductionService(
        kp_repository=KpRepository(db_path=tmp_plita),
        plan_repository=PlanRepository(db_path=tmp_plita),
        planning_service=service,
    )

    plan_dict = built["plan"]
    sorted_dates = sorted(plan_dict.get("days", {}).keys())
    assert sorted_dates, "ожидается хотя бы один день в плане"

    for date_key in sorted_dates:
        result = production.complete_day(plan_id=plan_id, target_date=date_key)
        assert result["completed"] is True, f"day {date_key} не завершился: {result}"

    with sqlite3.connect(tmp_plita) as conn:
        remaining_plan = conn.execute(
            "SELECT COUNT(*) FROM kp_plates WHERE plan_id = ?", (plan_id,),
        ).fetchone()[0]
        remaining_total = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = 1",
        ).fetchone()[0]
        completed_total = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id = 1",
        ).fetchone()[0]
        kp_status = conn.execute(
            "SELECT status FROM kp_meta WHERE kp_id = 1",
        ).fetchone()[0]

    assert remaining_plan == 0, (
        f"в kp_plates не должно остаться строк плана {plan_id}, "
        f"но осталось {remaining_plan}"
    )
    assert remaining_total == 0, (
        f"у KP=1 не должно остаться плит, но осталось qty={remaining_total}"
    )
    assert completed_total == 9, (
        f"в completed_plates ожидается 9 плит (3 primary + 6 secondary), "
        f"получено {completed_total}"
    )
    assert kp_status == "выполнено", (
        f"KP=1 должен иметь статус 'выполнено', получено '{kp_status}'"
    )
