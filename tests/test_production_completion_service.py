"""Интеграционные тесты для :class:`ProductionCompletionService`.

Покрывают регрессию (P0): после завершения дня бракованные плиты должны
оказаться в статусе ``'в производстве'`` (не ``'в плане'``) и снова
попадать в новые планы — иначе мастер планирования рапортует
«Не найдено плит для планирования».

Также проверяют helper :meth:`ProductionCompletionService._return_rejected`
на идемпотентность: повторный вызов с теми же данными не дублирует возвраты.
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
KP_ID = 1


@pytest.fixture
def tmp_plita(tmp_path) -> str:
    """Готовит ``plita.db`` с одним КП и тремя плитами в производстве."""
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
    """Изолируем planning-файлы и отключаем праздники/выходные."""
    plans_dir = tmp_path / "plans"
    plans_dir.mkdir()
    metadata_path = tmp_path / "plans_metadata.json"

    monkeypatch.setattr(plan_manager, "PLANS_DIR", plans_dir)
    monkeypatch.setattr(plan_manager, "PLANS_METADATA_PATH", metadata_path)
    monkeypatch.setattr(plan_manager, "load_holidays", lambda: set())
    monkeypatch.setattr(plan_manager, "load_extra_workdays", lambda: set())


@pytest.fixture
def planning_service(tmp_plita, monkeypatch):
    """Собирает :class:`ProductionPlanningService` с мок-оптимизатором."""
    from app.services.production_planning_service import ProductionPlanningService

    service = ProductionPlanningService(
        plita_db_path=tmp_plita,
        pb_db_path=tmp_plita,
    )

    def fake_optimize(self, *, orders_2d):
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


def _make_production_service(planning_service, db_path):
    from app.repositories.kp_repository import KpRepository
    from app.repositories.plan_repository import PlanRepository
    from app.services.production_service import ProductionService

    return ProductionService(
        kp_repository=KpRepository(db_path=db_path),
        plan_repository=PlanRepository(),
        planning_service=planning_service,
    )


def _kp_plate_rows(db_path: str) -> list[tuple[str, int, str | None]]:
    """Возвращает (status, qty, plan_id) по всем строкам kp_plates КП."""
    with sqlite3.connect(db_path) as conn:
        return conn.execute(
            "SELECT status, qty, plan_id FROM kp_plates "
            "WHERE kp_id = ? ORDER BY id",
            (KP_ID,),
        ).fetchall()


def _kp_status(db_path: str) -> str:
    with sqlite3.connect(db_path) as conn:
        return conn.execute(
            "SELECT status FROM kp_meta WHERE kp_id = ?",
            (KP_ID,),
        ).fetchone()[0]


def _completed_total(db_path: str) -> int:
    with sqlite3.connect(db_path) as conn:
        return conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id = ?",
            (KP_ID,),
        ).fetchone()[0]


def _day_completed(plan_id: str, date_key: str) -> bool:
    plan = plan_manager.load_plan(plan_id)
    assert plan is not None
    return bool((plan.get("days") or {}).get(date_key, {}).get("completed"))


def test_full_reject_returns_plate_to_production(planning_service, tmp_plita):
    """Полностью забракованная позиция возвращается в 'в производстве'."""
    built = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    service = _make_production_service(planning_service, tmp_plita)
    result = service.complete_day(
        plan_id=plan_id,
        target_date="2026-04-21",
        rejected_plates=[{"track_number": 1, "plate_index": 0, "qty": 3}],
    )

    assert result["completed"] is True
    assert result["moved_plates"] == 0
    assert result["rejected_returned"] == 3
    assert result["completed_kps"] == []

    rows = _kp_plate_rows(tmp_plita)
    # Брак вернулся в производство, plan_id очищен
    assert rows == [("в производстве", 3, None)]
    assert _completed_total(tmp_plita) == 0
    assert _kp_status(tmp_plita) == "в работе"


def test_unmoved_plates_do_not_mark_day_completed(
    planning_service,
    tmp_plita,
    monkeypatch,
):
    """Если БД не списала запрошенные плиты, день нельзя закрывать."""
    from app.services.production_completion_service import ProductionCompletionError

    built = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    monkeypatch.setattr(kp_db, "move_plates_to_completed", lambda *args, **kwargs: 0)

    service = _make_production_service(planning_service, tmp_plita)
    with pytest.raises(
        ProductionCompletionError,
        match=r"не списано 3 плит.*Не хватает:.*КП 1: ПБ 60-12-8п — 3 шт",
    ):
        service.complete_day(plan_id=plan_id, target_date="2026-04-21")

    assert _day_completed(plan_id, "2026-04-21") is False
    assert _completed_total(tmp_plita) == 0
    rows = _kp_plate_rows(tmp_plita)
    assert rows == [("в плане", 3, plan_id)]


def test_plate_without_kp_id_does_not_mark_day_completed(
    planning_service,
    tmp_plita,
    monkeypatch,
):
    """Позиции без kp_id не должны превращаться в ложное completed=True."""
    from app.services.production_completion_service import ProductionCompletionError

    built = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    def fake_day_view(_target_date: str) -> dict:
        return {
            "date": "2026-04-21",
            "plans": [
                {
                    "plan_id": plan_id,
                    "plan_name": plan_id,
                    "completed": False,
                    "tracks": [
                        {
                            "track_number": 1,
                            "plates_info": [
                                {
                                    "plate_name": PLATE_NAME,
                                    "length_m": 6.0,
                                    "width_mm": 1200,
                                    "qty": 3,
                                    "load_code": 8,
                                    "kp_id": None,
                                }
                            ],
                        }
                    ],
                }
            ],
            "plans_count": 1,
            "total_tracks": 1,
        }

    monkeypatch.setattr(
        "app.services.production_completion_service.build_day_view_detail",
        fake_day_view,
    )

    service = _make_production_service(planning_service, tmp_plita)
    with pytest.raises(ProductionCompletionError, match="нет привязки к КП"):
        service.complete_day(plan_id=plan_id, target_date="2026-04-21")

    assert _day_completed(plan_id, "2026-04-21") is False
    assert _completed_total(tmp_plita) == 0
    rows = _kp_plate_rows(tmp_plita)
    assert rows == [("в плане", 3, plan_id)]


def test_partial_reject_splits_correctly(planning_service, tmp_plita):
    """Частичный брак: часть в completed_plates, остаток в 'в производстве'."""
    built = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    service = _make_production_service(planning_service, tmp_plita)
    result = service.complete_day(
        plan_id=plan_id,
        target_date="2026-04-21",
        rejected_plates=[{"track_number": 1, "plate_index": 0, "qty": 1}],
    )

    assert result["completed"] is True
    assert result["moved_plates"] == 2
    assert result["rejected_returned"] == 1
    assert result["completed_kps"] == []

    # В completed_plates ушли 2 плиты, в kp_plates остался брак: qty=1, 'в производстве'
    assert _completed_total(tmp_plita) == 2

    rows = _kp_plate_rows(tmp_plita)
    in_prod = [r for r in rows if r[0] == "в производстве"]
    assert len(in_prod) == 1
    assert in_prod[0][1] == 1  # qty
    assert in_prod[0][2] is None  # plan_id очищен

    # Никаких залипших 'в плане' остаться не должно — это и есть весь смысл фикса
    in_plan = [r for r in rows if r[0] == "в плане"]
    assert sum(r[1] for r in in_plan) == 0

    assert _kp_status(tmp_plita) == "в работе"


def test_rejected_plates_visible_in_next_planning(planning_service, tmp_plita):
    """Регрессия: после полного брака мастер планирования снова видит плиту.

    Это сердце исходного бага: до фикса плиты залипали в ``'в плане'`` и
    повторный ``build_plan`` падал с ``«Не найдено плит для планирования»``.
    """
    from app.services.production_planning_service import ProductionPlanBuildError

    built = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    service = _make_production_service(planning_service, tmp_plita)
    service.complete_day(
        plan_id=plan_id,
        target_date="2026-04-21",
        rejected_plates=[{"track_number": 1, "plate_index": 0, "qty": 3}],
    )

    # Повторное планирование на следующую дату должно построиться,
    # так как все 3 плиты вернулись в 'в производстве'.
    try:
        rebuilt = planning_service.build_plan(
            start_date="2026-04-22",
            tracks_count=3,
            filter_method="all",
        )
    except ProductionPlanBuildError as exc:  # pragma: no cover — диагностический хвост
        pytest.fail(
            f"После возврата брака повторное планирование не должно падать, "
            f"но получили: {exc}"
        )

    assert rebuilt["plan"]["id"]
    assert rebuilt["summary"]["selected_plates_count"] == 3


def test_kp_marked_done_only_when_no_remaining_plates(planning_service, tmp_plita):
    """С браком КП остаётся 'в работе'; без брака — становится 'выполнено'."""
    # 1) С полным браком → 'в работе'
    built = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    service = _make_production_service(planning_service, tmp_plita)
    service.complete_day(
        plan_id=plan_id,
        target_date="2026-04-21",
        rejected_plates=[{"track_number": 1, "plate_index": 0, "qty": 3}],
    )
    assert _kp_status(tmp_plita) == "в работе"

    # 2) Перепланируем + завершим без брака → 'выполнено'
    built2 = planning_service.build_plan(
        start_date="2026-04-22",
        tracks_count=3,
        filter_method="all",
    )
    plan_id2 = built2["plan"]["id"]

    service.complete_day(plan_id=plan_id2, target_date="2026-04-22")
    assert _kp_status(tmp_plita) == "выполнено"


def test_return_rejected_helper_idempotent(planning_service, tmp_plita):
    """Повторный вызов helper'а возврата брака не дублирует возвраты в БД."""
    from app.services.production_completion_service import ProductionCompletionService

    built = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]
    assert plan_id

    rejected_payload = [{"kp_id": KP_ID, "plate_name": PLATE_NAME, "qty": 3}]

    first = ProductionCompletionService._return_rejected(rejected_payload, tmp_plita)
    second = ProductionCompletionService._return_rejected(rejected_payload, tmp_plita)

    assert first == 3
    # После первого вызова в БД нет строк со status='в плане' для этой плиты,
    # поэтому второй вызов ничего не возвращает — никаких дублей.
    assert second == 0

    rows = _kp_plate_rows(tmp_plita)
    assert sum(r[1] for r in rows if r[0] == "в производстве") == 3
    assert sum(r[1] for r in rows if r[0] == "в плане") == 0
