"""Тесты режима ``fill_targets`` в :class:`ProductionPlanningService`.

Покрывают четыре ключевых инварианта:

* split — корзина из нескольких дней раскладывает дорожки строго по таргетам;
* too-many — таргет, превышающий свободные слоты, валится с
  ``ProductionPlanBuildError``;
* duplicate — Pydantic-валидатор схемы запрещает дублирующиеся даты;
* active_plan_id-ignored — режим дозаполнения всегда создаёт новый план,
  игнорируя переданный ``active_plan_id``.

Оптимизация подменяется фейком, чтобы тест не зависел от тяжёлой CP-логики.
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
    """KP с большим qty: 8 плит ``в производстве`` под одно КП."""
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
            ) VALUES (1, 1, ?, 6.0, 1.2, 800, 8, 'в производстве')
            """,
            (PLATE_NAME,),
        )
        conn.commit()
    return db_path


@pytest.fixture(autouse=True)
def _isolate_plans(tmp_path, monkeypatch):
    """Изолируем planning-файлы и убираем праздники, чтобы дата всегда «рабочая»."""
    plans_dir = tmp_path / "plans"
    plans_dir.mkdir()
    metadata_path = tmp_path / "plans_metadata.json"

    monkeypatch.setattr(plan_manager, "PLANS_DIR", plans_dir)
    monkeypatch.setattr(plan_manager, "PLANS_METADATA_PATH", metadata_path)
    monkeypatch.setattr(plan_manager, "load_holidays", lambda: set())
    monkeypatch.setattr(plan_manager, "load_extra_workdays", lambda: set())


@pytest.fixture
def planning_service(tmp_plita, monkeypatch):
    """Сервис с фейк-оптимизатором: ровно 8 одиночных дорожек."""
    from app.services.production_planning_service import ProductionPlanningService

    service = ProductionPlanningService(
        plita_db_path=tmp_plita,
        pb_db_path=tmp_plita,
    )

    def fake_optimize(self, *, orders_2d):
        # 8 дорожек, по 1 плите в каждой — удобно проверять обрезание.
        if not orders_2d:
            return [], {}
        order = orders_2d[0]
        tracks: list[dict] = []
        assignments: list[dict] = []
        for _ in range(8):
            item = {
                "kp_id": order["kp_id"],
                "plate_name": order["plate_name"],
                "length": order["length"],
                "width": order["width"],
                "load_code": order["load_code"],
            }
            tracks.append({"label": "ОСНОВНАЯ", "items": [item]})
            assignments.append(
                {
                    "source": "primary",
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"],
                    "load_code": order["load_code"],
                }
            )
        optimization_result = {
            "total_plates": 8,
            "plate_assignments": assignments,
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


def test_fill_targets_split_across_days(planning_service, tmp_plita):
    """2 дня в корзине: 2+3 = 5 дорожек, лишние 3 плиты остаются 'в производстве'."""
    result = planning_service.build_plan(
        start_date="2026-04-21",  # будет перезаписан min(fill_targets.date)
        tracks_count=10,            # игнорируется в режиме fill_targets
        filter_method="all",
        fill_targets=[
            {"date": "2026-04-27", "tracks": 2},
            {"date": "2026-04-28", "tracks": 3},
        ],
    )

    days = result["plan"]["days"]
    # Ровно две даты, без переноса лишнего на следующий рабочий день.
    assert set(days.keys()) == {"2026-04-27", "2026-04-28"}
    assert len(days["2026-04-27"]["tracks"]) == 2
    assert len(days["2026-04-28"]["tracks"]) == 3

    # 5 плит ушли в план, 3 должны остаться в производстве.
    with sqlite3.connect(tmp_plita) as conn:
        rows = conn.execute(
            "SELECT status, qty FROM kp_plates "
            "WHERE kp_id = 1 AND plate_name = ? ORDER BY status",
            (PLATE_NAME,),
        ).fetchall()
    statuses = {status: qty for status, qty in rows}
    assert statuses.get("в плане") == 5
    assert statuses.get("в производстве") == 3


def test_fill_targets_validation_too_many(planning_service):
    """Запрос 4 дорожек на день, где свободно лишь 2 — ошибка планирования."""
    from app.services.production_planning_service import ProductionPlanBuildError

    # Подсовываем фейковую занятость: на 27.04 уже занято 2 из 3 (max=3).
    # Свободно 1 → запрос 2 уже не пройдёт.
    target_date = "2026-04-27"
    fake_max_per_day = 3

    # Перехватываем плановый MAX_TRACKS_PER_DAY и occupancy через monkeypatch
    # на сервисе (через fixture monkeypatch требуется отдельное обращение).
    import app.services.production_planning_service as svc_mod
    from unittest.mock import patch

    with patch.object(svc_mod.plan_manager, "MAX_TRACKS_PER_DAY", fake_max_per_day), \
         patch.object(
             svc_mod.plan_manager,
             "get_global_day_occupancy",
             return_value={target_date: 2},
         ):
        with pytest.raises(ProductionPlanBuildError) as exc_info:
            planning_service.build_plan(
                start_date="2026-04-27",
                tracks_count=2,
                filter_method="all",
                fill_targets=[{"date": target_date, "tracks": 2}],
            )
    assert "свободно" in str(exc_info.value).lower()


def test_fill_targets_duplicate_dates_rejected():
    """Pydantic-схема ругается на повторяющиеся даты в корзине."""
    from pydantic import ValidationError

    from app.schemas.production import BuildPlanRequest

    with pytest.raises(ValidationError):
        BuildPlanRequest(
            start_date="2026-04-27",
            tracks_count=2,
            filter_method="all",
            fill_targets=[
                {"date": "2026-04-27", "tracks": 2},
                {"date": "2026-04-27", "tracks": 1},
            ],
        )


def test_fill_targets_active_plan_id_ignored(planning_service, tmp_plita):
    """Режим дозаполнения никогда не дописывает в чужой план."""
    # Любой существующий plan_id — фейковый, всё равно должен быть проигнорирован.
    result = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=5,
        filter_method="all",
        active_plan_id="some-other-plan",
        fill_targets=[
            {"date": "2026-04-27", "tracks": 2},
            {"date": "2026-04-28", "tracks": 3},
        ],
    )

    plan = result["plan"]
    # add_tracks_to_plan для несуществующего id создаёт новый план,
    # но новый id обязан отличаться от того, что мы передали, чтобы
    # подтвердить семантику «дозаполнение всегда новый план».
    assert plan["id"] != "some-other-plan"
    assert set(plan["days"].keys()) == {"2026-04-27", "2026-04-28"}
