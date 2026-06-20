"""Smoke-сценарий для режима ``fill_targets``: 3 дня в корзине, излишек плит.

Имитирует ручной smoke из плана: пользователь добавил три разных дня
(2 + 3 + 1 = 6 дорожек), а оптимизатор выдал 10 одиночных дорожек. Лишние
4 дорожки должны быть обрезаны, плиты под ними — остаться в статусе
``в производстве``, а 6 распределённых строго по таргетам — стать
``в плане``.
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


@pytest.fixture
def tmp_plita(tmp_path) -> str:
    """Готовит plita.db с одним КП и 10 плитами в производстве."""
    db_path = str(tmp_path / "plita.db")
    kp_db.init_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date, execution_terms, customer_name) "
            "VALUES (1, '2026-01-01', '21.04.2026', 'СмокеКлиент')"
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status) VALUES (1, 'в работе')"
        )
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (1, 1, ?, 6.0, 1.2, 800, 10, 'в производстве')
            """,
            (PLATE_NAME,),
        )
        conn.commit()
    return db_path


@pytest.fixture(autouse=True)
def _isolate_plans(tmp_path, monkeypatch):
    plans_dir = tmp_path / "plans"
    plans_dir.mkdir()
    monkeypatch.setattr(plan_manager, "PLANS_DIR", plans_dir)
    monkeypatch.setattr(
        plan_manager, "PLANS_METADATA_PATH", tmp_path / "plans_metadata.json"
    )
    monkeypatch.setattr(plan_manager, "load_holidays", lambda: set())
    monkeypatch.setattr(plan_manager, "load_extra_workdays", lambda: set())


@pytest.fixture
def planning_service(tmp_plita, monkeypatch):
    """Сервис с фейковой оптимизацией: 10 одиночных дорожек."""
    from app.services.production_planning_service import ProductionPlanningService

    service = ProductionPlanningService(plita_db_path=tmp_plita, pb_db_path=tmp_plita)

    def fake_optimize(self, *, orders_2d, **kwargs):
        if not orders_2d:
            return [], {}
        order = orders_2d[0]
        tracks: list[dict] = []
        assignments: list[dict] = []
        for _ in range(10):
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
        return tracks, {"total_plates": 10, "plate_assignments": assignments}

    monkeypatch.setattr(
        ProductionPlanningService, "_run_optimization_and_split", fake_optimize
    )
    monkeypatch.setattr(
        "core.production.planning.get_reinforcement",
        lambda **kwargs: 999.0,
    )
    return service


def test_smoke_three_days_with_excess_plates(planning_service, tmp_plita):
    """3 дня (2+3+1=6 дорожек), оптимизатор выдал 10 → обрезаем 4 лишних.

    Контролируем:
    - tracks_by_day строго совпадает с корзиной;
    - 6 плит ушли «в плане», 4 остались «в производстве»;
    - суммарное количество плит сохранилось (10 = 6 + 4).
    """
    fill_targets = [
        {"date": "2026-04-27", "tracks": 2},
        {"date": "2026-04-28", "tracks": 3},
        {"date": "2026-04-29", "tracks": 1},
    ]

    result = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=10,
        filter_method="all",
        fill_targets=fill_targets,
    )

    days = result["plan"]["days"]
    assert set(days.keys()) == {"2026-04-27", "2026-04-28", "2026-04-29"}
    assert len(days["2026-04-27"]["tracks"]) == 2
    assert len(days["2026-04-28"]["tracks"]) == 3
    assert len(days["2026-04-29"]["tracks"]) == 1

    # plate_assignments в финальном плане должен содержать ровно 6 элементов:
    # столько же, сколько помещено в дорожки.
    assignments = result["plan"]["optimization_result"].get("plate_assignments", [])
    assert len([a for a in assignments if a.get("source") == "primary"]) == 6

    # Статусы в БД: 6 → "в плане", 4 → "в производстве".
    # P5: один identity может теперь занимать несколько строк kp_plates
    # (по дням), поэтому сравниваем СУММЫ статусов.
    with sqlite3.connect(tmp_plita) as conn:
        rows = conn.execute(
            "SELECT status, SUM(qty) FROM kp_plates "
            "WHERE kp_id = 1 AND plate_name = ? GROUP BY status",
            (PLATE_NAME,),
        ).fetchall()
    statuses = dict(rows)
    assert statuses.get("в плане") == 6
    assert statuses.get("в производстве") == 4
    assert sum(qty for _, qty in rows) == 10
