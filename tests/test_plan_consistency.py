"""Инварианты целостности плана (P7).

Шесть ключевых инвариантов после фикса «не списываются N плит из M»:

1. ``test_kp_plates_day_view_invariant`` — для каждой пары (plan_id, day_number)
   суммарный qty в kp_plates == qty в plates_info.
2. ``test_complete_day_idempotent_on_failure`` — pre-flight отбракованный
   запрос НЕ изменяет БД.
3. ``test_complete_day_atomic_rollback`` — ошибка в середине списания
   не оставляет частичных записей в completed_plates.
4. ``test_rescue_tracks_deterministic`` — ``build_rescue_tracks`` детерминирован.
5. ``test_canonical_plate_name`` — «Плиты ПБ ...» и «ПБ ...» — одна плита.
6. ``test_secondary_no_identity_theft`` — secondary 4.5×600 НЕ забирает
   identity primary 4.5×1200.

Тесты используют общий tmp_path-fixture и мокированный оптимизатор, чтобы
не зависеть от тяжёлой CP-логики.
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
        "core.production.planning.get_reinforcement",
        lambda **kwargs: 999.0,
    )
    return service


def _make_production_service(planning_service, db_path):
    from app.repositories.kp_repository import KpRepository
    from app.repositories.plan_repository import PlanRepository
    from app.services.production_service import ProductionService

    return ProductionService(
        kp_repository=KpRepository(db_path=db_path),
        plan_repository=PlanRepository(db_path=db_path),
        planning_service=planning_service,
    )


def test_kp_plates_day_view_invariant(planning_service, tmp_plita):
    """Инвариант 1: SUM(qty) в kp_plates == SUM(qty) в plates_info для (plan_id, day)."""
    from app.services.day_view_service import build_day_view_detail

    built = planning_service.build_plan(
        start_date="2026-04-21", tracks_count=3, filter_method="all"
    )
    plan_id = built["plan"]["id"]

    with sqlite3.connect(tmp_plita) as conn:
        kp_total = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM kp_plates "
            "WHERE plan_id = ? AND day_number = 1 AND status = 'в плане'",
            (plan_id,),
        ).fetchone()[0]

    day_view = build_day_view_detail("2026-04-21", db_path=tmp_plita)
    assert day_view is not None

    plates_info_total = sum(
        int(p.get("qty") or 0)
        for block in day_view["plans"]
        if block["plan_id"] == plan_id
        for track in block["tracks"]
        for p in track["plates_info"]
    )
    assert kp_total == plates_info_total
    assert kp_total > 0


def test_complete_day_idempotent_on_failure(planning_service, tmp_plita):
    """Инвариант 2: pre-flight ошибка не меняет БД."""
    from app.services.production_completion_service import ProductionCompletionError

    built = planning_service.build_plan(
        start_date="2026-04-21", tracks_count=3, filter_method="all"
    )
    plan_id = built["plan"]["id"]

    # Удаляем kp_plates руками: после этого pre-flight должен поднять ошибку.
    with sqlite3.connect(tmp_plita) as conn:
        conn.execute("DELETE FROM kp_plates WHERE plan_id = ?", (plan_id,))
        conn.commit()

    snap_kp = _snapshot(tmp_plita, "kp_plates")
    snap_completed = _snapshot(tmp_plita, "completed_plates")

    service = _make_production_service(planning_service, tmp_plita)
    with pytest.raises(ProductionCompletionError):
        service.complete_day(plan_id=plan_id, target_date="2026-04-21")

    assert _snapshot(tmp_plita, "kp_plates") == snap_kp
    assert _snapshot(tmp_plita, "completed_plates") == snap_completed


def test_complete_day_atomic_rollback(planning_service, tmp_plita, monkeypatch):
    """Инвариант 3: ошибка в середине цикла → completed_plates не растёт."""
    from app.services.production_completion_service import ProductionCompletionError

    built = planning_service.build_plan(
        start_date="2026-04-21", tracks_count=3, filter_method="all"
    )
    plan_id = built["plan"]["id"]

    snap_kp_before = _snapshot(tmp_plita, "kp_plates")
    snap_completed_before = _snapshot(tmp_plita, "completed_plates")

    # Симулируем падение в середине списания.
    from app.services.plate_completion_service import PlateCompletionService

    def boom(*args, **kwargs):
        raise RuntimeError("boom")

    monkeypatch.setattr(PlateCompletionService, "move_plates_to_completed", boom)

    service = _make_production_service(planning_service, tmp_plita)
    with pytest.raises((ProductionCompletionError, RuntimeError)):
        service.complete_day(plan_id=plan_id, target_date="2026-04-21")

    # Атомарность: completed_plates не должен пополниться, kp_plates не должен схлопнуться.
    assert _snapshot(tmp_plita, "completed_plates") == snap_completed_before
    assert _snapshot(tmp_plita, "kp_plates") == snap_kp_before


def test_rescue_tracks_deterministic():
    """Инвариант 4: core/rescue_tracks даёт детерминированные missing_counts (web path).

    Phase 4: подсчёт идёт по identity (kp_id, plate_name) из plate_assignments.
    """
    from core.rescue_tracks import build_rescue_tracks

    orders_2d = [
        {
            "kp_id": 1,
            "plate_name": "ПБ 60-12-8п",
            "length": 6.0,
            "width": 1200,
            "load_code": 8,
            "qty": 3,
        },
    ]
    plate_assignments = [
        {
            "source": "primary",
            "kp_id": 1,
            "plate_name": "ПБ 60-12-8п",
            "length": 6.0,
            "width": 1200,
            "load_code": 8,
        },
    ]

    rescue_tracks_a, missing_a, rescue_pa_a = build_rescue_tracks(
        orders_2d, list(plate_assignments)
    )
    rescue_tracks_b, missing_b, rescue_pa_b = build_rescue_tracks(
        orders_2d, list(plate_assignments)
    )
    assert missing_a == missing_b
    # Не хватает 2 плит → 2 rescue items и 2 rescue assignments
    total_rescue = sum(len(t.get("items") or []) for t in rescue_tracks_a)
    assert total_rescue == 2
    assert len(rescue_pa_a) == 2
    assert len(rescue_pa_a) == len(rescue_pa_b)


def test_canonical_plate_name():
    """Инвариант 5: «Плиты ПБ ...» и «ПБ ...» — одна плита по canonical()."""
    from core import plate_name as pn

    assert pn.canonical("Плиты ПБ 45-12-6п") == "ПБ 45-12-6п"
    assert pn.canonical("ПБ 45-12-6п") == "ПБ 45-12-6п"
    assert pn.canonical("Плиты  ПБ  45-12-6п ") == "ПБ 45-12-6п"
    assert pn.equal("Плиты ПБ 45-12-6п", "ПБ 45-12-6п")
    assert pn.display("ПБ 45-12-6п") == "Плиты ПБ 45-12-6п"


def test_secondary_no_identity_theft():
    """Инвариант 6: secondary 4.5×600 не получает identity primary 4.5×1200."""
    from app.services.day_view_service import build_smart_lookup

    plate_lookup_exact = {
        (4.5, 1200): [
            {
                "kp_id": 42,
                "kp_date": "2026-01-01",
                "customer": "Тест",
                "plate_name": "ПБ 45-12-6п",
                "reinforcement": 0,
                "load_code": 6,
                "qty_remaining": 1,
            }
        ],
    }
    plate_lookup_by_length = {
        4.5: list(plate_lookup_exact[(4.5, 1200)]),
    }
    lookup = build_smart_lookup(plate_lookup_exact, plate_lookup_by_length)

    # Secondary с шириной 600 не должен забрать identity primary 1200.
    info = lookup(4.5, 600)
    assert info.get("kp_id") is None
    assert info.get("plate_name") == ""

    # Primary 4.5×1200 — должен получить identity.
    info = lookup(4.5, 1200)
    assert info.get("kp_id") == 42


def test_kp_plates_day_view_invariant_with_secondary_cuts(tmp_plita, monkeypatch):
    """P9: инвариант ``kp_plates ↔ day_view`` сохраняется и для треков с
    ``secondary_cuts`` без явного ``target_length``.

    Сценарий из ошибки пользователя: secondary без identity и без
    ``target_length`` — после backfill_track_items_identity он получает
    ``kp_id``+``plate_name``, попадает в commit, листится в day_view
    и потом списывается complete_day.
    """
    secondary_plate_name = "ПБ 60-3,2-8п"

    with sqlite3.connect(tmp_plita) as conn:
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (?, 2, ?, 6.0, 0.32, 800, 2, 'в производстве')
            """,
            (KP_ID, secondary_plate_name),
        )
        conn.commit()

    from app.services.production_planning_service import ProductionPlanningService

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
                        "label": "[2] secondary без target_length",
                        "load_code": primary_order["load_code"],
                    }
                ],
            })
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

        # Mock замещает _run_optimization_and_split целиком, поэтому
        # обязаны сами вызвать backfill_track_items_identity — точно так же,
        # как это делает реальный метод после rescue.
        from core.plate_attribution import backfill_track_items_identity
        backfill_track_items_identity(tracks, orders_2d)

        return tracks, optimization_result

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
    )

    built = service.build_plan(
        start_date="2026-04-21", tracks_count=3, filter_method="all"
    )
    plan_id = built["plan"]["id"]

    with sqlite3.connect(tmp_plita) as conn:
        rows = conn.execute(
            "SELECT plate_name, day_number, qty, status FROM kp_plates "
            "WHERE plan_id = ?",
            (plan_id,),
        ).fetchall()

    assert rows, "ожидается, что план пометил какие-то плиты"
    null_day_rows = [r for r in rows if r[1] is None and r[3] == "в плане"]
    assert null_day_rows == [], (
        f"в плане не должно быть строк с day_number IS NULL: {null_day_rows}"
    )

    plate_names_in_plan = {r[0] for r in rows if r[3] == "в плане"}
    assert PLATE_NAME in plate_names_in_plan, (
        f"primary {PLATE_NAME} должен быть в плане"
    )
    assert secondary_plate_name in plate_names_in_plan, (
        f"secondary {secondary_plate_name} должен попасть в план "
        "(после P9 backfill его identity)"
    )

    from app.services.day_view_service import build_day_view_detail

    day_view = build_day_view_detail("2026-04-21", db_path=tmp_plita)
    assert day_view is not None

    plates_info_by_name: dict[str, int] = {}
    for block in day_view["plans"]:
        if block["plan_id"] != plan_id:
            continue
        for track in block["tracks"]:
            for p in track.get("plates_info") or []:
                name = (p.get("plate_name") or "").strip()
                if not name:
                    continue
                plates_info_by_name[name] = (
                    plates_info_by_name.get(name, 0) + int(p.get("qty") or 0)
                )

    canonical_pairs = {(PLATE_NAME, 3), (secondary_plate_name, 2)}
    actual_canonical_pairs: set[tuple[str, int]] = set()
    for name, qty in plates_info_by_name.items():
        if "ПБ 60-12" in name:
            actual_canonical_pairs.add((PLATE_NAME, qty))
        elif "ПБ 60-3,2" in name:
            actual_canonical_pairs.add((secondary_plate_name, qty))

    assert canonical_pairs == actual_canonical_pairs, (
        f"plates_info не совпадает с ожидаемым {canonical_pairs}: "
        f"actual={plates_info_by_name}"
    )


def _snapshot(db_path: str, table: str) -> list[tuple]:
    """Возвращает упорядоченный snapshot всех строк таблицы для сравнения."""
    with sqlite3.connect(db_path) as conn:
        cur = conn.execute(f"SELECT * FROM {table} ORDER BY 1")
        return cur.fetchall()
