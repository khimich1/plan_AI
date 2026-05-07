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

from app.planning import plan_manager
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

    def fake_day_view(_target_date: str, **_kwargs) -> dict:
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


def test_day_view_write_off_completed_false_before_complete_true_after_snapshot(
    planning_service,
    tmp_plita,
):
    """До complete_day плиты из kp_plates без флага списания; после — снимок с write_off_completed."""
    from app.services.day_view_service import build_day_view_detail

    built = planning_service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    before = build_day_view_detail("2026-04-21", db_path=tmp_plita)
    assert before is not None
    before_plates = [
        p
        for block in before["plans"]
        if block["plan_id"] == plan_id
        for tr in block["tracks"]
        for p in (tr.get("plates_info") or [])
    ]
    assert before_plates, "плиты должны отображаться до списания дня"
    assert all(not p.get("write_off_completed") for p in before_plates), (
        f"живые строки kp_plates не помечаются write_off_completed: {before_plates}"
    )

    service = _make_production_service(planning_service, tmp_plita)
    result = service.complete_day(plan_id=plan_id, target_date="2026-04-21")
    assert result["completed"] is True
    assert result["moved_plates"] == 3

    after = build_day_view_detail("2026-04-21", db_path=tmp_plita)
    assert after is not None
    after_plates = [
        p
        for block in after["plans"]
        if block["plan_id"] == plan_id
        for tr in block["tracks"]
        for p in (tr.get("plates_info") or [])
    ]
    assert after_plates, "после complete_day day_view реаттачит плиты из журнала/снимка"
    assert all(p.get("write_off_completed") is True for p in after_plates), (
        f"позиции со снимка должны быть write_off_completed=True: {after_plates}"
    )


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


def test_to_completed_payload_prefers_explicit_load_class():
    from app.services.production_completion_service import ProductionCompletionService

    payload = ProductionCompletionService._to_completed_plate_payload(
        {
            "kp_id": 7,
            "plate_name": "ПБ 55-12-12,5п",
            "length_m": 5.5,
            "width_mm": 1200,
            "qty": 2,
            "load_code": 12,
            "load_class": 1250,
        }
    )

    assert payload["load_class"] == 1250


def test_to_completed_payload_legacy_fallback_uses_load_code():
    from app.services.production_completion_service import ProductionCompletionService

    payload = ProductionCompletionService._to_completed_plate_payload(
        {
            "kp_id": 7,
            "plate_name": "ПБ 60-12-8п",
            "length_m": 6.0,
            "width_mm": 1200,
            "qty": 1,
            "load_code": 8,
        }
    )

    assert payload["load_class"] == 800


def test_verify_plates_exist_preflight_accepts_1250(tmp_path):
    from app.services.production_completion_service import ProductionCompletionService

    db_path = str(tmp_path / "plita_1250.db")
    kp_db.init_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (?, 1, ?, 5.5, 1.2, 1250, 6, 'в плане')
            """,
            (KP_ID, "ПБ 55-12-12,5п"),
        )
        conn.commit()

        missing = ProductionCompletionService._verify_plates_exist_in_db(
            {
                KP_ID: [
                    {
                        "kp_id": KP_ID,
                        "plate_name": "ПБ 55-12-12,5п",
                        "length_m": 5.5,
                        "width_m": 1.2,
                        "load_class": 1250,
                        "qty": 6,
                    }
                ]
            },
            conn,
        )

    assert missing == []


def test_complete_day_succeeds_for_1250_load_class(tmp_path, monkeypatch):
    from app.services.production_completion_service import ProductionCompletionService

    class _PlanRepositoryStub:
        def load_plan(self, _plan_id: str) -> dict:
            return {"id": "plan-1250", "days": {"2026-04-30": {"day_number": 1}}}

    db_path = str(tmp_path / "plita_1250.db")
    kp_db.init_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date, execution_terms, customer_name) "
            "VALUES (?, '2026-04-01', '06.06.2026', 'Ромашка')",
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
                load_class, qty, status, plan_id, day_number
            ) VALUES (?, 1, ?, 5.5, 1.2, 1250, 6, 'в плане', ?, 1)
            """,
            (KP_ID, "ПБ 55-12-12,5п", "plan-1250"),
        )
        conn.commit()

    def _fake_day_view(_target_date: str, **_kwargs) -> dict:
        return {
            "date": "2026-04-30",
            "plans": [
                {
                    "plan_id": "plan-1250",
                    "plan_name": "plan-1250",
                    "completed": False,
                    "tracks": [
                        {
                            "track_number": 1,
                            "plates_info": [
                                {
                                    "kp_id": KP_ID,
                                    "plate_name": "ПБ 55-12-12,5п",
                                    "length_m": 5.5,
                                    "width_mm": 1200,
                                    "qty": 6,
                                    "load_code": 12,
                                    "load_class": 1250,
                                }
                            ],
                        }
                    ],
                }
            ],
        }

    monkeypatch.setattr(
        "app.services.production_completion_service.build_day_view_detail",
        _fake_day_view,
    )

    service = ProductionCompletionService(
        db_path=db_path,
        plan_repository=_PlanRepositoryStub(),
    )
    result = service.complete_day(
        plan_id="plan-1250",
        target_date="2026-04-30",
    )

    assert result["moved_plates"] == 6
    with sqlite3.connect(db_path) as conn:
        in_plan_qty = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE status='в плане' AND kp_id=?",
            (KP_ID,),
        ).fetchone()[0]
        completed_qty = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id=?",
            (KP_ID,),
        ).fetchone()[0]

    assert in_plan_qty == 0
    assert completed_qty == 6


def test_complete_day_handles_secondary_with_backfilled_identity(
    planning_service,
    tmp_plita,
    monkeypatch,
):
    """P9: secondary с backfilled identity листится в plates_info day_view
    и списывается complete_day, попадая в completed_plates.

    Сценарий: в БД есть КП с primary 6x1.2 и secondary 6x0.32 без явного
    target_length у secondary cut. После build_plan + complete_day обе
    позиции переходят в completed_plates.
    """
    from app.services.production_planning_service import ProductionPlanningService

    secondary_plate_name = "ПБ 60-3,2-8п"

    with sqlite3.connect(tmp_plita) as conn:
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (?, 2, ?, 6.0, 0.32, 800, 3, 'в производстве')
            """,
            (KP_ID, secondary_plate_name),
        )
        conn.commit()

    def fake_optimize_with_secondary(self, *, orders_2d):
        if not orders_2d:
            return [], {}
        primary_order = next(
            o for o in orders_2d if int(round(float(o["width"]))) == 1200
        )
        secondary_order = next(
            o for o in orders_2d if int(round(float(o["width"]))) == 320
        )

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

    service = ProductionPlanningService(
        plita_db_path=tmp_plita,
        pb_db_path=tmp_plita,
    )
    monkeypatch.setattr(
        "app.services.production_planning_service.get_reinforcement",
        lambda **kwargs: 999.0,
    )

    built = service.build_plan(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    plan_id = built["plan"]["id"]

    from app.services.day_view_service import build_day_view_detail
    day_view = build_day_view_detail("2026-04-21", db_path=tmp_plita)
    assert day_view is not None

    plate_names_in_day_view: set[str] = set()
    for block in day_view["plans"]:
        if block["plan_id"] != plan_id:
            continue
        for track in block["tracks"]:
            for p in track.get("plates_info") or []:
                name = (p.get("plate_name") or "").strip()
                if name:
                    plate_names_in_day_view.add(name)

    assert any("ПБ 60-3,2" in n for n in plate_names_in_day_view), (
        f"secondary должна быть в day_view plates_info, "
        f"но видим: {plate_names_in_day_view}"
    )

    production = _make_production_service(service, tmp_plita)
    result = production.complete_day(plan_id=plan_id, target_date="2026-04-21")

    assert result["completed"] is True
    assert result["moved_plates"] == 6, (
        f"ожидаем списание 3 primary + 3 secondary = 6, получено {result}"
    )

    with sqlite3.connect(tmp_plita) as conn:
        completed_by_name = dict(
            conn.execute(
                "SELECT plate_name, COALESCE(SUM(qty), 0) "
                "FROM completed_plates WHERE kp_id = ? GROUP BY plate_name",
                (KP_ID,),
            ).fetchall()
        )
        remaining = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = ?",
            (KP_ID,),
        ).fetchone()[0]

    assert completed_by_name.get(PLATE_NAME) == 3
    assert completed_by_name.get(secondary_plate_name) == 3
    assert remaining == 0
