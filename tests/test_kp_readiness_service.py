"""Unit tests for KpReadinessService (read-only aggregation)."""

from __future__ import annotations

import sqlite3

import pytest

from app.domain.enums import KpStatus, PlateStatus
from app.planning import plan_storage
from app.repositories.plan_repository import PlanRepository
from app.schemas.archive import KpReadinessStepState
from app.schemas.sgp import SgpProgress
from app.services.kp_readiness_service import KpReadinessService
from tests.helpers import kp_db_fixtures as fx

PLATE = "ПБ 59-12-8"
PLAN_ID = "plan-readiness-sgp"


@pytest.fixture
def db(tmp_path) -> str:
    return fx.make_iso_db(tmp_path)


def _patch_plan_db(monkeypatch, db_path: str) -> PlanRepository:
    repo = PlanRepository(db_path=db_path)
    monkeypatch.setattr(plan_storage, "_repo_override", repo)
    from app.core.settings import get_settings

    monkeypatch.setattr(
        type(get_settings()),
        "plita_db_path",
        property(lambda self: db_path),
        raising=False,
    )
    return repo


def _seed_plan_with_two_days(repo: PlanRepository, *, plan_id: str = PLAN_ID) -> None:
    repo.create(
        {
            "id": plan_id,
            "name": "test",
            "days": {
                "2026-08-10": {"day_number": 1, "tracks": []},
                "2026-08-14": {"day_number": 3, "tracks": []},
            },
            "start_date": "2026-08-10",
        }
    )


def _seed_fully_scheduled_kp(
    db_path: str,
    *,
    kp_id: int = 1,
    on_sgp_qty: int = 6,
    in_plan_day1_qty: int = 2,
    in_plan_day3_qty: int = 2,
    plan_id: str = PLAN_ID,
) -> None:
    fx.seed_kp_offer(db_path, kp_id, status=KpStatus.IN_WORK.value)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_meta SET ordered_qty = ? WHERE kp_id = ?",
            (on_sgp_qty + in_plan_day1_qty + in_plan_day3_qty, kp_id),
        )
        conn.commit()
    fx.seed_plate(
        db_path,
        kp_id=kp_id,
        plate_name=PLATE,
        length_m=5.9,
        width_m=1.2,
        load_class=800,
        qty=in_plan_day1_qty,
        status=PlateStatus.IN_PLAN.value,
        position_number=1,
        plan_id=plan_id,
        day_number=1,
    )
    fx.seed_plate(
        db_path,
        kp_id=kp_id,
        plate_name=PLATE,
        length_m=5.9,
        width_m=1.2,
        load_class=800,
        qty=in_plan_day3_qty,
        status=PlateStatus.IN_PLAN.value,
        position_number=2,
        plan_id=plan_id,
        day_number=3,
    )
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, plan_id
            ) VALUES (?, ?, 5.9, 1.2, 800, ?, '28.07.2026', 1, ?)
            """,
            (kp_id, PLATE, on_sgp_qty, plan_id),
        )
        conn.commit()


def _seed_ideation_fixture(db_path: str, *, kp_id: int = 1) -> None:
    """10 ordered: 4 in_plan, 6 on SGP, 0 remaining (ideation example)."""
    fx.seed_kp_offer(db_path, kp_id, status=KpStatus.IN_WORK.value)
    fx.seed_plate(
        db_path,
        kp_id=kp_id,
        plate_name=PLATE,
        length_m=5.9,
        width_m=1.2,
        load_class=800,
        qty=4,
        status=PlateStatus.IN_PLAN.value,
        position_number=1,
    )
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_meta SET ordered_qty = ? WHERE kp_id = ?",
            (10, kp_id),
        )
        conn.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, plan_id
            ) VALUES (?, ?, 5.9, 1.2, 800, 6, '28.07.2026', 1, 'plan-1')
            """,
            (kp_id, PLATE),
        )
        conn.commit()


class TestListPositions:
    def test_ideation_example_single_row(self, db: str) -> None:
        _seed_ideation_fixture(db)
        svc = KpReadinessService(db_path=db)
        items = svc.list_positions(1, status=KpStatus.IN_WORK.value)

        assert len(items) == 1
        row = items[0]
        assert row.label == PLATE
        assert row.ordered == 10
        assert row.in_plan == 4
        assert row.on_sgp == 6
        assert row.remaining == 0
        assert row.ordered == row.in_plan + row.on_sgp + row.remaining

    def test_qty_conservation_per_row(self, db: str) -> None:
        fx.seed_kp_offer(db, 1, status=KpStatus.IN_WORK.value)
        fx.seed_plate(
            db,
            kp_id=1,
            plate_name=PLATE,
            length_m=5.9,
            width_m=1.2,
            load_class=800,
            qty=3,
            status=PlateStatus.IN_PLAN.value,
            position_number=1,
        )
        fx.seed_plate(
            db,
            kp_id=1,
            plate_name=PLATE,
            length_m=5.9,
            width_m=1.2,
            load_class=800,
            qty=2,
            status=PlateStatus.IN_PRODUCTION.value,
            position_number=2,
        )
        with sqlite3.connect(db) as conn:
            conn.execute(
                """
                INSERT INTO completed_plates (
                    kp_id, plate_name, length_m, width_m, load_class, qty, completed_date
                ) VALUES (1, ?, 5.9, 1.2, 800, 1, '28.07.2026')
                """,
                (PLATE,),
            )
            conn.commit()

        items = KpReadinessService(db_path=db).list_positions(1, status=KpStatus.IN_WORK.value)
        assert len(items) == 1
        row = items[0]
        assert row.in_plan == 3
        assert row.remaining == 2
        assert row.on_sgp == 1
        assert row.ordered == 6

    def test_returns_empty_for_archived_status(self, db: str) -> None:
        _seed_ideation_fixture(db)
        svc = KpReadinessService(db_path=db)
        assert svc.list_positions(1, status=KpStatus.ARCHIVED.value) == []


class TestBuildSummary:
    def test_returns_none_for_archived(self, db: str) -> None:
        _seed_ideation_fixture(db)
        with sqlite3.connect(db) as conn:
            conn.execute(
                "UPDATE kp_meta SET status = ? WHERE kp_id = 1",
                (KpStatus.ARCHIVED.value,),
            )
            conn.commit()
        assert KpReadinessService(db_path=db).build_summary(1, status=KpStatus.ARCHIVED.value) is None

    def test_partial_sgp_summary_text(self, db: str) -> None:
        _seed_ideation_fixture(db)
        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        assert summary.sgp_progress is not None
        assert summary.sgp_progress.n == 6
        assert summary.sgp_progress.m == 10
        assert summary.issuable_qty == 6
        assert summary.in_production_qty == 4
        assert "6 из 10 шт на складе" in summary.summary_text
        assert "Можно выдать 6 шт" in summary.summary_text
        assert "Здравствуйте!" in summary.client_copy_text
        assert "№1" in summary.client_copy_text

    def test_no_data_summary(self, db: str) -> None:
        fx.seed_kp_offer(db, 1, status=KpStatus.IN_WORK.value)
        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        assert summary.summary_text == "Данных о производстве пока нет."
        assert "уточняем статус производства" in summary.client_copy_text

    def test_only_in_production_summary(self, db: str) -> None:
        fx.seed_kp_offer(db, 1, status=KpStatus.IN_WORK.value)
        fx.seed_plate(
            db,
            kp_id=1,
            plate_name=PLATE,
            length_m=5.9,
            width_m=1.2,
            load_class=800,
            qty=5,
            status=PlateStatus.IN_PRODUCTION.value,
        )
        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        assert summary.summary_text == "Заказ в производстве (5 шт). На складе пока нет."
        assert "на складе пока нет" in summary.client_copy_text

    def test_full_sgp_on_sgp_status(self, db: str) -> None:
        _seed_ideation_fixture(db)
        with sqlite3.connect(db) as conn:
            conn.execute(
                "UPDATE kp_meta SET status = ? WHERE kp_id = 1",
                (KpStatus.ON_SGP.value,),
            )
            conn.execute(
                "DELETE FROM kp_plates WHERE kp_id = 1 AND status = ?",
                (PlateStatus.IN_PLAN.value,),
            )
            conn.execute(
                """
                INSERT INTO completed_plates (
                    kp_id, plate_name, length_m, width_m, load_class, qty, completed_date
                ) VALUES (1, ?, 5.9, 1.2, 800, 4, '28.07.2026')
                """,
                (PLATE,),
            )
            conn.commit()

        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.ON_SGP.value)
        assert summary is not None
        assert summary.sgp_progress is not None
        assert summary.sgp_progress.n == 10

    def test_stepper_states_partial_sgp(self, db: str) -> None:
        _seed_ideation_fixture(db)
        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        steps = {s.id: s for s in summary.steps}
        assert steps["kp"].state == KpReadinessStepState.DONE
        assert steps["production"].state == KpReadinessStepState.ACTIVE
        assert steps["sgp"].state == KpReadinessStepState.ACTIVE
        assert steps["release"].state == KpReadinessStepState.DISABLED
        assert steps["closed"].state == KpReadinessStepState.DISABLED
        assert steps["sgp"].hint == "6/10"
        assert summary.release_note == "Выдача с СГП — в следующем обновлении"

    def test_stepper_production_done_when_only_sgp(self, db: str) -> None:
        fx.seed_kp_offer(db, 1, status=KpStatus.IN_WORK.value)
        with sqlite3.connect(db) as conn:
            conn.execute(
                "UPDATE kp_meta SET ordered_qty = 5 WHERE kp_id = 1",
            )
            conn.execute(
                """
                INSERT INTO completed_plates (
                    kp_id, plate_name, length_m, width_m, load_class, qty, completed_date
                ) VALUES (1, ?, 5.9, 1.2, 800, 5, '28.07.2026')
                """,
                (PLATE,),
            )
            conn.commit()

        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        steps = {s.id: s for s in summary.steps}
        assert steps["production"].state == KpReadinessStepState.DONE
        assert steps["sgp"].state == KpReadinessStepState.DONE

    def test_summary_when_all_on_sgp_no_pipeline(self, db: str) -> None:
        fx.seed_kp_offer(db, 1, status=KpStatus.IN_WORK.value)
        with sqlite3.connect(db) as conn:
            conn.execute("UPDATE kp_meta SET ordered_qty = 5 WHERE kp_id = 1")
            conn.execute(
                """
                INSERT INTO completed_plates (
                    kp_id, plate_name, length_m, width_m, load_class, qty, completed_date
                ) VALUES (1, ?, 5.9, 1.2, 800, 5, '28.07.2026')
                """,
                (PLATE,),
            )
            conn.commit()

        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        assert summary.in_production_qty == 0
        assert summary.sgp_progress is not None
        assert summary.sgp_progress.n == 5
        assert "на складе. Можно выдать" in summary.summary_text
        assert "на складе, можно забрать" in summary.client_copy_text

    def test_returns_none_for_completed(self, db: str) -> None:
        _seed_ideation_fixture(db)
        assert (
            KpReadinessService(db_path=db).build_summary(1, status=KpStatus.DONE.value) is None
        )


class TestExpectedSgpDate:
    def test_expected_sgp_two_days_same_plan_returns_max_date(self, db: str, monkeypatch) -> None:
        repo = _patch_plan_db(monkeypatch, db)
        _seed_plan_with_two_days(repo)
        _seed_fully_scheduled_kp(db)

        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        assert summary.expected_sgp_date == "2026-08-14"
        assert summary.expected_sgp_date_label == "14.08.2026"
        assert summary.fully_scheduled is True

    def test_expected_sgp_remaining_gt_zero_returns_null(self, db: str, monkeypatch) -> None:
        repo = _patch_plan_db(monkeypatch, db)
        _seed_plan_with_two_days(repo)
        _seed_fully_scheduled_kp(db)
        fx.seed_plate(
            db,
            kp_id=1,
            plate_name=PLATE,
            length_m=5.9,
            width_m=1.2,
            load_class=800,
            qty=2,
            status=PlateStatus.IN_PRODUCTION.value,
            position_number=3,
        )

        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        assert summary.expected_sgp_date is None
        assert summary.expected_sgp_date_label is None
        assert summary.fully_scheduled is False

    def test_expected_sgp_n_equals_m_returns_null(self, db: str, monkeypatch) -> None:
        repo = _patch_plan_db(monkeypatch, db)
        _seed_plan_with_two_days(repo)
        fx.seed_kp_offer(db, 1, status=KpStatus.IN_WORK.value)
        with sqlite3.connect(db) as conn:
            conn.execute("UPDATE kp_meta SET ordered_qty = 5 WHERE kp_id = 1")
            conn.execute(
                """
                INSERT INTO completed_plates (
                    kp_id, plate_name, length_m, width_m, load_class, qty, completed_date
                ) VALUES (1, ?, 5.9, 1.2, 800, 5, '28.07.2026')
                """,
                (PLATE,),
            )
            conn.commit()

        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        assert summary.sgp_progress is not None
        assert summary.sgp_progress.n == summary.sgp_progress.m
        assert summary.expected_sgp_date is None
        assert summary.expected_sgp_date_label is None
        assert summary.fully_scheduled is False

    def test_expected_sgp_client_copy_includes_date(self, db: str, monkeypatch) -> None:
        repo = _patch_plan_db(monkeypatch, db)
        _seed_plan_with_two_days(repo)
        _seed_fully_scheduled_kp(db)

        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        assert "Ожидаем полный комплект на складе к 14.08.2026." in summary.client_copy_text
        assert "6 из 10 шт на складе" in summary.summary_text

    def test_expected_sgp_resolver_skips_missing_plan(self, db: str, monkeypatch) -> None:
        _patch_plan_db(monkeypatch, db)
        _seed_fully_scheduled_kp(db, plan_id="missing-plan")

        svc = KpReadinessService(db_path=db)
        conn = sqlite3.connect(db)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            iso, label, fully = svc._resolve_expected_sgp_date(
                cur,
                1,
                remaining_total=0,
                progress=SgpProgress(n=6, m=10),
                in_plan_total=4,
                plan_mapping_cache={},
            )
        finally:
            conn.close()

        assert iso is None
        assert label is None
        assert fully is False

    def test_expected_sgp_max_across_two_plans(self, db: str, monkeypatch) -> None:
        repo = _patch_plan_db(monkeypatch, db)
        repo.create(
            {
                "id": "plan-a",
                "name": "plan-a",
                "days": {
                    "2026-08-10": {"day_number": 1, "tracks": []},
                    "2026-08-18": {"day_number": 5, "tracks": []},
                },
                "start_date": "2026-08-10",
            }
        )
        repo.create(
            {
                "id": "plan-b",
                "name": "plan-b",
                "days": {"2026-09-01": {"day_number": 1, "tracks": []}},
                "start_date": "2026-09-01",
            }
        )
        fx.seed_kp_offer(db, 1, status=KpStatus.IN_WORK.value)
        with sqlite3.connect(db) as conn:
            conn.execute("UPDATE kp_meta SET ordered_qty = 10 WHERE kp_id = 1")
            conn.commit()
        fx.seed_plate(
            db,
            kp_id=1,
            plate_name=PLATE,
            length_m=5.9,
            width_m=1.2,
            load_class=800,
            qty=3,
            status=PlateStatus.IN_PLAN.value,
            position_number=1,
            plan_id="plan-a",
            day_number=5,
        )
        fx.seed_plate(
            db,
            kp_id=1,
            plate_name=PLATE,
            length_m=5.9,
            width_m=1.2,
            load_class=800,
            qty=2,
            status=PlateStatus.IN_PLAN.value,
            position_number=2,
            plan_id="plan-b",
            day_number=1,
        )
        with sqlite3.connect(db) as conn:
            conn.execute(
                """
                INSERT INTO completed_plates (
                    kp_id, plate_name, length_m, width_m, load_class,
                    qty, completed_date, production_day, plan_id
                ) VALUES (1, ?, 5.9, 1.2, 800, 5, '28.07.2026', 1, 'plan-a')
                """,
                (PLATE,),
            )
            conn.commit()

        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        assert summary.expected_sgp_date == "2026-09-01"
        assert summary.expected_sgp_date_label == "01.09.2026"

    def test_expected_sgp_null_when_in_plan_missing_day_number(self, db: str, monkeypatch) -> None:
        repo = _patch_plan_db(monkeypatch, db)
        _seed_plan_with_two_days(repo)
        fx.seed_kp_offer(db, 1, status=KpStatus.IN_WORK.value)
        with sqlite3.connect(db) as conn:
            conn.execute("UPDATE kp_meta SET ordered_qty = 10 WHERE kp_id = 1")
            conn.commit()
        fx.seed_plate(
            db,
            kp_id=1,
            plate_name=PLATE,
            length_m=5.9,
            width_m=1.2,
            load_class=800,
            qty=4,
            status=PlateStatus.IN_PLAN.value,
            position_number=1,
            plan_id=PLAN_ID,
            day_number=None,
        )
        with sqlite3.connect(db) as conn:
            conn.execute(
                """
                INSERT INTO completed_plates (
                    kp_id, plate_name, length_m, width_m, load_class,
                    qty, completed_date, production_day, plan_id
                ) VALUES (1, ?, 5.9, 1.2, 800, 6, '28.07.2026', 1, ?)
                """,
                (PLATE, PLAN_ID),
            )
            conn.commit()

        summary = KpReadinessService(db_path=db).build_summary(1, status=KpStatus.IN_WORK.value)
        assert summary is not None
        assert summary.expected_sgp_date is None
        assert summary.expected_sgp_date_label is None
        assert summary.fully_scheduled is False
