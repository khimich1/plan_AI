"""KpRepository.get_plate_qty_remaining (orch-2026-08-12-podlozhki Task 3)."""

from __future__ import annotations

from app.domain.enums import PlateStatus
from app.repositories.kp_repository import KpRepository
from tests.helpers import kp_db_fixtures as fx


def _repo(tmp_path) -> KpRepository:
    return KpRepository(db_path=fx.make_iso_db(tmp_path))


def test_production_plate_qty_remaining_equals_qty(tmp_path) -> None:
    repo = _repo(tmp_path)
    fx.seed_kp_offer(repo.db_path, 1)
    plate_id = fx.seed_plate(
        repo.db_path,
        kp_id=1,
        plate_name="ПК 60.15",
        length_m=6.0,
        width_m=1.5,
        qty=5,
        status=PlateStatus.IN_PRODUCTION.value,
    )
    assert repo.get_plate_qty_remaining(plate_id) == 5


def test_planned_plate_with_plan_id_remaining_zero(tmp_path) -> None:
    repo = _repo(tmp_path)
    fx.seed_kp_offer(repo.db_path, 1)
    plate_id = fx.seed_plate(
        repo.db_path,
        kp_id=1,
        plate_name="ПК 60.15",
        length_m=6.0,
        width_m=1.5,
        qty=4,
        status=PlateStatus.IN_PLAN.value,
        plan_id="plan-1",
    )
    assert repo.get_plate_qty_remaining(plate_id) == 0


def test_stuck_in_plan_without_plan_id_keeps_qty(tmp_path) -> None:
    repo = _repo(tmp_path)
    fx.seed_kp_offer(repo.db_path, 1)
    plate_id = fx.seed_plate(
        repo.db_path,
        kp_id=1,
        plate_name="ПК 60.15",
        length_m=6.0,
        width_m=1.5,
        qty=3,
        status=PlateStatus.IN_PLAN.value,
        plan_id=None,
    )
    assert repo.get_plate_qty_remaining(plate_id) == 3


def test_missing_plate_id_returns_zero(tmp_path) -> None:
    repo = _repo(tmp_path)
    assert repo.get_plate_qty_remaining(999_999) == 0


def test_zero_qty_production_remaining_zero(tmp_path) -> None:
    repo = _repo(tmp_path)
    fx.seed_kp_offer(repo.db_path, 1)
    plate_id = fx.seed_plate(
        repo.db_path,
        kp_id=1,
        plate_name="ПК 60.15",
        length_m=6.0,
        width_m=1.5,
        qty=0,
        status=PlateStatus.IN_PRODUCTION.value,
    )
    assert repo.get_plate_qty_remaining(plate_id) == 0


def test_split_model_planned_and_remainder_rows(tmp_path) -> None:
    """After conceptual split: planned row → 0, remainder → its qty."""
    repo = _repo(tmp_path)
    fx.seed_kp_offer(repo.db_path, 1)
    planned_id = fx.seed_plate(
        repo.db_path,
        kp_id=1,
        plate_name="ПК 60.15",
        length_m=6.0,
        width_m=1.5,
        qty=2,
        status=PlateStatus.IN_PLAN.value,
        plan_id="plan-split",
        position_number=1,
    )
    remainder_id = fx.seed_plate(
        repo.db_path,
        kp_id=1,
        plate_name="ПК 60.15",
        length_m=6.0,
        width_m=1.5,
        qty=3,
        status=PlateStatus.IN_PRODUCTION.value,
        position_number=2,
    )
    assert repo.get_plate_qty_remaining(planned_id) == 0
    assert repo.get_plate_qty_remaining(remainder_id) == 3
