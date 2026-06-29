"""Regression tests for ``move_plates_to_completed`` / ``find_one_row`` heuristics."""

from __future__ import annotations

import pytest

from core import kp_db
from tests.helpers import kp_db_fixtures as fx

PLATE_60 = "ПБ 60-12-8п"
PLATE_598 = "ПБ 59,8-12-8п"
PLATE_599 = "ПБ 59,9-12-8п"
PLATE_611 = "ПБ 61,1-12-8п"
PLATE_612 = "ПБ 61,2-12-8п"


@pytest.fixture
def iso_db(tmp_path):
    return fx.make_iso_db(tmp_path)


def _complete(
    db: str,
    kp_id: int,
    plate_name: str,
    *,
    qty: int = 1,
    length_m: float = 6.0,
    width_m: float = 1.2,
    load_class: int = 800,
    length_dm_raw: str = "",
    kp_plate_id: int | None = None,
    allow_cross_kp: bool = False,
    plan_ids: list[str] | None = None,
    return_unmoved: bool = False,
):
    payload: dict = {
        "plate_name": plate_name,
        "qty": qty,
        "length_m": length_m,
        "width_m": width_m,
        "load_class": load_class,
        "length_dm_raw": length_dm_raw,
    }
    if kp_plate_id is not None:
        payload["kp_plate_id"] = kp_plate_id
    return kp_db.move_plates_to_completed(
        kp_id,
        [payload],
        production_day=1,
        db_path=db,
        allow_cross_kp=allow_cross_kp,
        plan_ids=plan_ids,
        return_unmoved=return_unmoved,
    )


def test_g0_length_dm_raw_distinguishes_rows(iso_db: str) -> None:
    """Step 0: same KP, different length_dm_raw — only matching row is deducted."""
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE_60,
        length_m=5.98,
        width_m=1.2,
        qty=2,
        status="в плане",
        length_dm_raw="598",
    )
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE_60,
        length_m=5.99,
        width_m=1.2,
        qty=3,
        status="в плане",
        length_dm_raw="599",
    )

    count = _complete(
        iso_db,
        1,
        PLATE_60,
        qty=2,
        length_m=5.98,
        length_dm_raw="598",
    )

    assert count == 2
    rows = fx.plates_snapshot(iso_db, 1)
    assert len(rows) == 1
    assert rows[0]["length_dm_raw"] == "599"
    assert rows[0]["qty"] == 3
    completed = fx.completed_snapshot(iso_db, 1)
    assert sum(r["qty"] for r in completed) == 2


def test_g1_exact_plate_name(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE_60,
        length_m=6.0,
        width_m=1.2,
        qty=2,
        status="в плане",
    )

    count = _complete(iso_db, 1, PLATE_60, qty=2, length_m=6.0)

    assert count == 2
    assert fx.total_plate_qty(iso_db, 1) == 0
    assert len(fx.completed_snapshot(iso_db, 1)) == 1


def test_g1_5_canonical_plate_prefix(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=f"Плиты {PLATE_60}",
        length_m=6.0,
        width_m=1.2,
        qty=1,
        status="в плане",
    )

    count = _complete(iso_db, 1, PLATE_60, qty=1, length_m=6.0)

    assert count == 1
    assert fx.completed_snapshot(iso_db, 1)[0]["plate_name"] == f"Плиты {PLATE_60}"


def test_g2_5_equivalent_59_8_59_9(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE_598,
        length_m=5.98,
        width_m=1.2,
        qty=1,
        status="в плане",
    )

    count = _complete(iso_db, 1, PLATE_599, qty=1, length_m=5.98)

    assert count == 1
    assert fx.completed_snapshot(iso_db, 1)[0]["plate_name"] == PLATE_598


def test_g_kp_plate_id_direct_lookup_wrong_name(iso_db: str) -> None:
    """P5: kp_plate_id selects row by id even when plate_name would not match (post-S4)."""
    fx.seed_kp_offer(iso_db, 1)
    plate_id = fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE_612,
        length_m=6.11,
        width_m=1.2,
        qty=1,
        status="в плане",
        day_number=1,
    )

    count = _complete(
        iso_db,
        1,
        PLATE_611,
        qty=1,
        length_m=6.11,
        kp_plate_id=plate_id,
    )

    assert count == 1
    assert fx.total_plate_qty(iso_db, 1) == 0
    assert fx.completed_snapshot(iso_db, 1)[0]["plate_name"] == PLATE_612


def test_g_legacy_mismatch_kp_id_returns_unmoved(iso_db: str) -> None:
    """After S4 (step 2.55 removed): wrong prefer_kp_id without kp_plate_id → unmoved."""
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_kp_offer(iso_db, 2, customer_name="Другой")
    fx.seed_plate(
        iso_db,
        kp_id=2,
        plate_name=PLATE_611,
        length_m=6.11,
        width_m=1.2,
        qty=1,
        status="в плане",
    )

    result = _complete(
        iso_db,
        1,
        PLATE_611,
        qty=1,
        length_m=6.11,
        allow_cross_kp=False,
        return_unmoved=True,
    )

    assert result == (
        0,
        [
            {
                "kp_id": 1,
                "plate_name": PLATE_611,
                "qty": 1,
                "length_m": 6.11,
                "width_m": 1.2,
                "load_class": 800,
            }
        ],
    )
    assert fx.total_plate_qty(iso_db, 2) == 1


def test_g2_55_cross_kp_61_1_61_2_when_disabled(iso_db: str) -> None:
    """S4: cross-KP disabled when stock only in other KP (step 2.55 removed)."""
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_kp_offer(iso_db, 2, customer_name="Другой")
    # No stock in KP1 — only KP2 has matching plate
    fx.seed_plate(
        iso_db,
        kp_id=2,
        plate_name=PLATE_612,
        length_m=6.11,
        width_m=1.2,
        qty=1,
        status="в плане",
    )

    count = _complete(
        iso_db,
        1,
        PLATE_611,
        qty=1,
        length_m=6.11,
        allow_cross_kp=False,
    )

    assert count == 0
    assert fx.completed_snapshot(iso_db) == []
    assert fx.total_plate_qty(iso_db, 2) == 1


def test_g2_6_length_tolerance_same_kp(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name="ПБ 62-12-8п",
        length_m=6.10,
        width_m=1.2,
        qty=1,
        status="в плане",
    )

    count = _complete(iso_db, 1, "ПБ 62-12-8п", qty=1, length_m=6.11)

    assert count == 1


def test_g_width_filter_avoids_narrow_secondary(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name="ПБ 63-12-8п",
        length_m=6.3,
        width_m=1.2,
        qty=1,
        status="в плане",
        position_number=1,
    )
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name="ПБ 63-12-8п",
        length_m=6.3,
        width_m=0.32,
        qty=1,
        status="в плане",
        position_number=2,
    )

    count = _complete(iso_db, 1, "ПБ 63-12-8п", qty=1, length_m=6.3, width_m=1.2)

    assert count == 1
    rows = fx.plates_snapshot(iso_db, 1)
    remaining = [r for r in rows if r["qty"] > 0]
    assert len(remaining) == 1
    assert remaining[0]["width_m"] == pytest.approx(0.32)


def test_g_partial_qty_deduct(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE_60,
        length_m=6.0,
        width_m=1.2,
        qty=5,
        status="в плане",
    )

    count = _complete(iso_db, 1, PLATE_60, qty=2, length_m=6.0)

    assert count == 2
    assert fx.total_plate_qty(iso_db, 1) == 3
    assert fx.completed_snapshot(iso_db, 1)[0]["qty"] == 2


def test_g_not_found_returns_zero(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)

    count = _complete(iso_db, 1, PLATE_60, qty=1, length_m=6.0)

    assert count == 0
    assert fx.completed_snapshot(iso_db) == []


def test_g_not_found_return_unmoved(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)

    result = _complete(
        iso_db,
        1,
        PLATE_60,
        qty=2,
        length_m=6.0,
        return_unmoved=True,
    )

    assert result == (0, [{"kp_id": 1, "plate_name": PLATE_60, "qty": 2, "length_m": 6.0, "width_m": 1.2, "load_class": 800}])


def test_g_cross_kp_with_plan_ids(iso_db: str) -> None:
    plan_id = "plan_test_cross"
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_kp_offer(iso_db, 2, customer_name="Другой")
    fx.seed_plate(
        iso_db,
        kp_id=2,
        plate_name="ПБ 70-12-8п",
        length_m=7.0,
        width_m=1.2,
        qty=1,
        status="в плане",
        plan_id=plan_id,
    )

    count = _complete(
        iso_db,
        1,
        "ПБ 70-12-8п",
        qty=1,
        length_m=7.0,
        allow_cross_kp=True,
        plan_ids=[plan_id],
    )

    assert count == 1
    assert fx.completed_snapshot(iso_db)[0]["kp_id"] == 2


def test_status_in_production_also_completable(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE_60,
        length_m=6.0,
        width_m=1.2,
        qty=1,
        status="в производстве",
    )

    count = _complete(iso_db, 1, PLATE_60, qty=1, length_m=6.0)

    assert count == 1
