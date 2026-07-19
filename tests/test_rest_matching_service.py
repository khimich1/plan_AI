"""Unit tests for RestMatchingService (A2 rests slice)."""

from __future__ import annotations

import pytest

from app.services.rest_matching_service import RestMatchingService
from core import kp_db
from core.domain.rest_matching import LONG_CUT_PRICE_PER_M, TRANSVERSE_CUT_PRICE
from tests.helpers import kp_db_fixtures as fx


@pytest.fixture
def iso_db(tmp_path):
    return fx.make_iso_db(tmp_path)


def _seed_exact(iso_db: str, *, qty: int = 1) -> None:
    fx.seed_kp_offer(iso_db, 1)
    kp_db.create_plate_rest(1, "ПБ 60-12-8п", 1200, 6.0, 1, qty=qty, db_path=iso_db)


def test_service_exact_match_zero_cut(iso_db: str) -> None:
    _seed_exact(iso_db)
    matches = RestMatchingService.find_matching_rests(6.0, 1200, 1, db_path=iso_db)
    assert len(matches) == 1
    assert matches[0]["match_type"] == "exact"
    assert matches[0]["cut_cost"] == 0.0


def test_service_width_cut_cost(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    kp_db.create_plate_rest(1, "W", 1500, 6.0, 1, qty=1, db_path=iso_db)
    matches = RestMatchingService.find_matching_rests(6.0, 1200, 1, db_path=iso_db)
    assert matches[0]["match_type"] == "width_cut"
    assert matches[0]["cut_cost"] == pytest.approx(LONG_CUT_PRICE_PER_M * 6.0)


def test_service_qty_cap_across_rests(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    kp_db.create_plate_rest(1, "A", 1200, 6.0, 1, qty=1, db_path=iso_db)
    kp_db.create_plate_rest(1, "B", 1200, 6.0, 1, qty=2, db_path=iso_db)
    matches = RestMatchingService.find_matching_rests(6.0, 1200, 3, db_path=iso_db)
    assert sum(m["qty_to_use"] for m in matches) == 3
