"""Проверки update_kp_logistics_cost: суммы в SQLite совпадают с calculate_total_cost (XLSX)."""

from __future__ import annotations

import sqlite3

import pytest

from core import kp_db
from core.commercial_offer_xlsx import calculate_total_cost


@pytest.fixture()
def iso_db(tmp_path_factory: pytest.TempPathFactory) -> str:
    db_path = str(tmp_path_factory.mktemp("kp_logistics") / "plita.db")
    kp_db.init_schema(db_path)
    return db_path


def _offer_financial_row(db_path: str, kp_id: int) -> tuple[float, float, float, float]:
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT logistics_cost, subtotal, vat_amount, total_amount FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        )
        row = cur.fetchone()
        assert row is not None
        return (float(row[0]), float(row[1]), float(row[2]), float(row[3]))
    finally:
        conn.close()


def test_update_kp_logistics_cost_totals_match_calculate_total_cost(iso_db: str) -> None:
    """После PATCH-логики пересчёта строки KP_offers согласованы с тем же calculate_total_cost, что и при сохранении."""
    order_data = [
        {
            "name": "ПБ",
            "qty": 66,
            "unit_price": 10.0,
            "length_m": 1.0,
            "width_m": 1.0,
        }
    ]
    kp_id = kp_db.save_kp_to_db(
        "01.03.2026",
        order_data,
        discount_percent=0.0,
        logistics_cost=100.0,
        db_path=iso_db,
        status="в архиве",
    )
    assert kp_db.update_kp_logistics_cost(kp_id, 125.55, iso_db) is True

    expected = calculate_total_cost(order_data, 0.0, logistics_cost=125.55)
    lg, sub, vat, total = _offer_financial_row(iso_db, kp_id)
    assert lg == pytest.approx(125.55)
    assert sub == pytest.approx(expected["subtotal"])
    assert vat == pytest.approx(expected["vat_amount"])
    assert total == pytest.approx(expected["total_with_vat"])


def test_update_kp_logistics_cost_with_discount_keeps_pdf_xlsx_formula(iso_db: str) -> None:
    order_data = [
        {
            "name": "ПБ 59-12-8п",
            "qty": 3,
            "unit_price": 100.0,
            "length_m": 1.0,
            "width_m": 1.0,
        }
    ]
    kp_id = kp_db.save_kp_to_db(
        "02.03.2026",
        order_data,
        discount_percent=15.0,
        logistics_cost=0.0,
        db_path=iso_db,
        status="в архиве",
    )
    trip = 2000.5
    assert kp_db.update_kp_logistics_cost(kp_id, trip, iso_db) is True

    expected = calculate_total_cost(order_data, 15.0, logistics_cost=trip)
    lg, sub, vat, total = _offer_financial_row(iso_db, kp_id)
    assert lg == pytest.approx(trip)
    assert sub == pytest.approx(expected["subtotal"])
    assert vat == pytest.approx(expected["vat_amount"])
    assert total == pytest.approx(expected["total_with_vat"])


def test_update_kp_logistics_cost_clamps_negative_and_matches_zero_trip_formula(iso_db: str) -> None:
    order_data = [
        {"name": "ПБ", "qty": 10, "unit_price": 5.0, "length_m": 1.0, "width_m": 1.0},
    ]
    kp_id = kp_db.save_kp_to_db(
        "03.03.2026",
        order_data,
        discount_percent=0.0,
        logistics_cost=50.0,
        db_path=iso_db,
        status="в архиве",
    )
    assert kp_db.update_kp_logistics_cost(kp_id, -99.0, iso_db) is True

    expected = calculate_total_cost(order_data, 0.0, logistics_cost=0.0)
    lg, sub, vat, total = _offer_financial_row(iso_db, kp_id)
    assert lg == pytest.approx(0.0)
    assert sub == pytest.approx(expected["subtotal"])
    assert vat == pytest.approx(expected["vat_amount"])
    assert total == pytest.approx(expected["total_with_vat"])


def test_update_kp_logistics_cost_returns_false_when_kp_missing(iso_db: str) -> None:
    assert kp_db.update_kp_logistics_cost(987654, 100.0, iso_db) is False
