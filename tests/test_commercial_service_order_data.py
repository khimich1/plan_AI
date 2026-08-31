from __future__ import annotations

from pathlib import Path

import pytest

from app.domain.models.parse_result import ParseResult
from app.domain.models.plate_order import PlateOrder
from app.services.commercial_calculation_service import CommercialCalculationService
from app.services.commercial_service import CommercialService
from core.commercial_pricing import collect_unpriced_positions


def _price_row(*, name: str, price: str | float) -> list:
    # [idx, name, qty, unit, week, contractor, weight, price, sum]
    return [1, name, 1, "шт", "", "", 0, price, 0]


def _build_one(
    *,
    length_m: float = 7.5,
    width_m: float = 1.2,
    qty: int = 1,
    load_code: int = 12,
    length_dm_raw: str = "75",
    price_rows: list | None = None,
) -> list[dict]:
    service = CommercialService()
    order = PlateOrder()
    parse_result = ParseResult(
        order=order,
        normalized_text="",
        line_plate_load_details=[
            {(length_m, width_m, float(load_code), length_dm_raw): qty},
        ],
    )
    procurement_items = [
        {
            "length": length_m,
            "width": width_m,
            "qty": qty,
            "load_code": load_code,
            "length_dm_raw": length_dm_raw,
        }
    ]
    return service._build_order_data(
        procurement_items,
        price_rows or [],
        order,
        parse_result,
    )


def test_build_order_data_zero_estimate_price_becomes_none() -> None:
    price_rows = [_price_row(name="ПБ 75-12-12п", price="0,00")]
    order_data = _build_one(price_rows=price_rows)

    assert len(order_data) == 1
    assert order_data[0]["unit_price"] is None
    assert order_data[0]["name"] == "ПБ 75-12-12п"


def test_build_order_data_missing_matching_row_unit_price_none() -> None:
    order_data = _build_one(price_rows=[])

    assert len(order_data) == 1
    assert order_data[0]["unit_price"] is None


def test_build_order_data_positive_price_is_float() -> None:
    price_rows = [_price_row(name="ПБ 70-12-12п", price="29 210,00")]
    order_data = _build_one(
        length_m=7.0,
        length_dm_raw="70",
        price_rows=price_rows,
    )

    assert order_data[0]["unit_price"] == pytest.approx(29210.0)


def test_build_order_data_unparseable_price_becomes_none() -> None:
    price_rows = [_price_row(name="ПБ 75-12-12п", price="не число")]
    order_data = _build_one(price_rows=price_rows)

    assert order_data[0]["unit_price"] is None


def test_unpriced_position_labels_finds_zero_priced_plate(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import sqlite3

    from core import commercial_offer_xlsx

    db_path = tmp_path / "pb.db"
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            "CREATE TABLE prices (length_dm INTEGER, load_code INTEGER, price REAL, "
            "PRIMARY KEY(length_dm, load_code))"
        )
        conn.execute(
            "INSERT INTO prices (length_dm, load_code, price) VALUES (75, 12, 0.0)"
        )
        conn.commit()
    finally:
        conn.close()

    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", str(db_path))
    monkeypatch.setattr(
        "app.services.commercial_calculation_service.DB_PATH",
        str(db_path),
    )

    order_data = [
        {
            "name": "ПБ 75-12-12п",
            "qty": 1,
            "length_m": 7.5,
            "width_m": 1.2,
            "load_class": 1200,
            "unit_price": None,
        }
    ]

    labels = CommercialCalculationService().unpriced_position_labels(order_data)
    assert "ПБ 75-12-12п" in labels

    assert collect_unpriced_positions(order_data, db_path=str(db_path)) == ["ПБ 75-12-12п"]
