from __future__ import annotations

import sqlite3

import pytest

from core.kp.offers_read import get_kp_by_id
from core.kp.plates_resolve import aggregate_completed_plates, resolve_plates_for_kp_documents
from core.kp_order_data import order_data_from_kp_info
from tests.helpers.kp_db_fixtures import make_iso_db, seed_kp_offer


def test_aggregate_completed_plates_sums_qty_by_dimensions() -> None:
    rows = [
        {
            "plate_name": "Плиты ПБ 67-15-12п",
            "length_m": 6.7,
            "width_m": 1.5,
            "load_class": 1200,
            "qty": 2,
        },
        {
            "plate_name": "Плиты ПБ 67-15-12п",
            "length_m": 6.7,
            "width_m": 1.5,
            "load_class": 1200,
            "qty": 1,
        },
        {
            "plate_name": "Плиты ПБ 39-5-12п",
            "length_m": 3.9,
            "width_m": 0.5,
            "load_class": 1200,
            "qty": 4,
        },
    ]

    aggregated = aggregate_completed_plates(rows)

    assert len(aggregated) == 2
    assert aggregated[0]["plate_name"] == "Плиты ПБ 67-15-12п"
    assert aggregated[0]["qty"] == 3
    assert aggregated[0]["position_number"] == 1
    assert aggregated[1]["qty"] == 4


def test_get_kp_by_id_falls_back_to_completed_plates(tmp_path) -> None:
    db_path = make_iso_db(tmp_path)
    seed_kp_offer(db_path, 1, status="выполнено")

    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class, qty,
                completed_date, production_day
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (1, "Плиты ПБ 67-15-12п", 6.7, 1.5, 1200, 2, "02.06.2026", 1),
        )
        conn.commit()

    kp_info = get_kp_by_id(1, db_path)

    assert kp_info is not None
    assert len(kp_info["plates"]) == 1
    assert kp_info["plates"][0]["qty"] == 2
    assert kp_info["plates"][0]["status"] == "выполнено"


def test_resolve_plates_prefers_active_kp_plates(tmp_path) -> None:
    db_path = make_iso_db(tmp_path)
    active = [{"plate_name": "ПБ", "length_m": 5.0, "width_m": 1.2, "load_class": 800, "qty": 1}]

    resolved = resolve_plates_for_kp_documents(active, kp_id=1, db_path=db_path)

    assert resolved == active


def test_order_data_from_completed_plates_omits_zero_unit_price() -> None:
    kp_info = {
        "discount_percent": 0,
        "plates": [
            {
                "plate_name": "Плиты ПБ 67-15-12п",
                "length_m": 6.7,
                "width_m": 1.5,
                "load_class": 1200,
                "qty": 2,
            }
        ],
    }

    order_data = order_data_from_kp_info(kp_info)

    assert len(order_data) == 1
    assert "unit_price" not in order_data[0]


def test_enrich_order_data_prices_from_xlsx(tmp_path) -> None:
    from core.commercial_offer_xlsx import generate_commercial_offer_xlsx
    from core.kp.xlsx_order_data import enrich_order_data_prices_from_xlsx

    source = [
        {
            "name": "Плиты ПБ 67-15-12п",
            "length_m": 6.7,
            "width_m": 1.5,
            "qty": 2,
            "load_class": 1200,
            "weight": 0,
        }
    ]
    xlsx_buffer = generate_commercial_offer_xlsx(
        [
            {
                "name": "Плиты ПБ 67-15-12п",
                "length_m": 6.7,
                "width_m": 1.5,
                "qty": 2,
                "load_class": 1200,
                "unit_price": 34125.0,
                "weight": 5695,
            }
        ],
        "1",
        "01.06.2026",
        customer_name="Тест",
    )
    enriched = enrich_order_data_prices_from_xlsx(source, xlsx_buffer.getvalue(), discount_percent=0)

    assert enriched[0]["unit_price"] == pytest.approx(34125.0)
