"""MARCH-301: save march KP to kp_marches with shared kp_id sequence."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd
import pytest

from core.kp_persistence_service import KpPersistenceService
from core.kp_db_schema import init_schema
from core.march_price_db import import_march_prices_from_xlsx


def _write_sample_march_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [1, "Лестничные марши 1ЛМ 27-11-14-4", 13993.72, 14150.79, 14271.10, 14391.41, 14639.53],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


@pytest.fixture()
def db_path(tmp_path: Path) -> str:
    path = str(tmp_path / "plita.db")
    init_schema(path)
    return path


def _march_order_item(**overrides: object) -> dict:
    base = {
        "product_kind": "march",
        "name": "1ЛМ 27-11-14-4",
        "mark": "1ЛМ 27-11-14-4",
        "concrete_grade": "B25",
        "qty": 2,
        "unit_price": 14391.41,
    }
    base.update(overrides)
    return base


def _plate_order_item(**overrides: object) -> dict:
    base = {
        "name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 1,
        "unit_price": 1000.0,
        "weight": 500.0,
        "length_dm_raw": "60",
    }
    base.update(overrides)
    return base


def test_save_march_kp_persists_kp_marches(db_path: str) -> None:
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        [_march_order_item()],
        customer_name="March Client",
        product_type="marches",
        db_path=db_path,
    )
    assert kp_id == 1

    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == "marches"
        cur.execute("SELECT COUNT(*) FROM kp_plates WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == 0
        cur.execute(
            "SELECT mark, concrete_grade, qty FROM kp_marches WHERE kp_id = ?",
            (kp_id,),
        )
        row = cur.fetchone()
        assert row["mark"] == "1ЛМ 27-11-14-4"
        assert row["concrete_grade"] == "B25"
        assert row["qty"] == 2


def test_save_march_kp_gets_next_kp_id_after_plates(db_path: str) -> None:
    first = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        [_plate_order_item()],
        customer_name="Plate 1",
        db_path=db_path,
    )
    second = KpPersistenceService.save_kp_to_db(
        "02.01.2026",
        [_plate_order_item(qty=3)],
        customer_name="Plate 2",
        db_path=db_path,
    )
    third = KpPersistenceService.save_kp_to_db(
        "03.01.2026",
        [_march_order_item(qty=5)],
        customer_name="March 1",
        product_type="marches",
        db_path=db_path,
    )

    assert first == 1
    assert second == 2
    assert third == 3

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM kp_marches WHERE kp_id = 3")
        assert cur.fetchone()[0] == 1
        cur.execute("SELECT COUNT(*) FROM kp_plates WHERE kp_id = 3")
        assert cur.fetchone()[0] == 0
