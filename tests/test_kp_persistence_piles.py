"""PILE-301: save pile KP to kp_piles with shared kp_id sequence."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd
import pytest

from core.kp_persistence_service import KpPersistenceService
from core.kp_db_schema import init_schema
from core.pile_price_db import import_pile_prices_from_xlsx


def _write_sample_pile_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [None, "35 СЕЧЕНИЕ", None, None, None, None, None],
        [69, "С120.35-12", 43760.31, 44108.15, 44371.09, 44634.03, 46159.37],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


@pytest.fixture()
def db_path(tmp_path: Path) -> str:
    path = str(tmp_path / "plita.db")
    init_schema(path)
    return path


def _pile_order_item(**overrides: object) -> dict:
    base = {
        "product_kind": "pile",
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "concrete_grade": "B25",
        "qty": 2,
        "unit_price": 44634.03,
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


def test_save_pile_kp_persists_kp_piles(db_path: str) -> None:
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        [_pile_order_item()],
        customer_name="Pile Client",
        product_type="piles",
        db_path=db_path,
    )
    assert kp_id == 1

    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == "piles"
        cur.execute("SELECT COUNT(*) FROM kp_plates WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == 0
        cur.execute(
            "SELECT mark, concrete_grade, qty FROM kp_piles WHERE kp_id = ?",
            (kp_id,),
        )
        row = cur.fetchone()
        assert row["mark"] == "С120.35-12"
        assert row["concrete_grade"] == "B25"
        assert row["qty"] == 2


def test_save_pile_kp_gets_next_kp_id_after_plates(db_path: str) -> None:
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
        [_pile_order_item(qty=5)],
        customer_name="Pile 1",
        product_type="piles",
        db_path=db_path,
    )

    assert first == 1
    assert second == 2
    assert third == 3

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM kp_piles WHERE kp_id = 3")
        assert cur.fetchone()[0] == 1
        cur.execute("SELECT COUNT(*) FROM kp_plates WHERE kp_id = 3")
        assert cur.fetchone()[0] == 0
