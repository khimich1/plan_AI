"""MARCH-002: kp_marches table (with concrete_grade)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from core import kp_db_schema


def test_fresh_schema_has_kp_marches(tmp_path: Path) -> None:
    db_path = str(tmp_path / "fresh.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()

        cur.execute("PRAGMA table_info(kp_marches)")
        cols = {row[1]: row for row in cur.fetchall()}
        assert cols.keys() >= {
            "id",
            "kp_id",
            "position_number",
            "mark",
            "concrete_grade",
            "qty",
            "unit_price",
            "discounted_price",
        }

        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (1, '2026-08-05')"
        )
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status, product_type) VALUES (1, 'в архиве', 'marches')"
        )
        cur.execute(
            """
            INSERT INTO kp_marches (
                kp_id, position_number, mark, concrete_grade,
                qty, unit_price, discounted_price
            ) VALUES (1, 1, '1ЛМ 27-11-14-4', 'B25', 2, 14391.41, 14391.41)
            """
        )
        conn.commit()

        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = 1")
        assert cur.fetchone()[0] == "marches"

        cur.execute("SELECT mark, concrete_grade, qty FROM kp_marches WHERE kp_id = 1")
        assert cur.fetchone() == ("1ЛМ 27-11-14-4", "B25", 2)


def test_ensure_schema_idempotent_with_kp_marches(tmp_path: Path) -> None:
    db_path = str(tmp_path / "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='kp_marches'"
        )
        assert cur.fetchone() is not None
