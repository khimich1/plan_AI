"""SGP-000: enums + schema migration for completed_plates / kp_meta.ordered_qty."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from app.domain.enums import KpStatus, PlateStatus, PlateTransitionReason
from core import kp_db_schema


def test_sgp_enums_values() -> None:
    assert PlateStatus.ON_SGP.value == "on_sgp"
    assert KpStatus.ON_SGP.value == "На СГП"
    assert PlateTransitionReason.SGP_SEND.value == "sgp_send"
    assert PlateTransitionReason.SGP_UNLINK.value == "sgp_unlink"
    assert PlateTransitionReason.SGP_RELINK.value == "sgp_relink"
    assert PlateTransitionReason.SGP_RESERVE.value == "sgp_reserve"


def test_fresh_schema_allows_null_kp_id_and_has_ordered_qty(tmp_path: Path) -> None:
    db_path = str(tmp_path / "fresh.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.cursor()

        cur.execute("PRAGMA table_info(completed_plates)")
        cp_cols = {row[1]: row for row in cur.fetchall()}
        assert "kp_id" in cp_cols
        assert cp_cols["kp_id"][3] == 0  # notnull == 0 → nullable
        assert "plan_id" in cp_cols

        cur.execute("PRAGMA table_info(kp_meta)")
        meta_cols = {row[1] for row in cur.fetchall()}
        assert "ordered_qty" in meta_cols

        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (1, '2026-01-01')"
        )
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, plan_id
            ) VALUES (NULL, 'ПБ 60-12-8п', 6.0, 1.2, 800, 2, '2026-07-27', 1, NULL)
            """
        )
        conn.commit()
        cur.execute("SELECT kp_id, qty, plan_id FROM completed_plates")
        row = cur.fetchone()
        assert row == (None, 2, None)


def test_migrate_existing_completed_plates_makes_kp_id_nullable(tmp_path: Path) -> None:
    """Existing DB with NOT NULL kp_id + CASCADE → rebuild to nullable + SET NULL."""
    db_path = str(tmp_path / "legacy.db")
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE KP_offers (
                kp_id INTEGER PRIMARY KEY AUTOINCREMENT,
                creation_date TEXT NOT NULL
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE kp_meta (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL UNIQUE,
                status TEXT DEFAULT 'в работе',
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE completed_plates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                plate_name TEXT NOT NULL,
                length_m REAL,
                width_m REAL,
                load_class INTEGER,
                qty INTEGER NOT NULL,
                completed_date TEXT NOT NULL,
                production_day INTEGER,
                nomenclature_id TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (10, '2026-01-01')"
        )
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, nomenclature_id
            ) VALUES (10, 'ПБ 60-12-8п', 6.0, 1.2, 800, 3, '2026-07-01', 1, 'n1')
            """
        )
        conn.commit()

    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(completed_plates)")
        cp_cols = {row[1]: row for row in cur.fetchall()}
        assert cp_cols["kp_id"][3] == 0
        assert "plan_id" in cp_cols

        cur.execute(
            "SELECT kp_id, plate_name, qty, nomenclature_id FROM completed_plates"
        )
        assert cur.fetchone() == (10, "ПБ 60-12-8п", 3, "n1")

        # INSERT with NULL kp_id must succeed after migration
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, qty, completed_date
            ) VALUES (NULL, 'free', 1, '2026-07-27')
            """
        )
        conn.commit()

        cur.execute("PRAGMA table_info(kp_meta)")
        assert "ordered_qty" in {row[1] for row in cur.fetchall()}


def test_ensure_schema_idempotent(tmp_path: Path) -> None:
    db_path = str(tmp_path / "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)  # second pass must not fail
