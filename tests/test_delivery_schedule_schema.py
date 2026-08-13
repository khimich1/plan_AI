"""Delivery schedule schema (T1): delivery_schedule / delivery_batch / delivery_batch_item."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core import kp_db_schema
from core.kp_db_common import _connect

_TS = "2026-08-07T12:00:00"


def _fresh_db(tmp_path: Path, name: str = "test.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _seed_kp_with_plate(conn: sqlite3.Connection, kp_id: int = 1) -> int:
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO KP_offers (kp_id, creation_date) VALUES (?, '2026-08-07')",
        (kp_id,),
    )
    cur.execute(
        """
        INSERT INTO kp_plates (kp_id, position_number, plate_name, qty)
        VALUES (?, 1, 'ПБ 60-12-8п', 10)
        """,
        (kp_id,),
    )
    return int(cur.lastrowid)


def _seed_schedule(
    conn: sqlite3.Connection, kp_id: int, plate_id: int
) -> tuple[int, int, int]:
    """schedule → batch → item chain; returns (schedule_id, batch_id, item_id)."""
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO delivery_schedule (
            kp_id, invoice_number, contract_number, status, created_at, updated_at
        ) VALUES (?, 'СЧ-101', 'Д-5', 'draft', ?, ?)
        """,
        (kp_id, _TS, _TS),
    )
    schedule_id = int(cur.lastrowid)
    cur.execute(
        """
        INSERT INTO delivery_batch (
            schedule_id, name, deliver_from, deliver_to, produce_by, sort_order
        ) VALUES (?, '1 этаж', '2026-09-01', '2026-09-10', '2026-08-25', 1)
        """,
        (schedule_id,),
    )
    batch_id = int(cur.lastrowid)
    cur.execute(
        "INSERT INTO delivery_batch_item (batch_id, plate_id, qty) VALUES (?, ?, 3)",
        (batch_id, plate_id),
    )
    return schedule_id, batch_id, int(cur.lastrowid)


def _count(conn: sqlite3.Connection, table: str) -> int:
    return int(conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])


def test_ensure_schema_idempotent_with_delivery_schedule(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)  # second pass must not fail
    kp_db_schema._init_schema_impl(db_path)  # direct re-init must not fail either


def test_delivery_tables_have_expected_columns(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "fresh.db")

    with _connect(db_path) as conn:
        cur = conn.cursor()

        cur.execute("PRAGMA table_info(delivery_schedule)")
        schedule_cols = {row[1]: row for row in cur.fetchall()}
        assert schedule_cols.keys() == {
            "id",
            "kp_id",
            "invoice_number",
            "contract_number",
            "status",
            "created_at",
            "updated_at",
        }
        assert schedule_cols["kp_id"][3] == 1  # NOT NULL
        assert schedule_cols["status"][4] == "'draft'"  # DEFAULT 'draft'

        cur.execute("PRAGMA table_info(delivery_batch)")
        batch_cols = {row[1]: row for row in cur.fetchall()}
        assert batch_cols.keys() == {
            "id",
            "schedule_id",
            "name",
            "deliver_from",
            "deliver_to",
            "produce_by",
            "sort_order",
        }
        assert batch_cols["schedule_id"][3] == 1  # NOT NULL
        assert batch_cols["sort_order"][4] == "0"  # DEFAULT 0

        cur.execute("PRAGMA table_info(delivery_batch_item)")
        item_cols = {row[1]: row for row in cur.fetchall()}
        assert item_cols.keys() == {"id", "batch_id", "plate_id", "qty"}
        assert item_cols["qty"][3] == 1  # NOT NULL


def test_delivery_schedule_unique_kp_id(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "unique.db")

    with _connect(db_path) as conn:
        _seed_kp_with_plate(conn, kp_id=1)
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO delivery_schedule (kp_id, status, created_at, updated_at)
            VALUES (1, 'draft', ?, ?)
            """,
            (_TS, _TS),
        )
        conn.commit()

        with pytest.raises(sqlite3.IntegrityError):
            cur.execute(
                """
                INSERT INTO delivery_schedule (kp_id, status, created_at, updated_at)
                VALUES (1, 'active', ?, ?)
                """,
                (_TS, _TS),
            )
        conn.rollback()

        assert _count(conn, "delivery_schedule") == 1


def test_delivery_schedule_fk_requires_existing_kp(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "fk.db")

    with _connect(db_path) as conn:
        cur = conn.cursor()
        with pytest.raises(sqlite3.IntegrityError):
            cur.execute(
                """
                INSERT INTO delivery_schedule (kp_id, status, created_at, updated_at)
                VALUES (999, 'draft', ?, ?)
                """,
                (_TS, _TS),
            )
        conn.rollback()


def test_delivery_schedule_defaults(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "defaults.db")

    with _connect(db_path) as conn:
        _seed_kp_with_plate(conn, kp_id=1)
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO delivery_schedule (kp_id, created_at, updated_at)
            VALUES (1, ?, ?)
            """,
            (_TS, _TS),
        )
        schedule_id = int(cur.lastrowid)
        cur.execute(
            """
            INSERT INTO delivery_batch (
                schedule_id, name, deliver_from, deliver_to, produce_by
            ) VALUES (?, 'партия', '2026-09-01', '2026-09-10', '2026-08-25')
            """,
            (schedule_id,),
        )
        conn.commit()

        cur.execute(
            "SELECT status, invoice_number, contract_number "
            "FROM delivery_schedule WHERE id = ?",
            (schedule_id,),
        )
        assert cur.fetchone() == ("draft", None, None)

        cur.execute(
            "SELECT sort_order FROM delivery_batch WHERE schedule_id = ?",
            (schedule_id,),
        )
        assert cur.fetchone() == (0,)


def test_delete_kp_cascades_to_schedule_batches_items(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "cascade_kp.db")

    with _connect(db_path) as conn:
        plate_id = _seed_kp_with_plate(conn, kp_id=1)
        _seed_schedule(conn, kp_id=1, plate_id=plate_id)
        conn.commit()
        assert _count(conn, "delivery_schedule") == 1
        assert _count(conn, "delivery_batch") == 1
        assert _count(conn, "delivery_batch_item") == 1

        conn.execute("DELETE FROM KP_offers WHERE kp_id = 1")
        conn.commit()

        assert _count(conn, "delivery_schedule") == 0
        assert _count(conn, "delivery_batch") == 0
        assert _count(conn, "delivery_batch_item") == 0


def test_delete_batch_cascades_only_to_own_items(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "cascade_batch.db")

    with _connect(db_path) as conn:
        plate_id = _seed_kp_with_plate(conn, kp_id=1)
        schedule_id, batch_id, _ = _seed_schedule(conn, kp_id=1, plate_id=plate_id)
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO delivery_batch (
                schedule_id, name, deliver_from, deliver_to, produce_by
            ) VALUES (?, '2 этаж', '2026-10-01', '2026-10-10', '2026-09-25')
            """,
            (schedule_id,),
        )
        other_batch_id = int(cur.lastrowid)
        cur.execute(
            "INSERT INTO delivery_batch_item (batch_id, plate_id, qty) VALUES (?, ?, 2)",
            (other_batch_id, plate_id),
        )
        conn.commit()

        conn.execute("DELETE FROM delivery_batch WHERE id = ?", (batch_id,))
        conn.commit()

        assert _count(conn, "delivery_schedule") == 1
        assert _count(conn, "delivery_batch") == 1
        assert _count(conn, "delivery_batch_item") == 1
        cur.execute("SELECT batch_id, qty FROM delivery_batch_item")
        assert cur.fetchone() == (other_batch_id, 2)


def test_delete_plate_cascades_to_batch_items(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "cascade_plate.db")

    with _connect(db_path) as conn:
        plate_id = _seed_kp_with_plate(conn, kp_id=1)
        schedule_id, batch_id, _ = _seed_schedule(conn, kp_id=1, plate_id=plate_id)
        conn.commit()

        conn.execute("DELETE FROM kp_plates WHERE id = ?", (plate_id,))
        conn.commit()

        assert _count(conn, "delivery_batch_item") == 0
        assert _count(conn, "delivery_batch") == 1
        assert _count(conn, "delivery_schedule") == 1


@pytest.mark.parametrize("bad_qty", [0, -1, -100])
def test_delivery_batch_item_qty_check_rejects_non_positive(
    tmp_path: Path, bad_qty: int
) -> None:
    db_path = _fresh_db(tmp_path, f"check_qty_{bad_qty}.db")

    with _connect(db_path) as conn:
        plate_id = _seed_kp_with_plate(conn, kp_id=1)
        _, batch_id, _ = _seed_schedule(conn, kp_id=1, plate_id=plate_id)
        conn.commit()

        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                "INSERT INTO delivery_batch_item (batch_id, plate_id, qty) "
                "VALUES (?, ?, ?)",
                (batch_id, plate_id, bad_qty),
            )
        conn.rollback()


def test_delivery_batch_item_qty_allows_one(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "check_qty_ok.db")

    with _connect(db_path) as conn:
        plate_id = _seed_kp_with_plate(conn, kp_id=1)
        _, batch_id, _ = _seed_schedule(conn, kp_id=1, plate_id=plate_id)
        conn.execute(
            "INSERT INTO delivery_batch_item (batch_id, plate_id, qty) "
            "VALUES (?, ?, 1)",
            (batch_id, plate_id),
        )
        conn.commit()

        assert _count(conn, "delivery_batch_item") == 2


def test_delivery_indexes_exist(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "indexes.db")

    with _connect(db_path) as conn:
        cur = conn.cursor()

        def indexed_columns(table: str) -> dict[str, list[str]]:
            cur.execute(f"PRAGMA index_list({table})")
            result: dict[str, list[str]] = {}
            for row in cur.fetchall():
                index_name = row[1]
                cur.execute(f"PRAGMA index_info({index_name})")
                result[index_name] = [info[2] for info in cur.fetchall()]
            return result

        batch_indexes = indexed_columns("delivery_batch")
        assert batch_indexes.get("idx_delivery_batch_schedule") == ["schedule_id"]

        item_indexes = indexed_columns("delivery_batch_item")
        assert item_indexes.get("idx_delivery_batch_item_batch") == ["batch_id"]
        assert item_indexes.get("idx_delivery_batch_item_plate") == ["plate_id"]
