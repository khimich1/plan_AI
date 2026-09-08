"""Isolated SQLite fixtures for ``core.kp_db`` regression tests."""

from __future__ import annotations

import sqlite3
from typing import Any

from core import kp_db


def make_iso_db(tmp_path) -> str:
    """Create temp plita.db with schema initialized."""
    db_path = str(tmp_path / "plita.db")
    kp_db.init_schema(db_path)
    return db_path


def seed_kp_offer(
    db_path: str,
    kp_id: int,
    *,
    customer_name: str = "ТестКлиент",
    status: str = "в работе",
) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO KP_offers (kp_id, creation_date, execution_terms, customer_name)
            VALUES (?, '2026-01-01', '21.04.2026', ?)
            """,
            (kp_id, customer_name),
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status) VALUES (?, ?)",
            (kp_id, status),
        )
        conn.commit()


def seed_plate(
    db_path: str,
    *,
    kp_id: int,
    plate_name: str,
    length_m: float,
    width_m: float,
    load_class: int = 800,
    qty: int,
    status: str,
    length_dm_raw: str = "",
    nomenclature_id: int | None = None,
    plan_id: str | None = None,
    day_number: int | None = None,
    position_number: int = 1,
) -> int:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status, length_dm_raw, nomenclature_id,
                plan_id, day_number
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                kp_id,
                position_number,
                plate_name,
                length_m,
                width_m,
                load_class,
                qty,
                status,
                length_dm_raw or "",
                nomenclature_id,
                plan_id,
                day_number,
            ),
        )
        conn.commit()
        return int(cur.lastrowid)


def plates_snapshot(db_path: str, kp_id: int | None = None) -> list[dict[str, Any]]:
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        if kp_id is not None:
            cur.execute(
                """
                SELECT id, kp_id, plate_name, length_m, width_m, load_class, qty,
                       status, length_dm_raw, nomenclature_id, plan_id
                FROM kp_plates
                WHERE kp_id = ?
                ORDER BY id
                """,
                (kp_id,),
            )
        else:
            cur.execute(
                """
                SELECT id, kp_id, plate_name, length_m, width_m, load_class, qty,
                       status, length_dm_raw, nomenclature_id, plan_id
                FROM kp_plates
                ORDER BY id
                """
            )
        return [dict(row) for row in cur.fetchall()]


def completed_snapshot(db_path: str, kp_id: int | None = None) -> list[dict[str, Any]]:
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        if kp_id is not None:
            cur.execute(
                """
                SELECT kp_id, plate_name, length_m, width_m, load_class, qty,
                       production_day, nomenclature_id
                FROM completed_plates
                WHERE kp_id = ?
                ORDER BY id
                """,
                (kp_id,),
            )
        else:
            cur.execute(
                """
                SELECT kp_id, plate_name, length_m, width_m, load_class, qty,
                       production_day, nomenclature_id
                FROM completed_plates
                ORDER BY id
                """
            )
        return [dict(row) for row in cur.fetchall()]


def total_plate_qty(db_path: str, kp_id: int) -> int:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = ?",
            (kp_id,),
        )
        row = cur.fetchone()
        return int(row[0]) if row else 0


def seed_test_counterparty(db_path: str, *, name: str = "ООО Тест") -> int:
    """Idempotent client row for wizard save/calculate tests (CTR-006)."""
    import hashlib

    from app.repositories.counterparties_repository import CounterpartiesRepository

    repo = CounterpartiesRepository(db_path=db_path)
    code = (
        "TEST-00"
        if name == "ООО Тест"
        else "TEST-" + hashlib.md5(name.encode("utf-8")).hexdigest()[:10]
    )
    existing = repo.get_by_code_1c(code)
    if existing is not None:
        return int(existing["id"])
    created = repo.insert(
        code_1c=code,
        name=name,
        inn="7701000001",
        kpp="770101001",
        is_client=True,
        source="manual",
    )
    return int(created["id"])


def client_meta_payload(client_name: str, **extra: Any) -> dict[str, Any]:
    """Meta fields for create-path calculate: catalog pick + display name."""
    from app.core.settings import get_settings

    cp_id = seed_test_counterparty(str(get_settings().plita_db_path), name=client_name)
    return {"client_name": client_name, "counterparty_id": cp_id, **extra}


def bind_draft_counterparty(client: Any, draft_id: str, db_path: str | None = None) -> int:
    from app.core.settings import get_settings

    path = db_path or str(get_settings().plita_db_path)
    cp_id = seed_test_counterparty(path)
    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={"counterparty_id": cp_id},
    )
    assert response.status_code == 200, response.text
    return cp_id

