from __future__ import annotations

import sqlite3
from pathlib import Path

from core import commercial_offer, commercial_offer_xlsx
from core.price_db import length_m_to_price_length_dm


def _create_price_db(path: Path) -> None:
    conn = sqlite3.connect(path)
    try:
        cur = conn.cursor()
        cur.execute(
            "CREATE TABLE prices (length_dm INTEGER, load_code INTEGER, price REAL, PRIMARY KEY(length_dm, load_code))"
        )
        cur.execute(
            "INSERT INTO prices (length_dm, load_code, price) VALUES (?, ?, ?)",
            (28, 8, 12345.0),
        )
        conn.commit()
    finally:
        conn.close()


def test_length_m_to_price_length_dm_uses_ceil() -> None:
    assert length_m_to_price_length_dm(2.73) == 28
    assert length_m_to_price_length_dm(2.7) == 27
    assert length_m_to_price_length_dm(5.5) == 55
    assert length_m_to_price_length_dm(2.7000000000000002) == 27


def test_pdf_fallback_price_uses_new_length_key(tmp_path: Path, monkeypatch) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)
    monkeypatch.setattr(commercial_offer, "DB_PATH", str(db_path))

    price = commercial_offer.get_plate_price(2.73, 1.2, 800)

    assert price == 12345.0


def test_xlsx_fallback_price_uses_new_length_key(tmp_path: Path, monkeypatch) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)
    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", str(db_path))

    price = commercial_offer_xlsx.get_plate_price(2.73, 1.2, 800)

    assert price == 12345.0
