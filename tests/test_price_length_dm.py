from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest
from bot.handlers import commercial as commercial_handler
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


@pytest.mark.skip(
    reason="Устаревший контракт bot.handlers.commercial: нет _price_row_matches_item (OPT-010).",
)
def test_commercial_row_match_uses_normalized_length_dm() -> None:
    row = [1, "Плиты ПБ 68-12-8п", 1, "шт", "—", "—", "1000", "27 300,00", "27 300,00"]

    assert commercial_handler._price_row_matches_item(
        row,
        length_m=6.75,
        width_m=1.2,
        load_code=8,
    )
    assert not commercial_handler._price_row_matches_item(
        row,
        length_m=6.7,
        width_m=1.2,
        load_code=8,
    )


@pytest.mark.skip(
    reason="Устаревший контракт bot.handlers.commercial: нет _price_row_matches_item (OPT-010).",
)
def test_commercial_row_match_handles_12_5_load() -> None:
    row = [1, "Плиты ПБ 67-12-12,5п", 1, "шт", "—", "—", "1000", "27 300,00", "27 300,00"]

    assert commercial_handler._price_row_matches_item(
        row,
        length_m=6.7,
        width_m=1.2,
        load_code=12.5,
    )


@pytest.mark.skip(
    reason="Устаревший контракт bot.handlers.commercial: нет get_basic_plate_price (OPT-010).",
)
def test_commercial_unit_price_uses_fallback_instead_of_zero(monkeypatch) -> None:
    monkeypatch.setattr(commercial_handler, "get_basic_plate_price", lambda *args, **kwargs: 54321.0)

    assert commercial_handler._extract_unit_price_from_row_or_fallback(
        None,
        length_m=6.75,
        width_m=1.2,
        load_code=8,
    ) == 54321.0
