"""Unit tests for core.kp_db_nomenclature lookup and enrich logic."""

from __future__ import annotations

import sqlite3

import pytest

import core.kp_db_nomenclature as nom


def _make_prays_db(path: str, rows: list[tuple[str, str]]) -> None:
    """Create pb.db stub with prays_plity (nomenclature_id, product_name)."""
    with sqlite3.connect(path) as conn:
        conn.execute(
            'CREATE TABLE prays_plity ('
            '"Уникальный идентификатор (Номенклатура)" TEXT, '
            '"Товар" TEXT)'
        )
        conn.executemany(
            'INSERT INTO prays_plity VALUES (?, ?)',
            rows,
        )
        conn.commit()


@pytest.mark.parametrize(
    ("name", "expected"),
    [
        ("Плиты ПБ 63-12-8п", 63.0),
        ("ПБ 59,81-12-8п", 59.81),
        ("Плиты ПБ 40,0-12-8п", 40.0),
        ("not a plate", None),
        ("", None),
    ],
)
def test_extract_length_dm_val(name: str, expected: float | None) -> None:
    assert nom._extract_length_dm_val(name) == expected


def test_lookup_exact_match(tmp_path) -> None:
    db_path = str(tmp_path / "pb.db")
    canonical = "Плиты ПБ 63-12-8п"
    _make_prays_db(db_path, [("NOM-1", canonical)])

    with sqlite3.connect(db_path) as conn:
        name, nom_id, match_type = nom.lookup_nomenclature_by_plate_name(
            canonical, conn.cursor()
        )

    assert name == canonical
    assert nom_id == "NOM-1"
    assert match_type == "exact"


def test_lookup_prays_width_variant(tmp_path) -> None:
    """Bot-style 0.3 m ribbon maps to prays '-3,0-' variant."""
    db_path = str(tmp_path / "pb.db")
    prays_name = "Плиты ПБ 42-3,0-8п"
    bot_name = "Плиты ПБ 42-0.3-8п"
    _make_prays_db(db_path, [("NOM-RIB", prays_name)])

    with sqlite3.connect(db_path) as conn:
        name, nom_id, match_type = nom.lookup_nomenclature_by_plate_name(
            bot_name, conn.cursor()
        )

    assert name == prays_name
    assert nom_id == "NOM-RIB"
    assert match_type == "exact_prays_variant"


def test_lookup_prays_length_variant_bridges_40_and_40_0(tmp_path) -> None:
    """Целая длина 40 в марке сопоставляется со справочником 40,0 через prays variant."""
    db_path = str(tmp_path / "pb.db")
    canonical = "Плиты ПБ 40,0-12-8п"
    _make_prays_db(db_path, [("NOM-40", canonical)])

    with sqlite3.connect(db_path) as conn:
        name, nom_id, match_type = nom.lookup_nomenclature_by_plate_name(
            "Плиты ПБ 40-12-8п", conn.cursor()
        )

    assert name == canonical
    assert nom_id == "NOM-40"
    assert match_type == "exact_prays_variant"


def test_lookup_like_rejects_length_mismatch(tmp_path) -> None:
    db_path = str(tmp_path / "pb.db")
    _make_prays_db(
        db_path,
        [("NOM-MIX", "Плиты ПБ 40,1-12-8п и 40,8-12-8п")],
    )

    with sqlite3.connect(db_path) as conn:
        name, nom_id, match_type = nom.lookup_nomenclature_by_plate_name(
            "Плиты ПБ 40,8-12-8п", conn.cursor()
        )

    assert name is None
    assert nom_id is None
    assert match_type is None


def test_lookup_like_accepts_matching_fractional_length(tmp_path) -> None:
    db_path = str(tmp_path / "pb.db")
    canonical = "Плиты ПБ 59,81-12-8п"
    _make_prays_db(db_path, [("NOM-5981", canonical)])

    with sqlite3.connect(db_path) as conn:
        name, nom_id, match_type = nom.lookup_nomenclature_by_plate_name(
            "ПБ 59,81-12-8п", conn.cursor()
        )

    assert name == canonical
    assert nom_id == "NOM-5981"
    assert match_type == "like"


def test_enrich_order_data_fills_nomenclature(monkeypatch, tmp_path) -> None:
    db_path = str(tmp_path / "pb.db")
    canonical = "Плиты ПБ 63-12-8п"
    _make_prays_db(db_path, [("NOM-63", canonical)])
    monkeypatch.setattr(nom, "_PB_DB_PATH", db_path)

    order = [{"name": canonical, "qty": 2}]
    result = nom.enrich_order_data_with_nomenclature(order)

    assert result is order
    assert order[0]["name"] == canonical
    assert order[0]["nomenclature_id"] == "NOM-63"


def test_enrich_order_data_skips_existing_nomenclature_id(monkeypatch, tmp_path) -> None:
    db_path = str(tmp_path / "pb.db")
    _make_prays_db(db_path, [("NOM-OTHER", "Плиты ПБ 99-12-8п")])
    monkeypatch.setattr(nom, "_PB_DB_PATH", db_path)

    order = [{"name": "Плиты ПБ 99-12-8п", "nomenclature_id": 42, "qty": 1}]
    nom.enrich_order_data_with_nomenclature(order)

    assert order[0]["nomenclature_id"] == 42


def test_enrich_order_data_missing_db_returns_unchanged(monkeypatch, tmp_path) -> None:
    missing = str(tmp_path / "no_pb.db")
    monkeypatch.setattr(nom, "_PB_DB_PATH", missing)

    order = [{"name": "Плиты ПБ 1-12-8п", "qty": 1}]
    result = nom.enrich_order_data_with_nomenclature(order)

    assert result is order
    assert "nomenclature_id" not in order[0]
