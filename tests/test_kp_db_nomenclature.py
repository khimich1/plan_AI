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


@pytest.mark.parametrize(
    "rows",
    [
        [
            ("bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb", "Плиты ПБ 63-12-8п"),
            ("aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa", "Плиты ПБ 63-12-8п"),
        ],
        [
            ("aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa", "Плиты ПБ 63-12-8п"),
            ("bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb", "Плиты ПБ 63-12-8п"),
        ],
    ],
)
def test_lookup_duplicate_names_returns_stable_guid(
    tmp_path, rows: list[tuple[str, str]]
) -> None:
    """Same plate name, two 1C GUIDs: always the lexicographically smallest GUID."""
    db_path = str(tmp_path / "pb.db")
    _make_prays_db(db_path, rows)

    with sqlite3.connect(db_path) as conn:
        name, nom_id, match_type = nom.lookup_nomenclature_by_plate_name(
            "Плиты ПБ 63-12-8п", conn.cursor()
        )

    assert name == "Плиты ПБ 63-12-8п"
    assert nom_id == "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
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


def test_lookup_trailing_decimal_zero_matches_catalog_mark(tmp_path) -> None:
    """75,10 в заказе — та же плита, что 75,1 в prays_plity."""
    db_path = str(tmp_path / "pb.db")
    canonical = "Плиты ПБ 75,1-12-8п"
    _make_prays_db(db_path, [("NOM-751", canonical)])

    with sqlite3.connect(db_path) as conn:
        name, nom_id, match_type = nom.lookup_nomenclature_by_plate_name(
            "Плиты ПБ 75,10-12-8п", conn.cursor()
        )

    assert name == canonical
    assert nom_id == "NOM-751"
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


# --- NG-101: find_plate_card — порядок D-load12, LIKE сам не выбирает ------------

GUID_125 = "0de0f4a7-dc8e-11ea-8f69-04d9f5387845"
GUID_PV71 = "ad13b504-96c5-11eb-8f7e-04d9f5387845"
GUID_PV72 = "b3596b8f-96c5-11eb-8f7e-04d9f5387845"


def _find(rows: list[tuple[str, str]], tmp_path, name: str) -> "nom.PlateCardMatch":
    db_path = str(tmp_path / "pb.db")
    _make_prays_db(db_path, rows)
    with sqlite3.connect(db_path) as conn:
        return nom.find_plate_card(conn.cursor(), name)


def _cards_51_12() -> list[tuple[str, str]]:
    return [
        (GUID_PV71, "Плиты ПБ 51-12-12п Пв-7.1"),
        (GUID_PV72, "Плиты ПБ 51-12-12п Пв-7.2"),
        (GUID_125, "Плиты ПБ 51-12-12,5п"),
    ]


def test_find_plate_card_alias_beats_like_on_special_cards(tmp_path) -> None:
    match = _find(_cards_51_12(), tmp_path, "Плиты ПБ 51-12-12п")

    assert match.status == "found"
    assert match.canonical_name == "Плиты ПБ 51-12-12,5п"
    assert match.guid == GUID_125
    assert match.guid_count == 1


def test_lookup_and_count_follow_the_same_alias_rule(tmp_path) -> None:
    db_path = str(tmp_path / "pb.db")
    _make_prays_db(db_path, _cards_51_12())

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        name, nom_id, match_type = nom.lookup_nomenclature_by_plate_name(
            "Плиты ПБ 51-12-12п", cur
        )
        count = nom.count_nomenclature_by_plate_name("Плиты ПБ 51-12-12п", cur)

    assert (name, nom_id) == ("Плиты ПБ 51-12-12,5п", GUID_125)
    assert match_type == "load12_alias"
    assert count == 1


def test_find_plate_card_own_12_card_wins_over_12_5(tmp_path) -> None:
    rows = _cards_51_12() + [("cccccccc-0000-0000-0000-000000000000", "Плиты ПБ 51-12-12п")]

    match = _find(rows, tmp_path, "Плиты ПБ 51-12-12п")

    assert match.canonical_name == "Плиты ПБ 51-12-12п"
    assert match.guid == "cccccccc-0000-0000-0000-000000000000"


def test_find_plate_card_alias_with_two_guids_is_dup(tmp_path) -> None:
    rows = [("bbbbbbbb-0000-0000-0000-000000000000", "Плиты ПБ 33-12-12,5п"),
            ("aaaaaaaa-0000-0000-0000-000000000000", "Плиты ПБ 33-12-12,5п")]

    match = _find(rows, tmp_path, "Плиты ПБ 33-12-12п")

    assert match.status == "dup"
    assert match.canonical_name == "Плиты ПБ 33-12-12,5п"
    assert match.guid_count == 2
    assert match.guid == "aaaaaaaa-0000-0000-0000-000000000000"  # стабильный первый для lookup


def test_find_plate_card_dup_reports_card_name_not_typed_name(tmp_path) -> None:
    rows = [("aaaaaaaa-0000-0000-0000-000000000000", "Плиты ПБ 22,8-12-8п"),
            ("bbbbbbbb-0000-0000-0000-000000000000", "Плиты ПБ 22,8-12-8п")]

    match = _find(rows, tmp_path, "Плиты ПБ 22,80-12-8п")

    assert match.status == "dup"
    assert match.canonical_name == "Плиты ПБ 22,8-12-8п"


def test_find_plate_card_alias_keeps_width_variant(tmp_path) -> None:
    match = _find([(GUID_125, "Плиты ПБ 55-7,0-12,5п")], tmp_path, "Плиты ПБ 55-7-12п")

    assert match.status == "found"
    assert match.guid == GUID_125


def test_find_plate_card_is_one_sided_12_5_does_not_fall_back_to_12(tmp_path) -> None:
    rows = [(GUID_PV71, "Плиты ПБ 51-12-12п Пв-7.1")]

    match = _find(rows, tmp_path, "Плиты ПБ 51-12-12,5п")

    assert match.status == "none"


@pytest.mark.parametrize("name", ["Плиты ПБ 51-12-8п", "Плиты ПБ 51-12-10п"])
def test_find_plate_card_other_loads_are_not_aliased(tmp_path, name: str) -> None:
    match = _find([(GUID_125, "Плиты ПБ 51-12-12,5п")], tmp_path, name)

    assert match.status == "none"


def test_find_plate_card_like_with_several_candidates_does_not_choose(tmp_path) -> None:
    rows = [(GUID_PV72, "Плиты ПБ 66,1-12-12п с выбр Б"), (GUID_PV71, "Плиты ПБ 66,1-12-12п с выбр А")]
    db_path = str(tmp_path / "pb.db")
    _make_prays_db(db_path, rows)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        match = nom.find_plate_card(cur, "Плиты ПБ 66,1-12-12п")
        looked_up = nom.lookup_nomenclature_by_plate_name("Плиты ПБ 66,1-12-12п", cur)

    assert match.status == "variants"
    assert match.guid is None
    assert match.candidates == (
        "Плиты ПБ 66,1-12-12п с выбр А",
        "Плиты ПБ 66,1-12-12п с выбр Б",
    )
    assert looked_up == (None, None, None)
    assert nom.count_nomenclature_by_plate_name("Плиты ПБ 66,1-12-12п", sqlite3.connect(db_path).cursor()) == 0


def test_find_plate_card_like_checks_length_of_every_candidate(tmp_path) -> None:
    rows = [
        ("aaaaaaaa-0000-0000-0000-000000000000", "Плиты ПБ 40,1-12-8п и 40,8-12-8п"),
        ("bbbbbbbb-0000-0000-0000-000000000000", "Плиты ПБ 40,8-12-8п Пв-1"),
    ]

    match = _find(rows, tmp_path, "Плиты ПБ 40,8-12-8п")

    assert match.status == "found"
    assert match.canonical_name == "Плиты ПБ 40,8-12-8п Пв-1"
    assert match.match_type == "like"


def test_find_plate_card_like_same_name_two_guids_is_dup(tmp_path) -> None:
    rows = [("aaaaaaaa-0000-0000-0000-000000000000", "Плиты ПБ 59,81-12-8п"),
            ("bbbbbbbb-0000-0000-0000-000000000000", "Плиты ПБ 59,81-12-8п")]

    match = _find(rows, tmp_path, "ПБ 59,81-12-8п")

    assert match.status == "dup"
    assert match.guid_count == 2


@pytest.mark.parametrize(
    ("name", "alias"),
    [
        ("Плиты ПБ 71-12-12п", "Плиты ПБ 71-12-12,5п"),
        ("Плиты ПБ 71-12-12.5п", "Плиты ПБ 71-12-12,5п"),
        ("Плиты ПБ 71-12-12,50п", "Плиты ПБ 71-12-12,5п"),
        ("Плиты ПБ 71-12-12,5п", None),
        ("Плиты ПБ 71-12-8п", None),
        ("С30.30-12", None),
    ],
)
def test_plate_load_12_alias(name: str, alias: str | None) -> None:
    assert nom.plate_load_12_alias(name) == alias


# --- NG-102: имя КП = имя карточки, load_class из имени (12,5 → 1250) ------------


@pytest.mark.parametrize(
    ("name", "expected"),
    [
        ("Плиты ПБ 51-12-12,5п", 1250),
        ("Плиты ПБ 51-12-12п", 1200),
        ("Плиты ПБ 51-12-8п", 800),
        ("Плиты ПБ 51-12-12.5п", 1250),
        ("Плиты ПБ 51-12-12,50п", 1250),
        ("Плиты ПБ 51-12-10п", 1000),
        ("Плиты ПБ 51-12-12п Пв-7.1", None),
        ("Плиты ПБ 51-12", None),
        ("", None),
    ],
)
def test_load_class_from_plate_name(name: str, expected: int | None) -> None:
    assert nom.load_class_from_plate_name(name) == expected


def test_enrich_rewrites_name_and_load_class_from_alias_card(monkeypatch, tmp_path) -> None:
    db_path = str(tmp_path / "pb.db")
    _make_prays_db(db_path, _cards_51_12())
    monkeypatch.setattr(nom, "_PB_DB_PATH", db_path)

    def _forbidden(*_args, **_kwargs):
        raise AssertionError("parse_load_code_from_name даёт 13 для 12,5 — нельзя")

    monkeypatch.setattr("core.config_and_data.parse_load_code_from_name", _forbidden)
    order = [{"name": "Плиты ПБ 51-12-12п", "qty": 1, "load_class": 1200}]

    nom.enrich_order_data_with_nomenclature(order)

    assert order[0]["name"] == "Плиты ПБ 51-12-12,5п"
    assert order[0]["load_class"] == 1250
    assert order[0]["nomenclature_id"] == GUID_125
    assert "Пв-7.1" not in order[0]["name"]


def test_enrich_keeps_load_class_when_name_is_not_rewritten(monkeypatch, tmp_path) -> None:
    db_path = str(tmp_path / "pb.db")
    _make_prays_db(db_path, [("NOM-63", "Плиты ПБ 63-12-8п")])
    monkeypatch.setattr(nom, "_PB_DB_PATH", db_path)
    order = [{"name": "Плиты ПБ 63-12-8п", "qty": 2, "load_class": 800}]

    nom.enrich_order_data_with_nomenclature(order)

    assert order[0]["load_class"] == 800


def test_enrich_does_not_pick_one_of_several_special_cards(monkeypatch, tmp_path) -> None:
    db_path = str(tmp_path / "pb.db")
    _make_prays_db(db_path, _cards_51_12()[:2])
    monkeypatch.setattr(nom, "_PB_DB_PATH", db_path)
    order = [{"name": "Плиты ПБ 51-12-12п", "qty": 1, "load_class": 1200}]

    nom.enrich_order_data_with_nomenclature(order)

    assert order[0]["name"] == "Плиты ПБ 51-12-12п"
    assert order[0]["load_class"] == 1200
    assert order[0]["nomenclature_id"] is None
