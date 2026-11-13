"""NG-104…NG-106: «почему нет GUID» — та же карточка, срез своего вида, поиск."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.guid_explain import (
    SEARCH_MIN_CHARS,
    SLICE_LIMIT,
    GuidCandidate,
    explain_missing_guid,
    search_guid_marks,
)
from core.guid_gate import REASON_MISSING, OrderLine
from core.nomenclature_guid import ensure_schema as ensure_guid_schema
from core.nomenclature_guid import upsert
from core.plate_guid_choice import ensure_schema as ensure_choice_schema

G1 = "11111111-1111-1111-1111-111111111111"
G2 = "22222222-2222-2222-2222-222222222222"
G3 = "33333333-3333-3333-3333-333333333333"

_PRICE_TABLES = {
    "pile": "pile_prices",
    "bridge_pile": "bridge_pile_prices",
    "composite_pile": "composite_pile_prices",
    "fbs": "fbs_prices",
    "stair_flight": "march_prices",
    "stair_step": "step_prices",
}


@pytest.fixture
def conn(tmp_path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(tmp_path / "pb.db")
    ensure_guid_schema(connection)
    ensure_choice_schema(connection)
    connection.execute(
        'CREATE TABLE prays_plity ("Уникальный идентификатор (Номенклатура)" TEXT, "Товар" TEXT)'
    )
    try:
        yield connection
    finally:
        connection.close()


def _cards(conn: sqlite3.Connection, *rows: tuple[str, str]) -> None:
    conn.executemany("INSERT INTO prays_plity VALUES (?, ?)", rows)


def _price_mark(conn: sqlite3.Connection, kind: str, mark: str) -> None:
    table = _PRICE_TABLES[kind]
    conn.execute(f"CREATE TABLE IF NOT EXISTS {table} (mark TEXT)")
    conn.execute(f"INSERT INTO {table} (mark) VALUES (?)", (mark,))


def _by_mark(items) -> dict[str, GuidCandidate]:
    return {item.mark: item for item in items}


def _explain(conn: sqlite3.Connection, kind: str, mark: str):
    return explain_missing_guid(OrderLine(kind, mark), conn)


# ------------------------------------------------------------------ плиты (NG-104)


def test_plate_same_card_trailing_zero_one_number(conn) -> None:
    _cards(conn, (G1, "Плиты ПБ 22,8-12-8п"))

    result = _explain(conn, "plate", "Плиты ПБ 22,80-12-8п")

    assert result.same_card == GuidCandidate(
        mark="Плиты ПБ 22,8-12-8п", guid_count=1, label="лишний ноль"
    )


def test_plate_same_card_two_numbers_is_dup_and_not_manager_selectable(conn) -> None:
    _cards(conn, (G1, "Плиты ПБ 22,8-12-8п"), (G2, "Плиты ПБ 22,8-12-8п"))

    result = _explain(conn, "plate", "Плиты ПБ 22,80-12-8п")

    assert result.same_card is not None and result.same_card.guid_count == 2
    assert result.same_card.selectable_for("manager") is False
    assert result.same_card.selectable_for("economist") is True


def test_plate_two_guids_of_one_name_are_two_cards_with_that_name(conn) -> None:
    _cards(conn, (G2, "Плиты ПБ 22,8-12-8п"), (G1, "Плиты ПБ 22,8-12-8п"))

    result = _explain(conn, "plate", "Плиты ПБ 22,80-12-8п")

    assert result.same_card is not None
    assert [(card.name, card.guid) for card in result.same_card.cards] == [
        ("Плиты ПБ 22,8-12-8п", G1),
        ("Плиты ПБ 22,8-12-8п", G2),
    ]


def test_one_plate_guid_is_not_a_card_choice(conn) -> None:
    _cards(conn, (G1, "Плиты ПБ 22,8-12-8п"))

    result = _explain(conn, "plate", "Плиты ПБ 22,80-12-8п")

    assert result.same_card is not None
    assert result.same_card.guid_count == 1
    assert result.same_card.cards == ()


def test_plate_slice_with_two_guids_returns_both_cards(conn) -> None:
    _cards(conn, (G2, "Плиты ПБ 29-5,3-8п"), (G1, "Плиты ПБ 29-5,3-8п"))

    result = _explain(conn, "plate", "Плиты ПБ 29-5,25-8п")

    assert len(result.slice) == 1
    assert result.slice[0].guid_count == 2
    assert [(card.name, card.guid) for card in result.slice[0].cards] == [
        ("Плиты ПБ 29-5,3-8п", G1),
        ("Плиты ПБ 29-5,3-8п", G2),
    ]


def test_plate_nearby_width_is_in_slice_and_nothing_is_substituted(conn) -> None:
    _cards(conn, (G1, "Плиты ПБ 29-5,3-8п"), (G2, "Плиты ПБ 40-12-8п"))

    result = _explain(conn, "plate", "Плиты ПБ 29-5,25-8п")

    assert result.same_card is None
    assert result.mark == "Плиты ПБ 29-5,25-8п"
    assert [item.mark for item in result.slice] == ["Плиты ПБ 29-5,3-8п"]
    assert result.slice[0].label == "другая ширина"
    assert result.slice[0].selectable_for("manager") is True


def test_plate_slice_shows_widths_of_same_length_and_no_missing_width(conn) -> None:
    _cards(
        conn,
        (G1, "Плиты ПБ 49,5-12-8п"),
        (G2, "Плиты ПБ 49,5-10,2-8п"),
        (G3, "Плиты ПБ 50-15-8п"),
    )

    result = _explain(conn, "plate", "Плиты ПБ 49,5-15-8п")

    assert result.same_card is None
    assert {item.mark for item in result.slice} == {
        "Плиты ПБ 49,5-12-8п",
        "Плиты ПБ 49,5-10,2-8п",
    }
    assert all("-15-" not in item.mark for item in result.slice)


def test_plate_special_cards_are_listed_each_with_own_guid_count(conn) -> None:
    _cards(
        conn,
        (G1, "Плиты ПБ 66,1-12-12п с выбр А"),
        (G2, "Плиты ПБ 66,1-12-12п с выбр Б"),
        (G3, "Плиты ПБ 66,1-12-12п с выбр Б"),
        ("44444444-4444-4444-4444-444444444444", "Плиты ПБ 66,1-12-12п с выбр эскиз 8"),
        ("55555555-5555-5555-5555-555555555555", "Плиты ПБ 66,1-12-12п с выбр эскиз 28"),
    )

    result = _explain(conn, "plate", "Плиты ПБ 66,1-12-12п")

    assert result.reason != ""
    assert len(result.slice) == 4
    counts = {item.mark: item.guid_count for item in result.slice}
    assert counts["Плиты ПБ 66,1-12-12п с выбр Б"] == 2
    assert counts["Плиты ПБ 66,1-12-12п с выбр А"] == 1
    assert all(item.label is None for item in result.slice)  # особые: подписи нет
    assert result.same_card is None


def test_plate_label_only_when_one_difference(conn) -> None:
    _cards(
        conn,
        (G1, "Плиты ПБ 29-5,3-8п"),
        (G2, "Плиты ПБ 29-6-10п"),
        (G3, "Плиты ПБ 29-5,25-10п"),
    )

    result = _explain(conn, "plate", "Плиты ПБ 29-5,25-8п")

    labels = {item.mark: item.label for item in result.slice}
    assert labels["Плиты ПБ 29-5,3-8п"] == "другая ширина"
    assert labels["Плиты ПБ 29-5,25-10п"] == "другая нагрузка"
    assert labels["Плиты ПБ 29-6-10п"] is None  # ширина и нагрузка


def test_plate_slice_is_capped_and_nearest_first(conn) -> None:
    for index in range(40):
        width = f"{index + 20}"
        _cards(conn, (f"{index:08d}-0000-0000-0000-000000000000", f"Плиты ПБ 30-{width}-8п"))

    result = _explain(conn, "plate", "Плиты ПБ 30-21,5-8п")

    assert len(result.slice) == SLICE_LIMIT == 30
    assert result.slice[0].mark in {"Плиты ПБ 30-21-8п", "Плиты ПБ 30-22-8п"}


def test_plate_alias_same_card_for_load_12(conn) -> None:
    _cards(
        conn,
        (G1, "Плиты ПБ 51-12-12п Пв-7.1"),
        (G2, "Плиты ПБ 51-12-12,5п"),
    )

    result = _explain(conn, "plate", "Плиты ПБ 51-12-12п")

    assert result.same_card is not None
    assert result.same_card.mark == "Плиты ПБ 51-12-12,5п"
    assert result.same_card.guid_count == 1
    assert "Плиты ПБ 51-12-12п Пв-7.1" in {item.mark for item in result.slice}


def test_plate_reason_comes_from_gate_and_unparsed_name_has_no_slice(conn) -> None:
    result = _explain(conn, "plate", "что-то странное")

    assert result.reason == REASON_MISSING
    assert result.slice == ()
    assert result.same_card is None


def test_ready_line_has_empty_reason(conn) -> None:
    _cards(conn, (G1, "Плиты ПБ 22,8-12-8п"))

    result = _explain(conn, "plate", "Плиты ПБ 22,8-12-8п")

    assert result.reason == ""


# --------------------------------------------- сваи, мостовые, составные (NG-105)


def test_bridge_pile_priced_but_without_number_is_visible_not_manager_selectable(conn) -> None:
    _price_mark(conn, "bridge_pile", "С9-35T7")
    upsert(conn, "bridge_pile", "С9-35T7", match_status="missing")

    result = _explain(conn, "bridge_pile", "C9-35T7")

    assert result.same_card == GuidCandidate(
        mark="С9-35T7", guid_count=0, label="другая буква"
    )
    assert result.same_card.selectable_for("manager") is False
    assert result.same_card.selectable_for("economist") is True


def test_bridge_pile_other_mark_is_in_slice_not_same_card(conn) -> None:
    _price_mark(conn, "bridge_pile", "С9-35T7")
    upsert(conn, "bridge_pile", "С9-35T6", guid_1c=G1, match_status="auto")
    upsert(conn, "bridge_pile", "С9-40T7", guid_1c=G2, match_status="auto")

    result = _explain(conn, "bridge_pile", "C9-35T7")

    marks = {item.mark: item for item in result.slice}
    assert "С9-35T6" in marks and marks["С9-35T6"].guid_count == 1
    assert "С9-40T7" not in marks  # другое сечение
    assert result.same_card is not None and result.same_card.mark == "С9-35T7"


def test_bridge_pile_two_cards_is_two_numbers(conn) -> None:
    upsert(conn, "bridge_pile", "С15-35T7", match_status="ambiguous")

    result = _explain(conn, "bridge_pile", "C15-35T7")

    assert result.same_card is not None
    assert result.same_card.guid_count == 2
    assert result.same_card.selectable_for("manager") is False
    assert result.same_card.cards == ()


def test_ambiguous_mark_returns_card_names_from_1c(conn) -> None:
    from core.duplicate_candidates import upsert_candidates

    upsert(conn, "bridge_pile", "С15-35T7", match_status="ambiguous")
    upsert_candidates(
        conn,
        "bridge_pile",
        "С15-35T7",
        [
            (G1, "Сваи мостовые С15-35T7 завод", None),
            (G2, "Сваи мостовые С15-35T7 особая", None),
        ],
    )

    result = _explain(conn, "bridge_pile", "C15-35T7")

    assert result.same_card is not None
    assert [(card.name, card.guid) for card in result.same_card.cards] == [
        ("Сваи мостовые С15-35T7 завод", G1),
        ("Сваи мостовые С15-35T7 особая", G2),
    ]


def test_pile_u_row_offers_cards_of_the_base_mark(conn) -> None:
    from core.duplicate_candidates import upsert_candidates

    upsert(conn, "pile", "С70.35-9", match_status="ambiguous")
    upsert_candidates(
        conn,
        "pile",
        "С70.35-9",
        [(G1, "Сваи С 70.35-9", None), (G2, "Сваи С 70.35-9 завод", None)],
    )

    result = _explain(conn, "pile", "С70.35-9у")
    rows = _by_mark(result.slice + result.same_cards)

    assert [card.guid for card in rows["С70.35-9"].cards] == [G1, G2]
    assert [card.guid for card in rows["С70.35-9у"].cards] == [G1, G2]
    assert rows["С70.35-9"].cards[0].name == "Сваи С 70.35-9"
    assert rows["С70.35-9"].cards[1].name == "Сваи С 70.35-9 завод"


def test_pile_u_variants_are_two_rows_and_u_counts_only_guid_1c_u(conn) -> None:
    upsert(conn, "pile", "С110.30-9", guid_1c=None, guid_1c_u=G1, match_status="auto")

    result = _explain(conn, "pile", "С110.30-9")
    rows = _by_mark(result.slice + result.same_cards)

    assert rows["С110.30-9"].guid_count == 0
    assert rows["С110.30-9у"].guid_count == 1
    assert rows["С110.30-9у"].selectable_for("manager") is True
    assert rows["С110.30-9"].selectable_for("manager") is False


def test_pile_slice_is_same_section(conn) -> None:
    upsert(conn, "pile", "С100.30-10", guid_1c=G1, match_status="auto")
    upsert(conn, "pile", "С100.35-10", guid_1c=G2, match_status="auto")
    _price_mark(conn, "pile", "С90.30-8")

    result = _explain(conn, "pile", "С110.30-9")

    marks = {item.mark for item in result.slice}
    assert {"С100.30-10", "С90.30-8"} <= marks
    assert "С100.35-10" not in marks


def test_pile_search_universe_includes_price_only_marks(conn) -> None:
    _price_mark(conn, "pile", "С70.30-5")

    result = _explain(conn, "pile", "С70.30-5у")

    rows = _by_mark(result.slice + result.same_cards)
    assert rows["С70.30-5"].guid_count == 0


# ----------------------------------------------------- ФБС, ступени, марши (NG-106)


def test_fbs_same_card_by_lookup_rule_and_slice_by_width_height(conn) -> None:
    upsert(conn, "fbs", "ФБС 24.4.6-T", guid_1c=G1, match_status="auto")
    upsert(conn, "fbs", "ФБС 12.4.6-Т", guid_1c=G2, match_status="auto")
    upsert(conn, "fbs", "ФБС 12.5.3-Т", guid_1c=G3, match_status="auto")

    result = _explain(conn, "fbs", "фбс 24.4.6-т")

    assert result.same_card is not None and result.same_card.mark == "ФБС 24.4.6-T"
    assert {item.mark for item in result.slice} == {"ФБС 12.4.6-Т"}


def test_steps_slice_is_own_series_and_no_same_card_for_shorter_mark(conn) -> None:
    upsert(conn, "stair_step", "ЛС11-1", guid_1c=G1, match_status="auto")
    upsert(conn, "stair_step", "ЛС11-2", guid_1c=G2, match_status="auto")
    upsert(conn, "stair_step", "ЛС14-1", guid_1c=G3, match_status="auto")

    result = _explain(conn, "stair_step", "ЛС11")

    assert result.same_card is None
    assert {item.mark for item in result.slice} == {"ЛС11-1", "ЛС11-2"}


def test_marches_slice_is_own_series(conn) -> None:
    upsert(conn, "stair_flight", "1ЛМ 30-11-15-4", guid_1c=G1, match_status="auto")
    upsert(conn, "stair_flight", "1ЛМ 30-12-15-4", guid_1c=G2, match_status="auto")
    upsert(conn, "stair_flight", "1ЛМ 27-11-14-4", guid_1c=G3, match_status="auto")

    result = _explain(conn, "stair_flight", "1ЛМ 30-13-15-4")

    assert {item.mark for item in result.slice} == {"1ЛМ 30-11-15-4", "1ЛМ 30-12-15-4"}


def test_unparsed_non_plate_mark_has_no_slice(conn) -> None:
    upsert(conn, "fbs", "ФБС 12.4.6-Т", guid_1c=G1, match_status="auto")

    result = _explain(conn, "fbs", "что-то")

    assert result.slice == ()
    assert result.same_card is None


# ------------------------------------------------------------------------- поиск


def test_search_requires_three_chars(conn) -> None:
    _cards(conn, (G1, "Плиты ПБ 22,8-12-8п"))

    assert SEARCH_MIN_CHARS == 3
    assert search_guid_marks(conn, "plate", "ПБ") == ()


def test_search_plates_is_case_insensitive_and_capped(conn) -> None:
    for index in range(35):
        _cards(conn, (f"{index:08d}-0000-0000-0000-000000000000", f"Плиты ПБ 30-{index + 20}-8п"))
    _cards(conn, (G1, "Плиты ПБ 22,8-12-8п"))

    found = search_guid_marks(conn, "plate", "плиты пб 30")

    assert len(found) == 30
    assert all(item.mark.startswith("Плиты ПБ 30-") for item in found)


def test_search_stays_in_own_kind_and_shows_zero_number_rows(conn) -> None:
    _price_mark(conn, "bridge_pile", "С9-35T7")
    upsert(conn, "bridge_pile", "С9-35T6", guid_1c=G1, match_status="auto")
    upsert(conn, "fbs", "ФБС 12.4.6-Т", guid_1c=G2, match_status="auto")

    found = _by_mark(search_guid_marks(conn, "bridge_pile", "С9-35"))

    assert set(found) == {"С9-35T6", "С9-35T7"}
    assert found["С9-35T7"].guid_count == 0
    assert found["С9-35T6"].guid_count == 1


def test_search_finds_latin_query_for_cyrillic_mark(conn) -> None:
    _price_mark(conn, "bridge_pile", "С9-35T7")

    found = search_guid_marks(conn, "bridge_pile", "C9-35T")

    assert [item.mark for item in found] == ["С9-35T7"]


def test_search_returns_both_cards_of_a_duplicate_plate_name(conn) -> None:
    _cards(conn, (G2, "Плиты ПБ 22,8-10,2-8п"), (G1, "Плиты ПБ 22,8-10,2-8п"))

    found = search_guid_marks(conn, "plate", "10,2")

    assert len(found) == 1
    assert [(card.name, card.guid) for card in found[0].cards] == [
        ("Плиты ПБ 22,8-10,2-8п", G1),
        ("Плиты ПБ 22,8-10,2-8п", G2),
    ]
