"""GPS-003: bootstrap unique non-plate GUIDs into nomenclature_guid."""

from __future__ import annotations

import sqlite3
from collections import defaultdict
from pathlib import Path

import pytest

from core.nomenclature_guid import ensure_schema, get_by_mark, upsert
from scripts.bootstrap_nomenclature_guid import (
    BootstrapCandidate,
    apply_bootstrap,
    collect_auto_candidates,
    run_bootstrap,
)
from scripts.build_guid_queue import Item1C, match_catalog_row, norm

PLAIN_GUID = "11111111-1111-1111-1111-111111111111"
U_GUID = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
OTHER_GUID = "22222222-2222-2222-2222-222222222222"
THIRD_GUID = "33333333-3333-3333-3333-333333333333"
FBS_GUID = "44444444-4444-4444-4444-444444444444"
BRIDGE_GUID = "55555555-5555-5555-5555-555555555555"
MARCH_GUID = "66666666-6666-6666-6666-666666666666"
STEP_GUID = "77777777-7777-7777-7777-777777777777"


@pytest.fixture
def conn(tmp_path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(tmp_path / "pb.db")
    try:
        yield connection
    finally:
        connection.close()


@pytest.fixture
def db(conn: sqlite3.Connection) -> sqlite3.Connection:
    ensure_schema(conn)
    return conn


def _item(guid: str, name: str) -> Item1C:
    return Item1C(guid=guid.lower(), name=name, price=None, unit="шт")


def _name_idx(items: dict[str, list[Item1C]]):
    idx = {
        key: defaultdict(list)
        for key in ("piles", "bridge", "fbs", "marches", "steps", "plates")
    }
    for group, rows in items.items():
        for item in rows:
            idx[group][norm(item.name)].append(item)
    return idx


def _catalog(**groups: list[tuple[str, str, str]]) -> dict[str, list[tuple[str, str, str]]]:
    out = {
        "piles": [],
        "bridge": [],
        "fbs": [],
        "marches": [],
        "steps": [],
        "plates": [],
    }
    out.update(groups)
    return out


def _count(db: sqlite3.Connection) -> int:
    return db.execute("SELECT COUNT(*) FROM nomenclature_guid").fetchone()[0]


def _ready_fixtures() -> tuple[dict, dict]:
    name_idx = _name_idx(
        {
            "piles": [
                _item(PLAIN_GUID, "Сваи С 60.30-6"),
                _item(U_GUID, "Сваи С 60.30-6у"),
            ],
            "bridge": [_item(BRIDGE_GUID, "Сваи мостовые С13-40T7")],
            "fbs": [_item(FBS_GUID, "Блоки ФБС 9.3.6-Т")],
            "marches": [_item(MARCH_GUID, "Лестничные марши 1ЛМ 30-11-15-4")],
            "steps": [_item(STEP_GUID, "Лестничные ступени ЛС-11")],
        }
    )
    catalog = _catalog(
        piles=[("С60.30-6", "С60.30-6", "pile_prices")],
        bridge=[("С13-40T7", "С13-40T7", "bridge_pile_prices")],
        fbs=[("ФБС 9.3.6-Т", "ФБС 9.3.6-Т", "fbs_prices")],
        marches=[("Лестничные марши 1ЛМ 30-11-15-4", "1ЛМ 30-11-15-4", "march_prices")],
        steps=[("Лестничные ступени ЛС-11", "ЛС11", "step_prices")],
    )
    return name_idx, catalog


def test_auto_upserts_then_idempotent_rerun(db: sqlite3.Connection) -> None:
    name_idx, catalog = _ready_fixtures()

    first = run_bootstrap(db, name_idx, catalog)
    db.commit()
    assert first.inserted == 5
    assert first.updated == 0
    assert first.skipped_manual == 0
    assert first.skipped_dup == 0
    assert first.skipped_not_ready == 0
    assert first.by_kind == {
        "pile": 1,
        "bridge_pile": 1,
        "fbs": 1,
        "stair_flight": 1,
        "stair_step": 1,
    }
    assert _count(db) == 5

    pile = get_by_mark(db, "pile", "С60.30-6")
    assert pile is not None
    assert pile.guid_1c == PLAIN_GUID
    assert pile.guid_1c_u == U_GUID
    assert pile.match_status == "auto"
    assert get_by_mark(db, "pile", "С60.30-6у") is None

    march = get_by_mark(db, "stair_flight", "1ЛМ 30-11-15-4")
    assert march is not None
    assert march.guid_1c == MARCH_GUID
    step = get_by_mark(db, "stair_step", "ЛС11")
    assert step is not None
    assert step.guid_1c == STEP_GUID

    second = run_bootstrap(db, name_idx, catalog)
    assert second.inserted == 0
    assert second.updated == 0
    assert second.skipped_manual == 0
    assert second.by_kind["pile"] == 1
    assert _count(db) == 5
    assert get_by_mark(db, "pile", "С60.30-6").guid_1c == PLAIN_GUID


def test_duplicate_names_not_inserted_as_auto(db: sqlite3.Connection) -> None:
    name_idx = _name_idx(
        {
            "piles": [
                _item(PLAIN_GUID, "Сваи С 50.30-8"),
                _item(OTHER_GUID, "Сваи C 50.30-8"),
            ]
        }
    )
    catalog = _catalog(piles=[("С50.30-8", "С50.30-8", "pile_prices")])

    matched = match_catalog_row("piles", "С50.30-8", "С50.30-8", name_idx)
    assert matched.status == "ambig"

    report = run_bootstrap(db, name_idx, catalog)
    assert report.inserted == 0
    assert report.skipped_not_ready == 1
    assert _count(db) == 0
    assert get_by_mark(db, "pile", "С50.30-8") is None


def test_u_guid_fills_guid_1c_u_of_plain_mark(db: sqlite3.Connection) -> None:
    name_idx = _name_idx(
        {
            "piles": [
                _item(PLAIN_GUID, "Сваи С 40.30-6"),
                _item(U_GUID, "Сваи С 40.30-6у"),
            ]
        }
    )
    catalog = _catalog(piles=[("С40.30-6", "С40.30-6", "pile_prices")])

    report = run_bootstrap(db, name_idx, catalog)
    assert report.inserted == 1
    assert report.with_guid_u == 1
    row = get_by_mark(db, "pile", "С40.30-6")
    assert row is not None
    assert row.guid_1c == PLAIN_GUID
    assert row.guid_1c_u == U_GUID
    assert _count(db) == 1


def test_u_only_pile_not_bootstrapped(db: sqlite3.Connection) -> None:
    name_idx = _name_idx({"piles": [_item(U_GUID, "Сваи С 30.30-3у")]})
    catalog = _catalog(piles=[("С30.30-3", "С30.30-3", "pile_prices")])

    report = run_bootstrap(db, name_idx, catalog)
    assert report.inserted == 0
    assert report.skipped_not_ready == 1
    assert _count(db) == 0


def test_missing_in_1c_not_bootstrapped(db: sqlite3.Connection) -> None:
    name_idx = _name_idx({})
    catalog = _catalog(piles=[("С99.30-12", "С99.30-12", "pile_prices")])

    report = run_bootstrap(db, name_idx, catalog)
    assert report.inserted == 0
    assert report.skipped_not_ready == 1
    assert _count(db) == 0


def test_manual_row_not_clobbered(db: sqlite3.Connection) -> None:
    upsert(
        db,
        "pile",
        "С60.30-6",
        guid_1c=THIRD_GUID,
        match_status="manual",
        match_note="выбрала бухгалтерия",
    )
    name_idx, catalog = _ready_fixtures()
    catalog = _catalog(piles=catalog["piles"])

    report = run_bootstrap(db, name_idx, catalog)
    assert report.inserted == 0
    assert report.updated == 0
    assert report.skipped_manual == 1
    row = get_by_mark(db, "pile", "С60.30-6")
    assert row is not None
    assert row.guid_1c == THIRD_GUID
    assert row.match_status == "manual"
    assert row.match_note == "выбрала бухгалтерия"


def test_plates_never_inserted(db: sqlite3.Connection) -> None:
    name_idx = _name_idx(
        {"plates": [_item(PLAIN_GUID, "Плиты ПБ 51-12-8")]}
    )
    catalog = _catalog(
        plates=[(PLAIN_GUID, "Плиты ПБ 51-12-8", "prays_plity")],
        piles=[],
    )

    report = run_bootstrap(db, name_idx, catalog, groups=("plates", "piles"))
    assert report.inserted == 0
    assert _count(db) == 0
    kinds = {
        row[0]
        for row in db.execute("SELECT DISTINCT product_kind FROM nomenclature_guid")
    }
    assert kinds == set()


def test_dry_run_does_not_write(db: sqlite3.Connection) -> None:
    name_idx, catalog = _ready_fixtures()

    report = run_bootstrap(db, name_idx, catalog, dry_run=True)
    assert report.inserted == 5
    assert _count(db) == 0

    run_bootstrap(db, name_idx, catalog, dry_run=False)
    assert _count(db) == 5


def test_duplicate_guid_across_marks_skipped(db: sqlite3.Connection) -> None:
    name_idx = _name_idx(
        {
            "piles": [_item(PLAIN_GUID, "Сваи С 60.30-6")],
            "fbs": [_item(PLAIN_GUID, "Блоки ФБС 9.3.6-Т")],
        }
    )
    catalog = _catalog(
        piles=[("С60.30-6", "С60.30-6", "pile_prices")],
        fbs=[("ФБС 9.3.6-Т", "ФБС 9.3.6-Т", "fbs_prices")],
    )

    candidates, skipped_not_ready, skipped_dup, dup_labels = collect_auto_candidates(
        name_idx, catalog
    )
    assert candidates == []
    assert skipped_not_ready == 0
    assert skipped_dup == 2
    assert set(dup_labels) == {"pile:С60.30-6", "fbs:ФБС 9.3.6-Т"}

    report = apply_bootstrap(db, candidates)
    assert report.inserted == 0
    assert _count(db) == 0


def test_updates_existing_missing_row(db: sqlite3.Connection) -> None:
    upsert(db, "fbs", "ФБС 9.3.6-Т", match_status="missing")
    name_idx = _name_idx({"fbs": [_item(FBS_GUID, "Блоки ФБС 9.3.6-Т")]})
    catalog = _catalog(fbs=[("ФБС 9.3.6-Т", "ФБС 9.3.6-Т", "fbs_prices")])

    report = run_bootstrap(db, name_idx, catalog)
    assert report.inserted == 0
    assert report.updated == 1
    row = get_by_mark(db, "fbs", "ФБС 9.3.6-Т")
    assert row is not None
    assert row.guid_1c == FBS_GUID
    assert row.match_status == "auto"
