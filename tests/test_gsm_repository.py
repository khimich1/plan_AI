"""GSM schema (9 gsm_* tables) + GsmRepository — Task T1 (TDD red)."""

from __future__ import annotations

import sqlite3
from datetime import date
from pathlib import Path

import pytest

from app.repositories.gsm_repository import GsmRepository
from core import kp_db_schema
from core.kp_db_common import _connect

GSM_TABLES = (
    "gsm_vehicle",
    "gsm_driver",
    "gsm_fuel_card",
    "gsm_station",
    "gsm_import_batch",
    "gsm_transaction",
    "gsm_route",
    "gsm_waybill",
    "gsm_setting",
)

VEHICLE_COLS = {
    "id",
    "name",
    "plate_number",
    "tank_volume_liters",
    "norm_summer",
    "norm_winter",
    "primary_driver_id",
    "is_active",
}
DRIVER_COLS = {
    "id",
    "full_name",
    "license_number",
    "license_issued_at",
    "personnel_number",
    "snils",
    "is_active",
}
CARD_COLS = {"id", "card_number", "vehicle_id", "assigned_at", "archived_at"}
STATION_COLS = {"id", "address", "brand", "lat", "lon", "geocode_source"}
BATCH_COLS = {
    "id",
    "filename",
    "period_from",
    "period_to",
    "rows_total",
    "sum_liters",
    "sum_amount",
    "uploaded_by",
    "uploaded_at",
}
TX_COLS = {
    "id",
    "card_id",
    "ts",
    "service_type",
    "fuel_grade",
    "qty_liters",
    "amount",
    "station_id",
    "raw_address",
    "batch_id",
}
ROUTE_COLS = {
    "id",
    "vehicle_id",
    "addr_a",
    "addr_b",
    "km",
    "frequency",
    "typical_station_ids",
}
WAYBILL_COLS = {
    "id",
    "vehicle_id",
    "date",
    "driver_id",
    "status",
    "source",
    "odometer_start",
    "odometer_end",
    "fuel_start",
    "fuel_issued",
    "fuel_end",
    "route_json",
    "warnings_json",
}
SETTING_COLS = {"key", "value"}


def _fresh_db(tmp_path: Path, name: str = "gsm.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _repo(tmp_path: Path, name: str = "gsm.db") -> GsmRepository:
    return GsmRepository(db_path=_fresh_db(tmp_path, name))


def _table_cols(conn: sqlite3.Connection, table: str) -> set[str]:
    return {row[1] for row in conn.execute(f"PRAGMA table_info({table})")}


def _indexed_columns(conn: sqlite3.Connection, table: str) -> dict[str, list[str]]:
    cur = conn.cursor()
    cur.execute(f"PRAGMA index_list({table})")
    result: dict[str, list[str]] = {}
    for row in cur.fetchall():
        index_name = row[1]
        cur.execute(f"PRAGMA index_info({index_name})")
        result[index_name] = [info[2] for info in cur.fetchall()]
    return result


def _unique_index_column_sets(conn: sqlite3.Connection, table: str) -> list[list[str]]:
    """Column lists of UNIQUE indexes (including those backing UNIQUE constraints)."""
    cur = conn.cursor()
    cur.execute(f"PRAGMA index_list({table})")
    sets: list[list[str]] = []
    for row in cur.fetchall():
        # row: (seq, name, unique, origin, partial)
        if not row[2]:
            continue
        index_name = row[1]
        cur.execute(f"PRAGMA index_info({index_name})")
        sets.append([info[2] for info in cur.fetchall()])
    return sets


def _seed_driver(repo: GsmRepository, *, full_name: str = "Кулигин Никита Валерьевич") -> int:
    return repo.create_driver(
        full_name=full_name,
        license_number="44 21 846315",
        license_issued_at="30.07.2015",
        personnel_number="143",
        snils="123-456-789 00",
    )


def _seed_vehicle(
    repo: GsmRepository,
    *,
    name: str = "Geely Tugella 848",
    plate_number: str = "О 848 ХР 44",
    primary_driver_id: int | None = None,
) -> int:
    return repo.create_vehicle(
        name=name,
        plate_number=plate_number,
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=primary_driver_id,
    )


def _seed_card(
    repo: GsmRepository,
    *,
    vehicle_id: int,
    card_number: str = "3005454268",
    assigned_at: str = "2025-01-01",
) -> int:
    return repo.create_card(
        card_number=card_number,
        vehicle_id=vehicle_id,
        assigned_at=assigned_at,
    )


# ---------------------------------------------------------------------------
# Schema: 9 tables + columns + indexes + uniqueness + idempotency
# ---------------------------------------------------------------------------


def test_ensure_schema_creates_nine_gsm_tables(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)

    with _connect(db_path) as conn:
        names = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'gsm_%'"
            )
        }
    assert names == set(GSM_TABLES)


def test_ensure_schema_gsm_table_columns(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)

    expected = {
        "gsm_vehicle": VEHICLE_COLS,
        "gsm_driver": DRIVER_COLS,
        "gsm_fuel_card": CARD_COLS,
        "gsm_station": STATION_COLS,
        "gsm_import_batch": BATCH_COLS,
        "gsm_transaction": TX_COLS,
        "gsm_route": ROUTE_COLS,
        "gsm_waybill": WAYBILL_COLS,
        "gsm_setting": SETTING_COLS,
    }
    with _connect(db_path) as conn:
        for table, cols in expected.items():
            assert _table_cols(conn, table) == cols, table


def test_ensure_schema_indexes_card_ts_and_vehicle_date(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)

    with _connect(db_path) as conn:
        tx_indexes = _indexed_columns(conn, "gsm_transaction")
        waybill_indexes = _indexed_columns(conn, "gsm_waybill")

    tx_col_sets = list(tx_indexes.values())
    waybill_col_sets = list(waybill_indexes.values())
    assert ["card_id", "ts"] in tx_col_sets
    assert ["vehicle_id", "date"] in waybill_col_sets


def test_gsm_transaction_unique_card_ts_qty_amount(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)

    with _connect(db_path) as conn:
        dedup_sql = conn.execute(
            "SELECT sql FROM sqlite_master "
            "WHERE type = 'index' AND name = 'idx_gsm_transaction_dedup'"
        ).fetchone()
        assert dedup_sql is not None
        assert dedup_sql[0] is not None
        assert "COALESCE(qty_liters,0)" in dedup_sql[0].replace(" ", "")
        unique_sets = _unique_index_column_sets(conn, "gsm_transaction")
    # Expression index: PRAGMA index_info returns NULL for COALESCE(...) column
    assert any(
        cols[:2] == ["card_id", "ts"] and cols[-1] == "amount" and len(cols) == 4
        for cols in unique_sets
    )

    # Behavioural check: duplicate fuel key raises IntegrityError
    with _connect(db_path) as conn:
        conn.execute(
            "INSERT INTO gsm_driver (full_name, license_number, is_active) "
            "VALUES ('Driver', '11 11 111111', 1)"
        )
        driver_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            "INSERT INTO gsm_vehicle "
            "(name, plate_number, tank_volume_liters, norm_summer, norm_winter, "
            "primary_driver_id, is_active) "
            "VALUES ('Car', 'A111AA44', 55, 9.4, 10.3, ?, 1)",
            (driver_id,),
        )
        vehicle_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            "INSERT INTO gsm_fuel_card (card_number, vehicle_id, assigned_at) "
            "VALUES ('3005454268', ?, '2025-01-01')",
            (vehicle_id,),
        )
        card_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            "INSERT INTO gsm_import_batch (filename, uploaded_at) "
            "VALUES ('t.xls', '2026-08-14T12:00:00')"
        )
        batch_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            "INSERT INTO gsm_transaction "
            "(card_id, ts, service_type, qty_liters, amount, raw_address, batch_id) "
            "VALUES (?, '2025-04-03T10:00:00', 'fuel', 40.0, 2500.0, 'АЗС 1', ?)",
            (card_id, batch_id),
        )
        conn.commit()

        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                "INSERT INTO gsm_transaction "
                "(card_id, ts, service_type, qty_liters, amount, raw_address, batch_id) "
                "VALUES (?, '2025-04-03T10:00:00', 'fuel', 40.0, 2500.0, 'АЗС 1', ?)",
                (card_id, batch_id),
            )
        conn.rollback()


def test_gsm_transaction_unique_dedupes_wash_with_null_qty(tmp_path: Path) -> None:
    """Washes store qty_liters NULL; UNIQUE must still reject duplicates (D14)."""
    db_path = _fresh_db(tmp_path)

    with _connect(db_path) as conn:
        conn.execute(
            "INSERT INTO gsm_driver (full_name, license_number, is_active) "
            "VALUES ('Driver', '11 11 111111', 1)"
        )
        driver_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            "INSERT INTO gsm_vehicle "
            "(name, plate_number, tank_volume_liters, norm_summer, norm_winter, "
            "primary_driver_id, is_active) "
            "VALUES ('Car', 'A111AA44', 55, 9.4, 10.3, ?, 1)",
            (driver_id,),
        )
        vehicle_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            "INSERT INTO gsm_fuel_card (card_number, vehicle_id, assigned_at) "
            "VALUES ('3005454268', ?, '2025-01-01')",
            (vehicle_id,),
        )
        card_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            "INSERT INTO gsm_import_batch (filename, uploaded_at) "
            "VALUES ('t.xls', '2026-08-14T12:00:00')"
        )
        batch_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            "INSERT INTO gsm_transaction "
            "(card_id, ts, service_type, qty_liters, amount, raw_address, batch_id) "
            "VALUES (?, '2025-04-03T11:00:00', 'wash', NULL, 500.0, 'Мойка 1', ?)",
            (card_id, batch_id),
        )
        conn.commit()

        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                "INSERT INTO gsm_transaction "
                "(card_id, ts, service_type, qty_liters, amount, raw_address, batch_id) "
                "VALUES (?, '2025-04-03T11:00:00', 'wash', NULL, 500.0, 'Мойка 1', ?)",
                (card_id, batch_id),
            )
        conn.rollback()


def test_ensure_schema_idempotent(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._init_schema_impl(db_path)

    with _connect(db_path) as conn:
        names = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'gsm_%'"
            )
        }
        cols = _table_cols(conn, "gsm_vehicle")
    assert names == set(GSM_TABLES)
    assert "plate_number" in cols
    assert "tank_volume_liters" in cols


# ---------------------------------------------------------------------------
# Repository: registry CRUD
# ---------------------------------------------------------------------------


def test_create_and_get_driver(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    driver_id = _seed_driver(repo)
    got = repo.get_driver(driver_id)
    assert got is not None
    assert got["id"] == driver_id
    assert got["full_name"] == "Кулигин Никита Валерьевич"
    assert got["license_number"] == "44 21 846315"
    assert got["is_active"] == 1


def test_list_and_update_driver(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    d1 = _seed_driver(repo, full_name="Driver One")
    d2 = repo.create_driver(
        full_name="Driver Two",
        license_number="44 22 000001",
        is_active=0,
    )
    active = repo.list_drivers(active_only=True)
    assert {d["id"] for d in active} == {d1}
    all_drivers = repo.list_drivers(active_only=False)
    assert {d["id"] for d in all_drivers} == {d1, d2}

    repo.update_driver(d1, personnel_number="999", snils="000-000-000 00")
    got = repo.get_driver(d1)
    assert got["personnel_number"] == "999"
    assert got["snils"] == "000-000-000 00"


def test_create_get_list_update_vehicle(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    driver_id = _seed_driver(repo)
    vehicle_id = _seed_vehicle(repo, primary_driver_id=driver_id)

    got = repo.get_vehicle(vehicle_id)
    assert got is not None
    assert got["name"] == "Geely Tugella 848"
    assert got["plate_number"] == "О 848 ХР 44"
    assert got["tank_volume_liters"] == pytest.approx(55.0)
    assert got["norm_summer"] == pytest.approx(9.4)
    assert got["norm_winter"] == pytest.approx(10.3)
    assert got["primary_driver_id"] == driver_id
    assert got["is_active"] == 1

    listed = repo.list_vehicles(active_only=True)
    assert len(listed) == 1
    assert listed[0]["id"] == vehicle_id

    repo.update_vehicle(vehicle_id, norm_summer=9.0, is_active=0)
    got2 = repo.get_vehicle(vehicle_id)
    assert got2["norm_summer"] == pytest.approx(9.0)
    assert got2["is_active"] == 0
    assert repo.list_vehicles(active_only=True) == []
    assert len(repo.list_vehicles(active_only=False)) == 1


def test_create_get_list_station(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    station_id = repo.create_station(
        address="г. Кострома, ул. Примерная, 1",
        brand="TATNEFT",
        lat=57.76,
        lon=40.92,
        geocode_source="cache",
    )
    got = repo.get_station(station_id)
    assert got["address"] == "г. Кострома, ул. Примерная, 1"
    assert got["brand"] == "TATNEFT"
    assert got["geocode_source"] == "cache"
    assert len(repo.list_stations()) == 1


def test_create_card_and_archive_sets_archived_at_no_delete(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    driver_id = _seed_driver(repo)
    vehicle_id = _seed_vehicle(repo, primary_driver_id=driver_id)
    card_id = _seed_card(repo, vehicle_id=vehicle_id)

    before = repo.get_card(card_id)
    assert before is not None
    assert before["archived_at"] is None
    assert before["card_number"] == "3005454268"

    active = repo.list_cards(include_archived=False)
    assert len(active) == 1

    repo.archive_card(card_id)

    after = repo.get_card(card_id)
    assert after is not None
    assert after["archived_at"] is not None
    assert isinstance(after["archived_at"], str)
    assert after["archived_at"] != ""

    # Row still exists in DB — no physical DELETE
    with _connect(repo.db_path) as conn:
        count = conn.execute(
            "SELECT COUNT(*) FROM gsm_fuel_card WHERE id = ?", (card_id,)
        ).fetchone()[0]
    assert count == 1

    assert repo.list_cards(include_archived=False) == []
    archived = repo.list_cards(include_archived=True)
    assert len(archived) == 1
    assert archived[0]["id"] == card_id


# ---------------------------------------------------------------------------
# Repository: transactions + waybills
# ---------------------------------------------------------------------------


def test_list_transactions_filters_by_vehicle_and_period(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    driver_id = _seed_driver(repo)
    v1 = _seed_vehicle(repo, name="Car 1", plate_number="A111AA44", primary_driver_id=driver_id)
    v2 = _seed_vehicle(repo, name="Car 2", plate_number="B222BB44", primary_driver_id=driver_id)
    c1 = _seed_card(repo, vehicle_id=v1, card_number="111")
    c2 = _seed_card(repo, vehicle_id=v2, card_number="222")

    batch_id = repo.create_import_batch(
        filename="card111.xls",
        uploaded_at="2026-08-14T12:00:00",
        uploaded_by="accountant",
        period_from="2025-04-01",
        period_to="2025-04-30",
    )
    # in period, vehicle 1
    repo.insert_transaction(
        card_id=c1,
        ts="2025-04-03T10:00:00",
        service_type="fuel",
        fuel_grade="АИ-95",
        qty_liters=40.0,
        amount=2500.0,
        raw_address="АЗС 1",
        batch_id=batch_id,
    )
    # outside period, vehicle 1
    repo.insert_transaction(
        card_id=c1,
        ts="2025-05-01T10:00:00",
        service_type="fuel",
        qty_liters=30.0,
        amount=1800.0,
        raw_address="АЗС 2",
        batch_id=batch_id,
    )
    # in period, other vehicle
    repo.insert_transaction(
        card_id=c2,
        ts="2025-04-10T12:00:00",
        service_type="wash",
        qty_liters=None,
        amount=500.0,
        raw_address="Мойка 1",
        batch_id=batch_id,
    )

    rows = repo.list_transactions(
        vehicle_id=v1,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
    )
    assert len(rows) == 1
    assert rows[0]["card_id"] == c1
    assert rows[0]["ts"].startswith("2025-04-03")
    assert rows[0]["qty_liters"] == pytest.approx(40.0)
    assert rows[0]["service_type"] == "fuel"

    # ISO strings also accepted for period bounds
    rows_str = repo.list_transactions(
        vehicle_id=v1,
        period_from="2025-04-01",
        period_to="2025-04-30",
    )
    assert len(rows_str) == 1


def test_upsert_waybill_insert_and_update(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    driver_id = _seed_driver(repo)
    vehicle_id = _seed_vehicle(repo, primary_driver_id=driver_id)

    waybill_id = repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date=date(2025, 4, 3),
        driver_id=driver_id,
        status="draft",
        source="auto",
        odometer_start=10000,
        odometer_end=10200,
        fuel_start=20.0,
        fuel_issued=40.0,
        fuel_end=35.0,
        route_json='[{"from":"A","to":"B","km":200}]',
        warnings_json=None,
    )
    assert isinstance(waybill_id, int)

    got = repo.get_waybill(vehicle_id, date(2025, 4, 3))
    assert got is not None
    assert got["id"] == waybill_id
    assert got["odometer_start"] == 10000
    assert got["fuel_end"] == pytest.approx(35.0)
    assert got["status"] == "draft"

    # Same (vehicle_id, date) → update, single row
    waybill_id2 = repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2025-04-03",
        driver_id=driver_id,
        status="confirmed",
        source="manual",
        odometer_start=10000,
        odometer_end=10250,
        fuel_start=20.0,
        fuel_issued=40.0,
        fuel_end=30.0,
        route_json='[{"from":"A","to":"C","km":250}]',
        warnings_json='["hook_above_threshold"]',
    )
    assert waybill_id2 == waybill_id

    with _connect(repo.db_path) as conn:
        count = conn.execute(
            "SELECT COUNT(*) FROM gsm_waybill WHERE vehicle_id = ? AND date = ?",
            (vehicle_id, "2025-04-03"),
        ).fetchone()[0]
    assert count == 1

    got2 = repo.get_waybill(vehicle_id, "2025-04-03")
    assert got2["status"] == "confirmed"
    assert got2["source"] == "manual"
    assert got2["odometer_end"] == 10250
    assert got2["fuel_end"] == pytest.approx(30.0)
    assert "hook_above_threshold" in got2["warnings_json"]


def test_create_and_list_routes(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    driver_id = _seed_driver(repo)
    v1 = _seed_vehicle(repo, name="Car 1", plate_number="A111AA44", primary_driver_id=driver_id)
    v2 = _seed_vehicle(repo, name="Car 2", plate_number="B222BB44", primary_driver_id=driver_id)

    r1 = repo.create_route(
        vehicle_id=v1,
        addr_a="Завод",
        addr_b="Объект А",
        km=42,
        frequency=2,
        typical_station_ids="[1,2]",
    )
    repo.create_route(
        vehicle_id=v2,
        addr_a="Завод",
        addr_b="Объект Б",
        km=10,
    )

    all_routes = repo.list_routes()
    assert len(all_routes) == 2
    for_v1 = repo.list_routes(vehicle_id=v1)
    assert len(for_v1) == 1
    assert for_v1[0]["id"] == r1
    assert for_v1[0]["km"] == 42
    assert for_v1[0]["typical_station_ids"] == "[1,2]"


def test_get_set_setting(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    assert repo.get_setting("winter_start") is None
    repo.set_setting("winter_start", "11-01")
    repo.set_setting("hook_threshold_km", "13")
    assert repo.get_setting("winter_start") == "11-01"
    assert repo.get_setting("hook_threshold_km") == "13"
    # upsert
    repo.set_setting("hook_threshold_km", "15")
    assert repo.get_setting("hook_threshold_km") == "15"
