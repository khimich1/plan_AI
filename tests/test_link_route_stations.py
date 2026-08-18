"""TDD: заполнение gsm_route.typical_station_ids (история + география).

Тесты на temp sqlite + мок геокода. Сеть не используется.
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path

from scripts.build_gsm_routes_map import save_address_cache
from scripts.build_gsm_trip_pool import normalize_address
from scripts.link_route_stations import link_route_stations

# Контрольные точки: Кострома Кузнецкая → Ярославль Домостроителей.
# Станция КТК Магистральная лежит на отрезке (~1.6 км).
ADDR_A = "г.Кострома, ул.Кузнецкая, д.18Б"
ADDR_B = "г.Ярославль, пр-д Домостроителей, д.1, стр.3"
ADDR_FAR = "г.Москва, Красная площадь, д.1"

LAT_A, LON_A = 57.7622656, 40.9582651
LAT_B, LON_B = 57.6647563, 39.930913
LAT_ST1, LON_ST1 = 57.7446259, 40.9238723
LAT_FAR, LON_FAR = 55.755, 37.617

STATION_1_ADDR = "Кострома, ул. Магистральная, д. 8"
STATION_51_ADDR = "г. Кострома, ул. Магистральная, д. 8"


def _make_db(tmp_path: Path) -> Path:
    db_path = tmp_path / "link.db"
    with sqlite3.connect(db_path) as conn:
        conn.executescript(
            """
            CREATE TABLE gsm_vehicle (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                plate_number TEXT NOT NULL,
                tank_volume_liters REAL NOT NULL DEFAULT 70,
                norm_summer REAL NOT NULL DEFAULT 10,
                norm_winter REAL NOT NULL DEFAULT 12,
                is_active INTEGER NOT NULL DEFAULT 1
            );
            CREATE TABLE gsm_station (
                id INTEGER PRIMARY KEY,
                address TEXT NOT NULL UNIQUE,
                brand TEXT,
                lat REAL,
                lon REAL,
                geocode_source TEXT
            );
            CREATE TABLE gsm_route (
                id INTEGER PRIMARY KEY,
                vehicle_id INTEGER NOT NULL,
                addr_a TEXT NOT NULL,
                addr_b TEXT NOT NULL,
                km INTEGER NOT NULL,
                frequency INTEGER NOT NULL DEFAULT 1,
                typical_station_ids TEXT
            );
            CREATE TABLE gsm_driver (
                id INTEGER PRIMARY KEY,
                full_name TEXT NOT NULL,
                license_number TEXT NOT NULL DEFAULT 'x',
                is_active INTEGER NOT NULL DEFAULT 1
            );
            CREATE TABLE gsm_fuel_card (
                id INTEGER PRIMARY KEY,
                card_number TEXT NOT NULL UNIQUE,
                vehicle_id INTEGER,
                assigned_at TEXT NOT NULL DEFAULT '2026-01-01',
                archived_at TEXT
            );
            CREATE TABLE gsm_import_batch (
                id INTEGER PRIMARY KEY,
                filename TEXT NOT NULL DEFAULT 't.xls',
                uploaded_at TEXT NOT NULL DEFAULT '2026-01-01'
            );
            CREATE TABLE gsm_transaction (
                id INTEGER PRIMARY KEY,
                card_id INTEGER NOT NULL,
                ts TEXT NOT NULL,
                service_type TEXT NOT NULL,
                qty_liters REAL,
                amount REAL NOT NULL DEFAULT 1,
                station_id INTEGER,
                raw_address TEXT NOT NULL DEFAULT '',
                batch_id INTEGER NOT NULL DEFAULT 1
            );
            CREATE TABLE gsm_waybill (
                id INTEGER PRIMARY KEY,
                vehicle_id INTEGER NOT NULL,
                date TEXT NOT NULL,
                driver_id INTEGER NOT NULL DEFAULT 1,
                status TEXT NOT NULL DEFAULT 'confirmed',
                source TEXT NOT NULL DEFAULT 'auto',
                route_json TEXT NOT NULL,
                UNIQUE(vehicle_id, date)
            );
            """
        )
        conn.execute(
            "INSERT INTO gsm_vehicle (id, name, plate_number) VALUES (1, 'Hyundai Palisade', 'О 521 УХ 44')"
        )
        conn.execute(
            "INSERT INTO gsm_driver (id, full_name) VALUES (1, 'Тестов Т.Т.')"
        )
        conn.execute("INSERT INTO gsm_import_batch (id) VALUES (1)")
    return db_path


def _insert_station(
    db_path: Path,
    *,
    station_id: int,
    address: str,
    brand: str = "КТК",
    lat: float | None = None,
    lon: float | None = None,
) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO gsm_station (id, address, brand, lat, lon, geocode_source)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (station_id, address, brand, lat, lon, "nominatim" if lat is not None else None),
        )


def _insert_route(
    db_path: Path,
    *,
    route_id: int,
    vehicle_id: int = 1,
    addr_a: str = ADDR_A,
    addr_b: str = ADDR_B,
    km: int = 95,
    typical_station_ids: str | None = None,
) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO gsm_route
                (id, vehicle_id, addr_a, addr_b, km, frequency, typical_station_ids)
            VALUES (?, ?, ?, ?, ?, 1, ?)
            """,
            (route_id, vehicle_id, addr_a, addr_b, km, typical_station_ids),
        )


def _write_cache(path: Path, mapping: dict[str, tuple[float, float]]) -> None:
    addresses: dict[str, dict] = {}
    for raw, (lat, lon) in mapping.items():
        addresses[normalize_address(raw)] = {
            "lat": lat,
            "lon": lon,
            "display": raw,
            "source": "test",
            "error": None,
        }
    save_address_cache(path, {"version": 1, "addresses": addresses})


def _typical_ids(db_path: Path, route_id: int) -> list[int]:
    with sqlite3.connect(db_path) as conn:
        raw = conn.execute(
            "SELECT typical_station_ids FROM gsm_route WHERE id = ?",
            (route_id,),
        ).fetchone()[0]
    if raw is None or raw == "":
        return []
    return [int(x) for x in json.loads(raw)]


def _run(
    db_path: Path,
    cache_path: Path,
    *,
    geocode_fn=None,
    offline: bool = False,
    threshold_km: float = 15.0,
):
    return link_route_stations(
        db_path,
        cache_path,
        offline=offline,
        threshold_km=threshold_km,
        geocode_fn=geocode_fn,
        sleep_sec=0,
    )


def test_station_on_segment_is_linked(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    _insert_route(db_path, route_id=10)
    _insert_station(
        db_path,
        station_id=1,
        address=STATION_1_ADDR,
        lat=LAT_ST1,
        lon=LON_ST1,
    )
    cache_path = tmp_path / "addresses.json"
    _write_cache(cache_path, {ADDR_A: (LAT_A, LON_A), ADDR_B: (LAT_B, LON_B)})

    report = _run(db_path, cache_path, offline=True)

    assert 1 in _typical_ids(db_path, 10)
    assert report.routes_filled >= 1


def test_far_station_is_not_linked(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    _insert_route(db_path, route_id=10)
    _insert_station(
        db_path,
        station_id=20,
        address=ADDR_FAR,
        brand="X",
        lat=LAT_FAR,
        lon=LON_FAR,
    )
    cache_path = tmp_path / "addresses.json"
    _write_cache(cache_path, {ADDR_A: (LAT_A, LON_A), ADDR_B: (LAT_B, LON_B)})

    _run(db_path, cache_path, offline=True)

    assert 20 not in _typical_ids(db_path, 10)
    assert _typical_ids(db_path, 10) == []


def test_station_without_coords_is_skipped(tmp_path: Path, capsys) -> None:
    db_path = _make_db(tmp_path)
    _insert_route(db_path, route_id=10)
    _insert_station(db_path, station_id=7, address="без координат, нигде")
    _insert_station(
        db_path,
        station_id=1,
        address=STATION_1_ADDR,
        lat=LAT_ST1,
        lon=LON_ST1,
    )
    cache_path = tmp_path / "addresses.json"
    _write_cache(cache_path, {ADDR_A: (LAT_A, LON_A), ADDR_B: (LAT_B, LON_B)})

    report = _run(db_path, cache_path, offline=True)

    assert 1 in _typical_ids(db_path, 10)
    assert 7 not in _typical_ids(db_path, 10)
    assert report.stations_skipped_no_coords >= 1
    captured = capsys.readouterr()
    assert "7" in captured.out
    assert "координат" in captured.out.lower() or "skip" in captured.out.lower()


def test_does_not_overwrite_existing_typical_ids_can_merge(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    _insert_route(db_path, route_id=10, typical_station_ids="[99]")
    _insert_station(
        db_path,
        station_id=1,
        address=STATION_1_ADDR,
        lat=LAT_ST1,
        lon=LON_ST1,
    )
    cache_path = tmp_path / "addresses.json"
    _write_cache(cache_path, {ADDR_A: (LAT_A, LON_A), ADDR_B: (LAT_B, LON_B)})

    _run(db_path, cache_path, offline=True)

    ids = _typical_ids(db_path, 10)
    assert 99 in ids
    assert 1 in ids


def test_palisade_prefers_station_id_1_over_duplicate(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    _insert_route(db_path, route_id=1, vehicle_id=1, addr_a=ADDR_A, addr_b=ADDR_B)
    _insert_station(
        db_path,
        station_id=1,
        address=STATION_1_ADDR,
        lat=LAT_ST1,
        lon=LON_ST1,
    )
    _insert_station(
        db_path,
        station_id=51,
        address=STATION_51_ADDR,
        lat=LAT_ST1,
        lon=LON_ST1,
    )
    cache_path = tmp_path / "addresses.json"
    _write_cache(cache_path, {ADDR_A: (LAT_A, LON_A), ADDR_B: (LAT_B, LON_B)})

    report = _run(db_path, cache_path, offline=True)

    ids = _typical_ids(db_path, 1)
    assert 1 in ids
    assert 51 not in ids
    assert report.palisade_with_station_1 >= 1


def test_history_imported_waybill_binds_fuel_station(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    _insert_route(
        db_path,
        route_id=515,
        vehicle_id=1,
        addr_a=ADDR_A,
        addr_b="г.Ковров, ул.Строителей, д.28, ООО СЗ \"СК Континент Доброе\"",
        km=205,
    )
    _insert_station(
        db_path,
        station_id=1,
        address=STATION_1_ADDR,
        lat=LAT_ST1,
        lon=LON_ST1,
    )
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO gsm_fuel_card (id, card_number, vehicle_id) VALUES (1, '3005454271', 1)"
        )
        conn.execute(
            """
            INSERT INTO gsm_transaction
                (id, card_id, ts, service_type, station_id, raw_address, batch_id)
            VALUES (1, 1, '2026-05-29 08:49:24', 'fuel', 1, ?, 1)
            """,
            (STATION_1_ADDR,),
        )
        conn.execute(
            """
            INSERT INTO gsm_waybill (id, vehicle_id, date, source, route_json)
            VALUES (4, 1, '2026-05-29', 'imported', ?)
            """,
            (
                json.dumps(
                    [
                        {
                            "seq": 1,
                            "addr_from": ADDR_A,
                            "addr_to": 'г.Ковров, ул.Строителей, д.28, ООО СЗ "СК Континент Доброе"',
                            "km": 205.0,
                        }
                    ],
                    ensure_ascii=False,
                ),
            ),
        )
    cache_path = tmp_path / "addresses.json"
    _write_cache(cache_path, {})

    _run(db_path, cache_path, offline=True)

    assert 1 in _typical_ids(db_path, 515)
