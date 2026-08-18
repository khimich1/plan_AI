"""TDD: разовый геокодинг gsm_station с lat IS NULL.

Тесты на temp sqlite — существующие координаты не затираются,
обновляются только NULL, повторный запуск с кэшем не ходит в сеть.
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path

from scripts.build_gsm_routes_map import load_address_cache, save_address_cache
from scripts.build_gsm_trip_pool import normalize_address
from scripts.geocode_gsm_stations import (
    coords_plausible_for_address,
    geocode_null_stations,
    highway_query_variants,
    locality_query_variants,
    station_query_variants,
)


def _make_db(tmp_path: Path) -> Path:
    db_path = tmp_path / "stations.db"
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE gsm_station (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                address TEXT NOT NULL UNIQUE,
                brand TEXT,
                lat REAL,
                lon REAL,
                geocode_source TEXT
            )
            """
        )
    return db_path


def _insert_station(
    db_path: Path,
    *,
    address: str,
    brand: str = "X",
    lat: float | None = None,
    lon: float | None = None,
    geocode_source: str | None = None,
) -> int:
    with sqlite3.connect(db_path) as conn:
        cur = conn.execute(
            """
            INSERT INTO gsm_station (address, brand, lat, lon, geocode_source)
            VALUES (?, ?, ?, ?, ?)
            """,
            (address, brand, lat, lon, geocode_source),
        )
        return int(cur.lastrowid)


def _get_station(db_path: Path, station_id: int) -> dict:
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT * FROM gsm_station WHERE id = ?",
            (station_id,),
        ).fetchone()
    assert row is not None
    return dict(row)


def test_does_not_overwrite_existing_lat_lon(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    kept_id = _insert_station(
        db_path,
        address="г. Кострома, ул. Магистральная, д. 8",
        brand="КТК",
        lat=57.766,
        lon=40.927,
        geocode_source="manual",
    )
    null_id = _insert_station(
        db_path,
        address="М8, 87 км, справа, Московская обл",
        brand="TATNEFT",
    )
    cache_path = tmp_path / "addresses.json"
    calls: list[str] = []

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        calls.append(query)
        return 56.2, 38.1, "М-8, 87 км"

    report = geocode_null_stations(
        db_path,
        cache_path,
        geocode_fn=fake_geocode,
        sleep_sec=0,
    )

    kept = _get_station(db_path, kept_id)
    assert kept["lat"] == 57.766
    assert kept["lon"] == 40.927
    assert kept["geocode_source"] == "manual"
    assert report.skipped == 1

    updated = _get_station(db_path, null_id)
    assert updated["lat"] == 56.2
    assert updated["lon"] == 38.1
    assert updated["geocode_source"] == "nominatim"
    assert calls  # NULL-станция пошла в геокодер


def test_updates_only_null_coords(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    _insert_station(
        db_path,
        address="уже есть",
        lat=55.0,
        lon=37.0,
        geocode_source="nominatim",
    )
    null_id = _insert_station(db_path, address="нужны координаты")
    cache_path = tmp_path / "addresses.json"

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        return 58.0, 41.0, query

    report = geocode_null_stations(
        db_path,
        cache_path,
        geocode_fn=fake_geocode,
        sleep_sec=0,
    )

    assert report.from_network == 1
    assert report.from_cache == 0
    assert report.failed == 0
    assert report.skipped == 1
    assert report.geocoded == 1

    updated = _get_station(db_path, null_id)
    assert updated["lat"] == 58.0
    assert updated["lon"] == 41.0
    assert updated["geocode_source"] == "nominatim"


def test_second_run_with_cache_does_not_call_network(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    address = "МКАД, 104-й км, Балашиха"
    station_id = _insert_station(db_path, address=address)
    cache_path = tmp_path / "addresses.json"
    calls: list[str] = []

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        calls.append(query)
        return 55.8, 37.9, "МКАД 104 км"

    first = geocode_null_stations(
        db_path,
        cache_path,
        geocode_fn=fake_geocode,
        sleep_sec=0,
    )
    assert first.from_network == 1
    assert calls

    # Сбрасываем только БД — кэш остаётся. Повтор не должен звать сеть.
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE gsm_station SET lat = NULL, lon = NULL, geocode_source = NULL"
            " WHERE id = ?",
            (station_id,),
        )

    calls.clear()
    second = geocode_null_stations(
        db_path,
        cache_path,
        geocode_fn=fake_geocode,
        sleep_sec=0,
    )
    assert calls == []
    assert second.from_cache == 1
    assert second.from_network == 0
    restored = _get_station(db_path, station_id)
    assert restored["lat"] == 55.8
    assert restored["lon"] == 37.9
    assert restored["geocode_source"] == "nominatim"


def test_failed_geocode_marks_source_keeps_null(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    station_id = _insert_station(
        db_path,
        address="Кумохинский сельсовет, нигде не найти",
        brand="Опти",
    )
    cache_path = tmp_path / "addresses.json"

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        return None

    report = geocode_null_stations(
        db_path,
        cache_path,
        geocode_fn=fake_geocode,
        sleep_sec=0,
    )

    assert report.failed == 1
    assert report.geocoded == 0
    row = _get_station(db_path, station_id)
    assert row["lat"] is None
    assert row["lon"] is None
    assert row["geocode_source"] == "failed"
    assert report.failed_stations
    assert report.failed_stations[0][0] == station_id


def test_force_does_not_overwrite_existing_db_coords(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    kept_id = _insert_station(
        db_path,
        address="уже в БД",
        lat=10.0,
        lon=20.0,
        geocode_source="manual",
    )
    null_id = _insert_station(db_path, address="ещё NULL")
    cache_path = tmp_path / "addresses.json"
    norm = normalize_address("ещё NULL")
    save_address_cache(
        cache_path,
        {
            "version": 1,
            "addresses": {
                norm: {
                    "lat": None,
                    "lon": None,
                    "display": "ещё NULL",
                    "query": "ещё NULL",
                    "source": "nominatim",
                    "error": "not_found",
                }
            },
        },
    )

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        return 1.0, 2.0, query

    report = geocode_null_stations(
        db_path,
        cache_path,
        force=True,
        geocode_fn=fake_geocode,
        sleep_sec=0,
    )

    kept = _get_station(db_path, kept_id)
    assert kept["lat"] == 10.0
    assert kept["lon"] == 20.0
    assert kept["geocode_source"] == "manual"
    assert report.skipped == 1

    updated = _get_station(db_path, null_id)
    assert updated["lat"] == 1.0
    assert updated["lon"] == 2.0
    assert updated["geocode_source"] == "nominatim"


def test_offline_uses_cache_only(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    cached_id = _insert_station(db_path, address="в кэше")
    missing_id = _insert_station(db_path, address="нет в кэше")
    cache_path = tmp_path / "addresses.json"
    save_address_cache(
        cache_path,
        {
            "version": 1,
            "addresses": {
                normalize_address("в кэше"): {
                    "lat": 57.1,
                    "lon": 40.2,
                    "display": "в кэше",
                    "query": "в кэше",
                    "source": "nominatim",
                    "error": None,
                }
            },
        },
    )

    def boom(query: str) -> tuple[float, float, str] | None:
        raise AssertionError(f"offline не должен звать сеть: {query}")

    report = geocode_null_stations(
        db_path,
        cache_path,
        offline=True,
        geocode_fn=boom,
        sleep_sec=0,
    )

    assert report.from_cache == 1
    assert report.from_network == 0
    assert report.failed == 1
    assert _get_station(db_path, cached_id)["lat"] == 57.1
    missing = _get_station(db_path, missing_id)
    assert missing["lat"] is None
    assert missing["geocode_source"] == "failed"


def test_highway_query_variants_m8_km() -> None:
    variants = highway_query_variants("М8, 87 км, справа, Московская обл")
    assert "М-8, 87 км, Московская область, Россия" in variants
    assert "М-8, 87 км, Россия" in variants


def test_locality_query_variants_village_and_postal() -> None:
    variants = locality_query_variants(
        "Ярославская обл., с. Деболовское, тер. Деболовское, зд. 4"
    )
    assert "Деболовское, Ярославская область, Россия" in variants
    postal = locality_query_variants(
        "Ярославское шоссе 43 км, ст1, Лесной, Московская обл., Россия, 141231"
    )
    assert "Лесной, Московская область, 141231, Россия" in postal


def test_highway_query_variants_mkad() -> None:
    variants = highway_query_variants(
        "МКАД, 104-й км, Балашиха, Московская обл., Россия, 105484"
    )
    assert "МКАД, 104 км, Россия" in variants


def test_highway_query_variants_m7_km_before_code() -> None:
    variants = highway_query_variants(
        "183 км, М7 Волга (справа), Крутово, Ивановская обл."
    )
    assert "М-7, 183 км, Россия" in variants


def test_locality_troitsk_also_tries_moscow() -> None:
    variants = locality_query_variants(
        "Калужское ш., 12, Троицк, Московская обл., Россия, 142191"
    )
    assert "Троицк, Москва, Россия" in variants


def test_locality_extracts_yam_and_pereslavl() -> None:
    yam = locality_query_variants(
        "М-8, а/д Холмогоры, село Ям, городской округ Переславль-Залесский, "
        "Ярославская область, Россия, 152024"
    )
    assert "Ям, Ярославская область, Россия" in yam
    assert "Переславль-Залесский, Ярославская область, Россия" in yam


def test_station_query_variants_include_highway_and_simplified() -> None:
    variants = station_query_variants("М8, 87 км, справа, Московская обл")
    assert "М-8, 87 км, Россия" in variants
    assert any("Московская" in item for item in variants)


def test_rejects_nominatim_hit_outside_address_region(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    station_id = _insert_station(
        db_path,
        address="183 км, М7 Волга, Крутово, Ивановская обл., Россия",
    )
    cache_path = tmp_path / "addresses.json"

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        if "Крутово" in query:
            return 56.76, 40.94, "Крутово, Ивановская область"
        if "М-7" in query or "М7" in query:
            return 53.63, 83.72, "Алтайский край"
        return None

    report = geocode_null_stations(
        db_path,
        cache_path,
        geocode_fn=fake_geocode,
        sleep_sec=0,
    )
    assert report.from_network == 1
    row = _get_station(db_path, station_id)
    assert abs(row["lat"] - 56.76) < 0.01
    assert abs(row["lon"] - 40.94) < 0.01
    assert coords_plausible_for_address(
        "Крутово, Ивановская обл.", 56.76, 40.94
    )
    assert not coords_plausible_for_address(
        "Крутово, Ивановская обл.", 53.63, 83.72
    )


def test_report_format_counts(tmp_path: Path) -> None:
    db_path = _make_db(tmp_path)
    _insert_station(db_path, address="есть", lat=1.0, lon=2.0)
    _insert_station(db_path, address="нет")
    cache_path = tmp_path / "addresses.json"

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        return 3.0, 4.0, query

    report = geocode_null_stations(
        db_path,
        cache_path,
        geocode_fn=fake_geocode,
        sleep_sec=0,
    )
    text = report.format()
    assert "1/1 геокодировано" in text
    assert "из кэша" in text
    assert "с сети" in text
    assert "failed" in text
    assert "пропущено" in text

    cached = load_address_cache(cache_path)
    entry = cached["addresses"][normalize_address("нет")]
    assert entry["lat"] == 3.0
    assert json.dumps(entry)  # запись кэша сериализуема
