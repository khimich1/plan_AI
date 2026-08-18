#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Разовый геокодинг станций gsm_station с lat/lon IS NULL.

Переиспользует Nominatim и кэш адресов из build_gsm_routes_map.
Существующие координаты в БД не затираются (только UPDATE NULL-полей).
"""

from __future__ import annotations

import argparse
import re
import sqlite3
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable

import requests

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.build_gsm_routes_map import (
    DEFAULT_GEO_CACHE,
    NOMINATIM_RATE_LIMIT_SEC,
    cache_entry_has_coords,
    cache_entry_should_skip_network,
    coerce_cache_entry_coords,
    load_address_cache,
    nominatim_geocode,
    nominatim_query_variants,
    save_address_cache,
)
from scripts.build_gsm_trip_pool import normalize_address

DEFAULT_DB = PROJECT_ROOT / "plita.db"

UPDATE_COORDS_SQL = """
UPDATE gsm_station
SET lat = ?, lon = ?, geocode_source = 'nominatim'
WHERE id = ? AND (lat IS NULL OR lon IS NULL)
"""
UPDATE_FAILED_SQL = """
UPDATE gsm_station
SET geocode_source = 'failed'
WHERE id = ? AND (lat IS NULL OR lon IS NULL)
"""

_FEDERAL_HW_RE = re.compile(
    r"(?:ФАД\s+)?(?P<code>[МM])[-\s]?(?P<num>\d+)\b",
    re.IGNORECASE,
)
_REGIONAL_HW_RE = re.compile(
    r"\b(?P<code>[РP])[-\s]?(?P<num>\d+)\b",
    re.IGNORECASE,
)
_MKAD_RE = re.compile(r"\bМКАД\b", re.IGNORECASE)
_KM_RE = re.compile(
    r"(?P<km>\d+)\s*(?:-?\s*й)?\s*(?:км\.?|километр)",
    re.IGNORECASE,
)
_BARE_KM_RE = re.compile(r"[,\s]+(?P<km>\d+)\b")
_NAMED_HW_RE = re.compile(
    r"(?P<stem>Ярославск|Калужск|Минск)(?:ое|ого)\s+ш(?:оссе\.?|\.)?",
    re.IGNORECASE,
)
_NAMED_CANON = {
    "ярославск": "Ярославское шоссе",
    "калужск": "Калужское шоссе",
    "минск": "Минское шоссе",
}
_REGION_RE = re.compile(
    r"(?P<name>Московская|Ярославская|Ивановская|Владимирская|"
    r"Костромская|Нижегородская|Калужская)\s+обл(?:асть|\.)?",
    re.IGNORECASE,
)
_POSTAL_RE = re.compile(r"\b(?P<postal>\d{6})\b")
_MARKED_SETTLEMENT_RE = re.compile(
    r"(?:^|[\s,])(?:г/п\s*|(?:г|с|д|п|пос)\.?\s*|село\s+|деревня\s+|"
    r"посёлок\s+|поселок\s+)"
    r"(?P<name>[А-ЯЁA-Z][А-Яа-яЁёA-Za-z\-]+(?:\s+[А-Яа-яЁёA-Za-z\-]+)?)",
)
_BARE_PLACE_RE = re.compile(
    r"(?:^|,\s*)(?P<name>[А-ЯЁ][А-Яа-яЁё\-]+(?:\s+[А-Яа-яЁё\-]+)?)(?=,|$)"
)
_POSELENIE_RE = re.compile(
    r"(?P<name>[А-ЯЁ][А-Яа-яЁё\-]+(?:ское|цкое|ский|цкий))\s+"
    r"(?:сельское\s+поселение|сельсовет)",
    re.IGNORECASE,
)
_SKIP_PLACE_RE = re.compile(
    r"область|\bобл\b|район|округ|поселение|сельсовет|шоссе|улица|километр|"
    r"справа|слева|автодорога|трасса|россия|м\.р-н|р-н|р-он|с/с|"
    r"тер\.|зд\.|ул\.|км\b|мкад|волга|крым|холмогоры",
    re.IGNORECASE,
)
_PERESLAVL_RE = re.compile(r"Переславл", re.IGNORECASE)
_OKRUG_RE = re.compile(
    r"(?P<name>[А-ЯЁ][А-Яа-яЁё\-]+(?:ий|ый|ой))\s+муниципальный\s+округ",
    re.IGNORECASE,
)
# Грубые bbox регионов (lat_min, lat_max, lon_min, lon_max).
_REGION_BBOX: dict[str, tuple[float, float, float, float]] = {
    "московская": (54.2, 57.0, 35.0, 40.5),
    "ярославская": (56.4, 58.9, 37.2, 41.6),
    "ивановская": (56.3, 57.9, 39.3, 43.3),
    "владимирская": (55.1, 56.9, 38.1, 42.6),
    "костромская": (57.0, 59.7, 40.2, 47.8),
    "нижегородская": (54.4, 58.0, 41.6, 48.0),
    "калужская": (53.3, 55.4, 33.3, 37.3),
}

GeocodeFn = Callable[[str], tuple[float, float, str] | None]


@dataclass(frozen=True)
class StationRow:
    id: int
    address: str
    brand: str
    lat: float | None
    lon: float | None


@dataclass(frozen=True)
class StationGeocodeReport:
    """Сводка прогона: кэш / сеть / failed / уже были координаты."""

    from_cache: int = 0
    from_network: int = 0
    failed: int = 0
    skipped: int = 0
    failed_stations: tuple[tuple[int, str, str], ...] = ()

    @property
    def geocoded(self) -> int:
        return self.from_cache + self.from_network

    def format(self) -> str:
        total_null = self.geocoded + self.failed
        return (
            f"{self.geocoded}/{total_null} геокодировано\n"
            f"  из кэша: {self.from_cache}\n"
            f"  с сети: {self.from_network}\n"
            f"  failed: {self.failed}\n"
            f"  пропущено (уже были координаты): {self.skipped}"
        )


def _extract_highway_code(text: str) -> str | None:
    federal = _FEDERAL_HW_RE.search(text)
    if federal:
        return f"М-{int(federal.group('num'))}"
    if _MKAD_RE.search(text):
        return "МКАД"
    regional = _REGIONAL_HW_RE.search(text)
    if regional:
        return f"Р-{int(regional.group('num'))}"
    return None


def _extract_km(text: str, *, after: int | None = None) -> int | None:
    explicit = _KM_RE.search(text)
    if explicit:
        return int(explicit.group("km"))
    if after is None:
        return None
    bare = _BARE_KM_RE.match(text[after:])
    if bare:
        return int(bare.group("km"))
    return None


def _highway_match_end(text: str) -> int | None:
    for pattern in (_FEDERAL_HW_RE, _MKAD_RE, _REGIONAL_HW_RE):
        match = pattern.search(text)
        if match:
            return match.end()
    return None


def _named_highway_km(text: str) -> tuple[str, int] | None:
    match = _NAMED_HW_RE.search(text)
    if not match:
        return None
    km_match = _KM_RE.search(text)
    if not km_match:
        return None
    stem = match.group("stem").casefold()
    name = _NAMED_CANON.get(stem)
    if not name:
        return None
    return name, int(km_match.group("km"))


def _extract_region(text: str) -> str | None:
    match = _REGION_RE.search(text)
    if not match:
        return None
    return f"{match.group('name').title()} область"


def _extract_postal(text: str) -> str | None:
    match = _POSTAL_RE.search(text)
    return match.group("postal") if match else None


def _is_useful_place(name: str) -> bool:
    cleaned = name.strip(" .,")
    if len(cleaned) < 2:
        return False
    return _SKIP_PLACE_RE.search(cleaned) is None


def _extract_settlements(text: str) -> list[str]:
    found: list[str] = []

    def _add(name: str) -> None:
        cleaned = re.sub(r"\s+", " ", name).strip(" .,")
        if _is_useful_place(cleaned) and cleaned not in found:
            found.append(cleaned)

    for match in _MARKED_SETTLEMENT_RE.finditer(text):
        _add(match.group("name"))
    for match in _POSELENIE_RE.finditer(text):
        _add(match.group("name"))
    for match in _OKRUG_RE.finditer(text):
        _add(match.group("name"))
    for match in _BARE_PLACE_RE.finditer(text):
        _add(match.group("name"))
    if _PERESLAVL_RE.search(text):
        _add("Переславль-Залесский")
    return found


def _add_unique(variants: list[str], query: str) -> None:
    q = query.strip()
    if q and q not in variants:
        variants.append(q)


def highway_query_variants(display: str) -> list[str]:
    """Варианты вроде «М-8, 87 км, Московская область, Россия»."""
    text = (display or "").strip()
    if not text:
        return []
    variants: list[str] = []
    region = _extract_region(text)
    code = _extract_highway_code(text)
    km = _extract_km(text, after=_highway_match_end(text) if code else None)
    if code and km is not None:
        if region:
            _add_unique(variants, f"{code}, {km} км, {region}, Россия")
        _add_unique(variants, f"{code}, {km} км, Россия")
        compact = code.replace("-", "")
        if compact != code:
            _add_unique(variants, f"{compact}, {km} км, Россия")
    elif code:
        if region:
            _add_unique(variants, f"{code}, {region}, Россия")
        _add_unique(variants, f"{code}, Россия")

    named = _named_highway_km(text)
    if named:
        name, named_km = named
        if region:
            _add_unique(variants, f"{name}, {named_km} км, {region}, Россия")
        _add_unique(variants, f"{name}, {named_km} км, Россия")
    return variants


def locality_query_variants(display: str) -> list[str]:
    """Населённый пункт + регион — надёжнее голого километра трассы."""
    text = (display or "").strip()
    if not text:
        return []
    variants: list[str] = []
    region = _extract_region(text)
    postal = _extract_postal(text)
    for place in _extract_settlements(text):
        if region and postal:
            _add_unique(variants, f"{place}, {region}, {postal}, Россия")
        if region:
            _add_unique(variants, f"{place}, {region}, Россия")
            if region.startswith("Московская"):
                _add_unique(variants, f"{place}, Москва, Россия")
        if "ярославск" in text.casefold():
            _add_unique(variants, f"{place}, Ярославль, Россия")
        if region is None:
            _add_unique(variants, f"{place}, Ярославская область, Россия")
            _add_unique(variants, f"{place}, Ярославль, Россия")
        _add_unique(variants, f"{place}, Россия")
    return variants


def coords_plausible_for_address(
    address: str, lat: float, lon: float
) -> bool:
    """Отсечь Nominatim-попадания в другой федеральный округ."""
    region = _extract_region(address)
    if not region:
        return True
    key = region.split()[0].casefold()
    bbox = _REGION_BBOX.get(key)
    if bbox is None:
        return True
    lat_min, lat_max, lon_min, lon_max = bbox
    return lat_min <= lat <= lat_max and lon_min <= lon <= lon_max


def station_query_variants(display: str) -> list[str]:
    """Населённый пункт, затем трасса+км+регион, затем остальные варианты."""
    variants: list[str] = []
    highway = highway_query_variants(display)
    specific = [query for query in highway if " км," in query and "область" in query]
    other_hw = [query for query in highway if query not in specific]
    for query in locality_query_variants(display):
        _add_unique(variants, query)
    for query in specific:
        _add_unique(variants, query)
    for query in other_hw:
        _add_unique(variants, query)
    for query in nominatim_query_variants(display):
        _add_unique(variants, query)
    return variants


def _has_coords(lat: Any, lon: Any) -> bool:
    return lat is not None and lon is not None


def _load_stations(db_path: Path) -> list[StationRow]:
    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            "SELECT id, address, brand, lat, lon FROM gsm_station ORDER BY id"
        ).fetchall()
    return [
        StationRow(
            id=int(row[0]),
            address=str(row[1] or ""),
            brand=str(row[2] or ""),
            lat=row[3],
            lon=row[4],
        )
        for row in rows
    ]


def _write_coords(db_path: Path, station_id: int, lat: float, lon: float) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(UPDATE_COORDS_SQL, (lat, lon, station_id))


def _write_failed(db_path: Path, station_id: int) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(UPDATE_FAILED_SQL, (station_id,))


def _cache_coords(entry: Any) -> tuple[float, float] | None:
    if not isinstance(entry, dict):
        return None
    coerce_cache_entry_coords(entry)
    if not cache_entry_has_coords(entry):
        return None
    return float(entry["lat"]), float(entry["lon"])


def _lookup_or_geocode(
    display: str,
    norm: str,
    addresses: dict[str, Any],
    *,
    offline: bool,
    force: bool,
    geocode_fn: GeocodeFn,
    sleep_fn: Callable[[float], None],
    sleep_sec: float,
    network_calls: list[int],
) -> tuple[tuple[float, float] | None, bool]:
    """Вернуть (coords|None, used_network). Пишет запись в addresses."""
    entry = addresses.get(norm)
    cached = _cache_coords(entry)
    cache_ok = cached is not None and coords_plausible_for_address(
        display, cached[0], cached[1]
    )
    if cache_ok and not force:
        return cached, False
    if offline:
        return (cached, False) if cache_ok else (None, False)
    if (
        not force
        and not cache_ok
        and cached is None
        and cache_entry_should_skip_network(entry, force_geocode=False)
    ):
        return None, False

    variants = station_query_variants(display)
    query = variants[0] if variants else display
    result: tuple[float, float, str] | None = None
    try:
        for query in variants:
            if network_calls[0] > 0 and sleep_sec > 0:
                sleep_fn(sleep_sec)
            network_calls[0] += 1
            hit = geocode_fn(query)
            if hit is None:
                continue
            if coords_plausible_for_address(display, hit[0], hit[1]):
                result = hit
                break
    except (requests.RequestException, ValueError, KeyError, TypeError) as exc:
        addresses[norm] = {
            "lat": None,
            "lon": None,
            "display": display,
            "query": query,
            "source": "nominatim",
            "error": str(exc),
        }
        return None, True

    if result is None:
        addresses[norm] = {
            "lat": None,
            "lon": None,
            "display": display,
            "query": query,
            "source": "nominatim",
            "error": "not_found",
        }
        return None, True

    lat, lon, nom_display = result
    addresses[norm] = {
        "lat": lat,
        "lon": lon,
        "display": nom_display,
        "query": query,
        "source": "nominatim",
        "error": None,
    }
    return (lat, lon), True


def geocode_null_stations(
    db_path: str | Path,
    cache_path: str | Path,
    *,
    offline: bool = False,
    force: bool = False,
    geocode_fn: GeocodeFn | None = None,
    sleep_fn: Callable[[float], None] = time.sleep,
    sleep_sec: float = NOMINATIM_RATE_LIMIT_SEC,
    session: requests.Session | None = None,
) -> StationGeocodeReport:
    """Геокодировать станции с NULL-координатами. Уже заполненные не трогает."""
    db = Path(db_path)
    cache = Path(cache_path)
    payload = load_address_cache(cache)
    addresses: dict[str, Any] = payload.setdefault("addresses", {})
    own_session = session is None and geocode_fn is None
    http = session or (requests.Session() if geocode_fn is None else None)
    resolve = geocode_fn or (
        lambda query: nominatim_geocode(query, session=http)
    )

    from_cache = 0
    from_network = 0
    failed = 0
    skipped = 0
    failed_rows: list[tuple[int, str, str]] = []
    dirty = False
    network_calls = [0]

    try:
        for station in _load_stations(db):
            if _has_coords(station.lat, station.lon):
                skipped += 1
                continue
            norm = normalize_address(station.address)
            coords, used_network = _lookup_or_geocode(
                station.address,
                norm,
                addresses,
                offline=offline,
                force=force,
                geocode_fn=resolve,
                sleep_fn=sleep_fn,
                sleep_sec=sleep_sec,
                network_calls=network_calls,
            )
            if used_network:
                dirty = True
            if coords is None:
                _write_failed(db, station.id)
                failed += 1
                failed_rows.append((station.id, station.brand, station.address))
                print(
                    f"failed id={station.id} {station.brand}: {station.address}",
                    flush=True,
                )
                continue
            _write_coords(db, station.id, coords[0], coords[1])
            if used_network:
                from_network += 1
            else:
                from_cache += 1
    finally:
        if dirty:
            save_address_cache(cache, payload)
        if own_session and http is not None:
            http.close()

    return StationGeocodeReport(
        from_cache=from_cache,
        from_network=from_network,
        failed=failed,
        skipped=skipped,
        failed_stations=tuple(failed_rows),
    )


def parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Геокодинг gsm_station с lat/lon IS NULL (кэш + Nominatim)."
    )
    parser.add_argument(
        "--db",
        type=Path,
        default=DEFAULT_DB,
        help=f"Путь к SQLite (по умолчанию: {DEFAULT_DB})",
    )
    parser.add_argument(
        "--cache",
        type=Path,
        default=DEFAULT_GEO_CACHE,
        help=f"Кэш адресов (по умолчанию: {DEFAULT_GEO_CACHE})",
    )
    parser.add_argument(
        "--offline",
        action="store_true",
        help="Только кэш, без запросов к Nominatim.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help=(
            "Повторить запрос для NULL-станций с отрицательным кэшем. "
            "Уже заполненные lat/lon в БД не затираются."
        ),
    )
    return parser.parse_args(list(argv) if argv is not None else None)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    report = geocode_null_stations(
        args.db,
        args.cache,
        offline=args.offline,
        force=args.force,
    )
    print(report.format())
    for station_id, brand, address in report.failed_stations:
        print(f"  failed id={station_id} {brand}: {address}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
