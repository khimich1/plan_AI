#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Заполнение gsm_route.typical_station_ids из истории ПЛ и географии.

Источники: (а) заправки дней imported-путевых; (б) станции ближе
``threshold_km`` к отрезку A–B. Существующие непустые списки не затираются
(только дополняются). Геокодинг концов маршрута — через кэш адресов,
сеть только при miss.
"""

from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.gsm.geo import GeoPoint, haversine_km, point_to_segment_km
from scripts.build_gsm_routes_map import (
    DEFAULT_GEO_CACHE,
    NOMINATIM_RATE_LIMIT_SEC,
    cache_entry_has_coords,
    coerce_cache_entry_coords,
    geocode_addresses,
    load_address_cache,
)
from scripts.build_gsm_trip_pool import normalize_address

DEFAULT_DB = PROJECT_ROOT / "plita.db"
DEFAULT_THRESHOLD_KM = 15.0
DEDUPE_NEARBY_KM = 0.2
PALISADE_VEHICLE_ID = 1
PALISADE_STATION_ID = 1

GeocodeFn = Callable[[str], tuple[float, float, str] | None]


@dataclass(frozen=True, slots=True)
class StationRow:
    id: int
    address: str
    lat: float | None
    lon: float | None

    @property
    def point(self) -> GeoPoint | None:
        if self.lat is None or self.lon is None:
            return None
        return GeoPoint(self.lat, self.lon)


@dataclass(frozen=True, slots=True)
class RouteRow:
    id: int
    vehicle_id: int
    addr_a: str
    addr_b: str
    km: int
    frequency: int
    typical_station_ids: str | None

    @property
    def addr_a_norm(self) -> str:
        return normalize_address(self.addr_a)

    @property
    def addr_b_norm(self) -> str:
        return normalize_address(self.addr_b)


@dataclass(frozen=True, slots=True)
class WaybillLeg:
    addr_from: str
    addr_to: str
    km: float
    route_id: int | None = None


@dataclass
class LinkReport:
    routes_total: int = 0
    routes_filled: int = 0
    routes_updated: int = 0
    history_bindings: int = 0
    geo_bindings: int = 0
    stations_skipped_no_coords: int = 0
    palisade_with_station_1: int = 0
    endpoints_from_cache: int = 0
    endpoints_from_network: int = 0
    endpoints_missing: int = 0

    def format(self) -> str:
        pct = (
            100.0 * self.routes_filled / self.routes_total
            if self.routes_total
            else 0.0
        )
        return (
            f"typical_station_ids: {self.routes_filled}/{self.routes_total} "
            f"маршрутов заполнено ({pct:.1f}%)\n"
            f"  обновлено: {self.routes_updated}\n"
            f"  из истории: {self.history_bindings}\n"
            f"  по географии: {self.geo_bindings}\n"
            f"  пропущено станций без координат: "
            f"{self.stations_skipped_no_coords}\n"
            f"  Palisade с station id=1: {self.palisade_with_station_1}\n"
            f"  концы маршрутов: кэш={self.endpoints_from_cache} "
            f"сеть={self.endpoints_from_network} нет={self.endpoints_missing}"
        )


def parse_typical_station_ids(raw: str | None) -> list[int]:
    """JSON-список id; NULL / '[]' / мусор → пустой список."""
    if raw is None:
        return []
    text = str(raw).strip()
    if text in ("", "[]", "null"):
        return []
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    ids: list[int] = []
    for item in data:
        try:
            ids.append(int(item))
        except (TypeError, ValueError):
            continue
    return ids


def merge_station_ids(existing: Iterable[int], incoming: Iterable[int]) -> list[int]:
    """Объединить списки: существующие id не удаляются."""
    return sorted(set(existing) | set(incoming))


def dedupe_nearby_stations(
    station_ids: Iterable[int],
    stations: dict[int, StationRow],
    *,
    radius_km: float = DEDUPE_NEARBY_KM,
) -> list[int]:
    """Оставить меньший id среди станций ближе ``radius_km``."""
    kept: list[int] = []
    for sid in sorted(set(station_ids)):
        station = stations.get(sid)
        point = station.point if station is not None else None
        if point is None:
            kept.append(sid)
            continue
        if any(
            _station_point(stations.get(kid)) is not None
            and haversine_km(point, _station_point(stations[kid])) < radius_km
            for kid in kept
        ):
            continue
        kept.append(sid)
    return kept


def _station_point(station: StationRow | None) -> GeoPoint | None:
    return None if station is None else station.point


def _connect(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def load_stations(db_path: Path) -> dict[int, StationRow]:
    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT id, address, lat, lon FROM gsm_station"
        ).fetchall()
    return {
        int(row["id"]): StationRow(
            id=int(row["id"]),
            address=str(row["address"] or ""),
            lat=row["lat"],
            lon=row["lon"],
        )
        for row in rows
    }


def load_routes(db_path: Path) -> list[RouteRow]:
    with _connect(db_path) as conn:
        rows = conn.execute(
            """
            SELECT id, vehicle_id, addr_a, addr_b, km, frequency,
                   typical_station_ids
            FROM gsm_route
            """
        ).fetchall()
    return [
        RouteRow(
            id=int(row["id"]),
            vehicle_id=int(row["vehicle_id"]),
            addr_a=str(row["addr_a"] or ""),
            addr_b=str(row["addr_b"] or ""),
            km=int(row["km"]),
            frequency=int(row["frequency"] or 1),
            typical_station_ids=row["typical_station_ids"],
        )
        for row in rows
    ]


def _parse_legs(route_json: str) -> list[WaybillLeg]:
    try:
        payload = json.loads(route_json or "[]")
    except json.JSONDecodeError:
        return []
    if not isinstance(payload, list):
        return []
    legs: list[WaybillLeg] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        addr_from = str(item.get("addr_from") or item.get("from") or "")
        addr_to = str(item.get("addr_to") or item.get("to") or "")
        if not addr_from or not addr_to:
            continue
        raw_km = item.get("km")
        try:
            km = float(raw_km) if raw_km is not None else 0.0
        except (TypeError, ValueError):
            km = 0.0
        raw_rid = item.get("route_id")
        route_id: int | None
        try:
            route_id = int(raw_rid) if raw_rid is not None else None
        except (TypeError, ValueError):
            route_id = None
        legs.append(
            WaybillLeg(
                addr_from=addr_from,
                addr_to=addr_to,
                km=km,
                route_id=route_id,
            )
        )
    return legs


def match_leg_to_route(leg: WaybillLeg, routes: list[RouteRow]) -> RouteRow | None:
    """Точный route_id, иначе vehicle-набор: нормализованные A/B + ближайший km."""
    if leg.route_id is not None:
        found = next((r for r in routes if r.id == leg.route_id), None)
        if found is not None:
            return found
    addr_a = normalize_address(leg.addr_from)
    addr_b = normalize_address(leg.addr_to)
    candidates = [
        r for r in routes if r.addr_a_norm == addr_a and r.addr_b_norm == addr_b
    ]
    if not candidates:
        return None
    return min(candidates, key=lambda r: (abs(r.km - leg.km), -r.frequency, r.id))


def collect_history_links(
    db_path: Path,
    routes: list[RouteRow],
) -> dict[int, set[int]]:
    """station_id заправки дня → маршруты плеч imported-ПЛ той же машины."""
    by_vehicle: dict[int, list[RouteRow]] = defaultdict(list)
    for route in routes:
        by_vehicle[route.vehicle_id].append(route)

    links: dict[int, set[int]] = defaultdict(set)
    with _connect(db_path) as conn:
        waybills = conn.execute(
            """
            SELECT id, vehicle_id, date, route_json
            FROM gsm_waybill
            WHERE source = 'imported'
            """
        ).fetchall()
        fuels = conn.execute(
            """
            SELECT c.vehicle_id AS vehicle_id,
                   substr(t.ts, 1, 10) AS day,
                   t.station_id AS station_id
            FROM gsm_transaction t
            JOIN gsm_fuel_card c ON c.id = t.card_id
            WHERE t.service_type = 'fuel'
              AND t.station_id IS NOT NULL
            """
        ).fetchall()

    day_stations: dict[tuple[int, str], set[int]] = defaultdict(set)
    for row in fuels:
        day_stations[(int(row["vehicle_id"]), str(row["day"]))].add(
            int(row["station_id"])
        )

    for wb in waybills:
        vehicle_id = int(wb["vehicle_id"])
        stations = day_stations.get((vehicle_id, str(wb["date"])), set())
        if not stations:
            continue
        for leg in _parse_legs(str(wb["route_json"] or "")):
            matched = match_leg_to_route(leg, by_vehicle.get(vehicle_id, []))
            if matched is None:
                continue
            links[matched.id].update(stations)
    return links


def _coords_from_cache(
    addresses: dict[str, Any], raw: str
) -> GeoPoint | None:
    entry = addresses.get(normalize_address(raw))
    if not isinstance(entry, dict):
        return None
    coerce_cache_entry_coords(entry)
    if not cache_entry_has_coords(entry):
        return None
    return GeoPoint(float(entry["lat"]), float(entry["lon"]))


def geocode_route_endpoints(
    routes: list[RouteRow],
    cache_path: Path,
    *,
    offline: bool,
    geocode_fn: GeocodeFn | None,
    sleep_sec: float,
) -> tuple[dict[str, Any], Any]:
    """Геокодировать уникальные A/B: кэш, сеть только на miss."""
    display_map: dict[str, str] = {}
    for route in routes:
        for raw in (route.addr_a, route.addr_b):
            norm = normalize_address(raw)
            if norm and norm not in display_map:
                display_map[norm] = raw
    stats = geocode_addresses(
        list(display_map.keys()),
        display_map,
        cache_path,
        offline=offline,
        sleep_sec=sleep_sec,
        geocode_fn=geocode_fn,
    )
    cache = load_address_cache(cache_path)
    return cache.get("addresses") or {}, stats


def collect_geo_links(
    routes: list[RouteRow],
    stations: dict[int, StationRow],
    addresses: dict[str, Any],
    *,
    threshold_km: float,
) -> tuple[dict[int, set[int]], int]:
    """Станции с отклонением от отрезка A–B < threshold. Без координат — skip."""
    skipped = 0
    usable: list[StationRow] = []
    for station in stations.values():
        if station.point is None:
            skipped += 1
            print(
                f"skip station id={station.id}: нет координат ({station.address})",
                flush=True,
            )
            continue
        usable.append(station)

    links: dict[int, set[int]] = defaultdict(set)
    for route in routes:
        point_a = _coords_from_cache(addresses, route.addr_a)
        point_b = _coords_from_cache(addresses, route.addr_b)
        if point_a is None or point_b is None:
            continue
        for station in usable:
            point = station.point
            if point is None:
                continue
            if point_to_segment_km(point, point_a, point_b) < threshold_km:
                links[route.id].add(station.id)
    return links, skipped


def _write_typical_ids(db_path: Path, route_id: int, ids: list[int]) -> None:
    payload = json.dumps(ids, ensure_ascii=False)
    with _connect(db_path) as conn:
        conn.execute(
            "UPDATE gsm_route SET typical_station_ids = ? WHERE id = ?",
            (payload, route_id),
        )


def link_route_stations(
    db_path: str | Path,
    cache_path: str | Path,
    *,
    offline: bool = False,
    threshold_km: float = DEFAULT_THRESHOLD_KM,
    geocode_fn: GeocodeFn | None = None,
    sleep_sec: float = NOMINATIM_RATE_LIMIT_SEC,
) -> LinkReport:
    """Заполнить typical_station_ids (история + география, merge без затирания)."""
    db = Path(db_path)
    cache = Path(cache_path)
    routes = load_routes(db)
    stations = load_stations(db)
    history = collect_history_links(db, routes)
    addresses, geo_stats = geocode_route_endpoints(
        routes,
        cache,
        offline=offline,
        geocode_fn=geocode_fn,
        sleep_sec=sleep_sec,
    )
    geo, skipped = collect_geo_links(
        routes, stations, addresses, threshold_km=threshold_km
    )

    report = LinkReport(
        routes_total=len(routes),
        history_bindings=sum(len(ids) for ids in history.values()),
        geo_bindings=sum(len(ids) for ids in geo.values()),
        stations_skipped_no_coords=skipped,
        endpoints_from_cache=int(getattr(geo_stats, "from_cache", 0) or 0),
        endpoints_from_network=int(getattr(geo_stats, "geocoded", 0) or 0),
        endpoints_missing=len(getattr(geo_stats, "missing", ()) or ()),
    )

    for route in routes:
        existing = parse_typical_station_ids(route.typical_station_ids)
        incoming = dedupe_nearby_stations(
            set(history.get(route.id, set())) | set(geo.get(route.id, set())),
            stations,
        )
        merged = merge_station_ids(existing, incoming)
        if merged and merged != existing:
            _write_typical_ids(db, route.id, merged)
            report.routes_updated += 1

    with _connect(db) as conn:
        filled_rows = conn.execute(
            """
            SELECT vehicle_id, typical_station_ids
            FROM gsm_route
            WHERE typical_station_ids IS NOT NULL
              AND typical_station_ids != '[]'
            """
        ).fetchall()
    report.routes_filled = len(filled_rows)
    report.palisade_with_station_1 = sum(
        1
        for row in filled_rows
        if int(row["vehicle_id"]) == PALISADE_VEHICLE_ID
        and PALISADE_STATION_ID
        in parse_typical_station_ids(row["typical_station_ids"])
    )
    return report


def parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Заполнение gsm_route.typical_station_ids "
            "(история imported-ПЛ + география A–B)."
        )
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
        help="Только кэш адресов, без запросов к Nominatim.",
    )
    parser.add_argument(
        "--threshold-km",
        type=float,
        default=DEFAULT_THRESHOLD_KM,
        help=f"Порог point-to-segment, км (по умолчанию: {DEFAULT_THRESHOLD_KM})",
    )
    return parser.parse_args(list(argv) if argv is not None else None)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    report = link_route_stations(
        args.db,
        args.cache,
        offline=args.offline,
        threshold_km=args.threshold_km,
    )
    print(report.format())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
