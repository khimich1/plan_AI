"""Unit-тесты чтения routes_ab, геокода, OSRM, HTML-карты (T1–T5)."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import requests
from openpyxl import Workbook

from scripts.build_gsm_routes_map import (
    VEHICLE_COLORS,
    address_display_map,
    build_routes_geojson,
    cache_entry_has_coords,
    collect_unique_addresses,
    distance_point_to_linestring_km,
    geocode_addresses,
    haversine_km,
    known_addresses_from_cache,
    load_address_cache,
    load_routes_ab,
    load_routes_geojson,
    nearest_routes,
    nominatim_geocode,
    nominatim_query,
    nominatim_query_variants,
    osrm_route,
    point_to_segment_km,
    route_osrm_cache_key,
    run,
    save_address_cache,
    save_routes_geojson,
    simplify_address_for_geocode,
    simplify_linestring_coords,
    write_map_html,
)
from scripts.build_gsm_trip_pool import normalize_address


def _write_routes_ab(path: Path, rows: list[tuple]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "routes_ab"
    headers = (
        "машина",
        "марка",
        "гос_номер",
        "адрес_A",
        "адрес_B",
        "адрес_A_норм",
        "адрес_B_норм",
        "км",
        "частота",
        "типичное_время_выезда",
        "водители",
        "топливо",
    )
    ws.append(list(headers))
    for row in rows:
        ws.append(list(row))
    wb.save(path)


def _sample_rows() -> list[tuple]:
    return [
        (
            "Hyundai Palisade",
            "Hyundai",
            "О 521 УХ 44",
            "г.Кострома, ул.Кузнецкая, д.18Б",
            "г.Ярославль, пр-д Домостроителей, д.1, стр.3",
            "",
            "",
            95,
            39,
            "07:10",
            "Иванов",
            "АИ-95",
        ),
        (
            "Geely Monjaro",
            "Geely",
            "О 165 ХУ 44",
            "г Кострома ул Кузнецкая дом 18Б",
            "г.Ярославль, пр-д Домостроителей, д.1, стр.3",
            "",
            "",
            96,
            2,
            "08:00",
            "Петров",
            "АИ-95",
        ),
        (
            "Skip",
            "",
            "",
            "",
            "",
            "",
            "",
            0,
            0,
            "",
            "",
            "",
        ),
    ]


def test_load_routes_ab_and_unique_addresses(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows())

    routes = load_routes_ab(xlsx)
    assert len(routes) == 2
    assert routes[0].vehicle == "Hyundai Palisade"
    assert routes[0].km == 95.0
    assert routes[0].frequency == 39
    assert routes[0].addr_a_norm == normalize_address(
        "г.Кострома, ул.Кузнецкая, д.18Б"
    )
    assert routes[1].addr_a_norm == routes[0].addr_a_norm

    addresses = collect_unique_addresses(routes)
    assert len(addresses) == 2
    assert normalize_address("г.Кострома, ул.Кузнецкая, д.18Б") in addresses
    assert (
        normalize_address("г.Ярославль, пр-д Домостроителей, д.1, стр.3")
        in addresses
    )


def test_simplify_address_for_geocode_strips_labels_and_org() -> None:
    assert simplify_address_for_geocode(
        'г.Кострома, ул. Кузнецкая, д.18Б, ООО "Ромашка", Россия'
    ) == "Кострома, Кузнецкая 18Б"
    assert simplify_address_for_geocode(
        "г.Ярославль, пр-д Домостроителей, д.1, стр.3"
    ) == "Ярославль, проезд Домостроителей 1"
    assert simplify_address_for_geocode(
        'г.Владимир, пр-кт Ленина, д.71Б, ООО "Строимгрупп"'
    ) == "Владимир, проспект Ленина 71Б"
    assert simplify_address_for_geocode(
        'Владимирская обл., г.Кольчугино, ул.Ленина, д.32, ООО "КольчугЭнергоСтрой М"'
    ) == "Владимирская область, Кольчугино, Ленина 32"
    assert simplify_address_for_geocode(
        'г.Владимир, пр-кт Октябрьский, д.27, СК "Гамма-Строй"'
    ) == "Владимир, проспект Октябрьский 27"
    assert simplify_address_for_geocode("Москва, Россия") == "Москва"
    assert simplify_address_for_geocode("Moscow, Russia") == "Moscow"


def test_nominatim_query_prefers_simplified_form() -> None:
    assert nominatim_query(
        'г.Кострома, ул. Кузнецкая, д.18Б, ООО "Х", Россия'
    ) == "Кострома, Кузнецкая 18Б"
    variants = nominatim_query_variants(
        'Владимирская обл., г.Кольчугино, ул.Ленина, д.32, ООО "Х"'
    )
    assert variants[0] == "Владимирская область, Кольчугино, Ленина 32"
    assert "Владимирская область, Кольчугино, Ленина 32, Россия" in variants
    assert "Кольчугино, Ленина 32" in variants


def test_load_save_address_cache_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / "geo_cache" / "addresses.json"
    cache = {
        "version": 1,
        "addresses": {
            "кострома": {
                "lat": 57.7,
                "lon": 40.9,
                "display": "Кострома",
                "query": "Кострома, Россия",
                "source": "manual",
                "error": None,
            }
        },
    }
    save_address_cache(path, cache)
    loaded = load_address_cache(path)
    assert loaded["version"] == 1
    assert loaded["addresses"]["кострома"]["lat"] == 57.7
    assert cache_entry_has_coords(loaded["addresses"]["кострома"])
    assert load_address_cache(tmp_path / "missing.json")["addresses"] == {}


def test_geocode_addresses_uses_cache_and_mocks_http(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows())
    routes = load_routes_ab(xlsx)
    norms = collect_unique_addresses(routes)
    display = address_display_map(routes)
    cache_path = tmp_path / "addresses.json"

    a_norm = normalize_address("г.Кострома, ул.Кузнецкая, д.18Б")
    b_norm = normalize_address("г.Ярославль, пр-д Домостроителей, д.1, стр.3")
    save_address_cache(
        cache_path,
        {
            "version": 1,
            "addresses": {
                a_norm: {
                    "lat": 57.766,
                    "lon": 40.927,
                    "display": "Кузнецкая",
                    "query": "…",
                    "source": "manual",
                    "error": None,
                }
            },
        },
    )

    calls: list[str] = []

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        calls.append(query)
        return 57.63, 39.87, "Ярославль"

    sleeps: list[float] = []

    stats = geocode_addresses(
        norms,
        display,
        cache_path,
        offline=False,
        sleep_sec=1.1,
        sleep_fn=sleeps.append,
        geocode_fn=fake_geocode,
    )

    assert stats.from_cache == 1
    assert stats.geocoded == 1
    assert stats.failed == 0
    assert stats.missing == ()
    assert len(calls) == 1
    assert calls[0] == "Ярославль, проезд Домостроителей 1"
    # Первый сетевой вызов без sleep; второй адрес — единственный запрос → sleep нет
    assert sleeps == []

    cached = load_address_cache(cache_path)
    assert cache_entry_has_coords(cached["addresses"][a_norm])
    assert cached["addresses"][b_norm]["lat"] == 57.63
    assert cached["addresses"][b_norm]["source"] == "nominatim"
    assert cached["addresses"][b_norm]["error"] is None


def test_geocode_rate_limit_sleep_between_requests(tmp_path: Path) -> None:
    cache_path = tmp_path / "addresses.json"
    norms = ["addr_a", "addr_b"]
    display = {"addr_a": "A", "addr_b": "B"}
    sleeps: list[float] = []
    n = {"i": 0}

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        n["i"] += 1
        return float(n["i"]), float(n["i"]), query

    stats = geocode_addresses(
        norms,
        display,
        cache_path,
        sleep_sec=1.1,
        sleep_fn=sleeps.append,
        geocode_fn=fake_geocode,
    )
    assert stats.geocoded == 2
    assert sleeps == [1.1]


def test_geocode_offline_no_network(tmp_path: Path) -> None:
    cache_path = tmp_path / "addresses.json"
    save_address_cache(
        cache_path,
        {
            "version": 1,
            "addresses": {
                "known": {
                    "lat": 1.0,
                    "lon": 2.0,
                    "display": "known",
                    "query": "known",
                    "source": "manual",
                    "error": None,
                },
                "failed_cached": {
                    "lat": None,
                    "lon": None,
                    "display": "failed",
                    "query": "failed",
                    "source": "nominatim",
                    "error": "not_found",
                },
            },
        },
    )

    def boom(_query: str) -> tuple[float, float, str] | None:
        raise AssertionError("network must not be used in offline mode")

    stats = geocode_addresses(
        ["known", "failed_cached", "missing_one"],
        {
            "known": "known",
            "failed_cached": "failed",
            "missing_one": "Missing",
        },
        cache_path,
        offline=True,
        geocode_fn=boom,
    )
    assert stats.from_cache == 2
    assert stats.geocoded == 0
    assert stats.failed == 0
    assert stats.missing == ("missing_one",)


def test_geocode_failed_not_found(tmp_path: Path) -> None:
    cache_path = tmp_path / "addresses.json"
    stats = geocode_addresses(
        ["nowhere"],
        {"nowhere": "нигде"},
        cache_path,
        sleep_sec=0,
        geocode_fn=lambda _q: None,
    )
    assert stats.failed == 1
    assert stats.geocoded == 0
    entry = load_address_cache(cache_path)["addresses"]["nowhere"]
    assert entry["lat"] is None
    assert entry["error"] == "not_found"


def test_geocode_tries_query_variants_until_hit(tmp_path: Path) -> None:
    cache_path = tmp_path / "addresses.json"
    calls: list[str] = []
    sleeps: list[float] = []

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        calls.append(query)
        if "Россия" in query:
            return 57.7, 40.9, query
        return None

    stats = geocode_addresses(
        ["a"],
        {"a": "г.Кострома, ул. Кузнецкая, д.18Б"},
        cache_path,
        sleep_sec=1.1,
        sleep_fn=sleeps.append,
        geocode_fn=fake_geocode,
    )
    assert stats.geocoded == 1
    assert calls[0] == "Кострома, Кузнецкая 18Б"
    assert calls[1] == "Кострома, Кузнецкая 18Б, Россия"
    assert sleeps == [1.1]
    entry = load_address_cache(cache_path)["addresses"]["a"]
    assert entry["query"] == "Кострома, Кузнецкая 18Б, Россия"
    assert entry["error"] is None


def test_geocode_negative_cache_no_second_network(tmp_path: Path) -> None:
    """Кэшированный not_found не должен повторно бить в Nominatim онлайн."""
    cache_path = tmp_path / "addresses.json"
    calls: list[str] = []

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        calls.append(query)
        return None

    first = geocode_addresses(
        ["nowhere"],
        {"nowhere": "нигде"},
        cache_path,
        sleep_sec=0,
        geocode_fn=fake_geocode,
    )
    assert first.failed == 1
    # Все варианты запроса исчерпаны (упрощённый + «, Россия»)
    assert calls == ["нигде", "нигде, Россия"]

    def boom(_query: str) -> tuple[float, float, str] | None:
        raise AssertionError("negative cache must not re-query Nominatim")

    second = geocode_addresses(
        ["nowhere"],
        {"nowhere": "нигде"},
        cache_path,
        offline=False,
        sleep_sec=0,
        geocode_fn=boom,
    )
    assert second.from_cache == 1
    assert second.geocoded == 0
    assert second.failed == 0
    assert second.missing == ()
    assert calls == ["нигде", "нигде, Россия"]


def test_geocode_force_and_retry_requery_negative_cache(tmp_path: Path) -> None:
    cache_path = tmp_path / "addresses.json"
    save_address_cache(
        cache_path,
        {
            "version": 1,
            "addresses": {
                "nowhere": {
                    "lat": None,
                    "lon": None,
                    "display": "нигде",
                    "query": "нигде, Россия",
                    "source": "nominatim",
                    "error": "not_found",
                }
            },
        },
    )
    calls: list[str] = []

    def fake_geocode(query: str) -> tuple[float, float, str] | None:
        calls.append(query)
        return 1.0, 2.0, query

    forced = geocode_addresses(
        ["nowhere"],
        {"nowhere": "нигде"},
        cache_path,
        force_geocode=True,
        sleep_sec=0,
        geocode_fn=fake_geocode,
    )
    assert forced.geocoded == 1
    assert len(calls) == 1

    save_address_cache(
        cache_path,
        {
            "version": 1,
            "addresses": {
                "nowhere": {
                    "lat": None,
                    "lon": None,
                    "display": "нигде",
                    "query": "нигде, Россия",
                    "source": "nominatim",
                    "error": "not_found",
                    "retry": True,
                }
            },
        },
    )
    retried = geocode_addresses(
        ["nowhere"],
        {"nowhere": "нигде"},
        cache_path,
        sleep_sec=0,
        geocode_fn=fake_geocode,
    )
    assert retried.geocoded == 1
    assert len(calls) == 2


def test_cache_entry_has_coords_coerces_numeric_strings() -> None:
    assert cache_entry_has_coords({"lat": "57.7", "lon": "40.9"})
    assert not cache_entry_has_coords({"lat": None, "lon": None})


def test_nominatim_geocode_http_mock() -> None:
    mock_resp = MagicMock()
    mock_resp.raise_for_status = MagicMock()
    mock_resp.json.return_value = [
        {
            "lat": "55.75",
            "lon": "37.61",
            "display_name": "Москва",
        }
    ]
    session = MagicMock()
    session.get.return_value = mock_resp

    result = nominatim_geocode("Москва, Россия", session=session)
    assert result == (55.75, 37.61, "Москва")
    kwargs = session.get.call_args.kwargs
    assert kwargs["headers"]["User-Agent"].startswith("ShishovGSM/")
    assert kwargs["headers"]["Accept-Language"] == "ru"
    assert kwargs["params"]["countrycodes"] == "ru"
    assert kwargs["params"]["q"] == "Москва, Россия"


def test_nominatim_geocode_empty_list() -> None:
    mock_resp = MagicMock()
    mock_resp.raise_for_status = MagicMock()
    mock_resp.json.return_value = []
    session = MagicMock()
    session.get.return_value = mock_resp
    assert nominatim_geocode("xxx", session=session) is None


def test_run_offline_reports_missing(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows())
    cache_path = tmp_path / "addresses.json"
    out = tmp_path / "map.html"

    def boom(_query: str) -> tuple[float, float, str] | None:
        raise AssertionError("no network")

    report = run(
        xlsx,
        out,
        offline=True,
        geo_cache_path=cache_path,
        geocode_fn=boom,
    )

    assert report["from_cache"] == 0
    assert report["geocoded"] == 0
    assert len(report["missing"]) == 2


def test_geocode_http_error_counts_failed(tmp_path: Path) -> None:
    cache_path = tmp_path / "addresses.json"

    def raise_http(_query: str) -> tuple[float, float, str] | None:
        raise requests.HTTPError("503")

    stats = geocode_addresses(
        ["x"],
        {"x": "X"},
        cache_path,
        sleep_sec=0,
        geocode_fn=raise_http,
    )
    assert stats.failed == 1
    assert "503" in load_address_cache(cache_path)["addresses"]["x"]["error"]


def _addr_cache_for_sample(routes: list) -> dict:
    """Минимальный addresses.json с координатами A/B первого маршрута."""
    a_norm = routes[0].addr_a_norm
    b_norm = routes[0].addr_b_norm
    return {
        "version": 1,
        "addresses": {
            a_norm: {
                "lat": 57.766,
                "lon": 40.927,
                "display": "A",
                "query": "A",
                "source": "manual",
                "error": None,
            },
            b_norm: {
                "lat": 57.63,
                "lon": 39.87,
                "display": "B",
                "query": "B",
                "source": "manual",
                "error": None,
            },
        },
    }


def _no_osrm(*_a: float) -> tuple[list[list[float]], float] | None:
    raise AssertionError("OSRM must not be called")


def test_route_osrm_cache_key_format() -> None:
    key = route_osrm_cache_key(40.927001, 57.766009, 39.870004, 57.630002)
    assert key == "40.92700,57.76601->39.87000,57.63000"


def test_osrm_route_http_mock() -> None:
    mock_resp = MagicMock()
    mock_resp.raise_for_status = MagicMock()
    mock_resp.json.return_value = {
        "code": "Ok",
        "routes": [
            {
                "distance": 12345.6,
                "geometry": {
                    "type": "LineString",
                    "coordinates": [[40.9, 57.7], [40.0, 57.6], [39.87, 57.63]],
                },
            }
        ],
    }
    session = MagicMock()
    session.get.return_value = mock_resp

    result = osrm_route(40.927, 57.766, 39.87, 57.63, session=session)
    assert result is not None
    coords, dist = result
    assert dist == 12345.6
    assert len(coords) == 3
    assert coords[0] == [40.9, 57.7]
    url = session.get.call_args.args[0]
    assert "40.927,57.766;39.87,57.63" in url
    assert session.get.call_args.kwargs["params"]["overview"] == "simplified"
    assert session.get.call_args.kwargs["params"]["geometries"] == "geojson"


def test_simplify_linestring_coords_keeps_ends_and_reduces() -> None:
    # Прямая линия + «шум» в пределах допуска — должна сжаться до 2 точек.
    coords = [[0.0, 0.0], [0.00001, 0.0], [0.00002, 0.0], [0.001, 0.0]]
    out = simplify_linestring_coords(coords, tolerance_m=50.0)
    assert out[0] == [0.0, 0.0]
    assert out[-1] == [0.001, 0.0]
    assert len(out) < len(coords)
    assert len(out) <= 200


def test_build_routes_geojson_mocks_osrm_and_caches(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows())
    routes = load_routes_ab(xlsx)
    address_cache = _addr_cache_for_sample(routes)
    routes_path = tmp_path / "routes.geojson"
    calls: list[tuple[float, float, float, float]] = []
    sleeps: list[float] = []

    def fake_osrm(
        lon1: float, lat1: float, lon2: float, lat2: float
    ) -> tuple[list[list[float]], float] | None:
        calls.append((lon1, lat1, lon2, lat2))
        return [[lon1, lat1], [lon2, lat2]], 95000.0

    stats = build_routes_geojson(
        routes,
        address_cache,
        routes_path,
        offline=False,
        sleep_sec=0.2,
        sleep_fn=sleeps.append,
        osrm_fn=fake_osrm,
    )

    # Оба маршрута — одни и те же координаты A→B → один OSRM-вызов, второй из by_geom
    assert stats.routed == 1
    assert stats.from_cache == 1
    assert stats.features == 2
    assert stats.skipped_missing_coords == 0
    assert len(calls) == 1
    assert sleeps == []  # только один сетевой вызов

    cached = load_routes_geojson(routes_path)
    assert cached["type"] == "FeatureCollection"
    assert len(cached["features"]) == 2
    props0 = cached["features"][0]["properties"]
    assert props0["машина"] == "Hyundai Palisade"
    assert props0["osrm_distance_m"] == 95000.0
    assert "cache_key" in props0
    assert cached["features"][0]["geometry"]["type"] == "LineString"

    # Повтор: без сети, полный from_cache
    stats2 = build_routes_geojson(
        routes,
        address_cache,
        routes_path,
        offline=False,
        sleep_sec=0.2,
        sleep_fn=sleeps.append,
        osrm_fn=_no_osrm,
    )
    assert stats2.routed == 0
    assert stats2.from_cache == 2
    assert stats2.features == 2
    assert len(calls) == 1


def test_build_routes_osrm_rate_limit_sleep(tmp_path: Path) -> None:
    """Две пары с разными координатами → sleep между OSRM-вызовами."""
    xlsx = tmp_path / "pool.xlsx"
    rows = [
        (
            "Car1",
            "",
            "",
            "Addr A1",
            "Addr B1",
            "",
            "",
            10,
            1,
            "",
            "",
            "",
        ),
        (
            "Car2",
            "",
            "",
            "Addr A2",
            "Addr B2",
            "",
            "",
            20,
            1,
            "",
            "",
            "",
        ),
    ]
    _write_routes_ab(xlsx, rows)
    routes = load_routes_ab(xlsx)
    address_cache = {
        "version": 1,
        "addresses": {
            routes[0].addr_a_norm: {
                "lat": 57.0,
                "lon": 40.0,
                "display": "A1",
                "query": "A1",
                "source": "manual",
                "error": None,
            },
            routes[0].addr_b_norm: {
                "lat": 57.1,
                "lon": 40.1,
                "display": "B1",
                "query": "B1",
                "source": "manual",
                "error": None,
            },
            routes[1].addr_a_norm: {
                "lat": 58.0,
                "lon": 41.0,
                "display": "A2",
                "query": "A2",
                "source": "manual",
                "error": None,
            },
            routes[1].addr_b_norm: {
                "lat": 58.1,
                "lon": 41.1,
                "display": "B2",
                "query": "B2",
                "source": "manual",
                "error": None,
            },
        },
    }
    sleeps: list[float] = []
    n = {"i": 0}

    def fake_osrm(
        lon1: float, lat1: float, lon2: float, lat2: float
    ) -> tuple[list[list[float]], float] | None:
        n["i"] += 1
        return [[lon1, lat1], [lon2, lat2]], float(n["i"] * 1000)

    stats = build_routes_geojson(
        routes,
        address_cache,
        tmp_path / "routes.geojson",
        sleep_sec=0.2,
        sleep_fn=sleeps.append,
        osrm_fn=fake_osrm,
    )
    assert stats.routed == 2
    assert sleeps == [0.2]


def test_build_routes_offline_no_network(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows())
    routes = load_routes_ab(xlsx)
    address_cache = _addr_cache_for_sample(routes)
    routes_path = tmp_path / "routes.geojson"

    stats = build_routes_geojson(
        routes,
        address_cache,
        routes_path,
        offline=True,
        osrm_fn=_no_osrm,
    )
    assert stats.routed == 0
    assert stats.from_cache == 0
    assert stats.skipped_offline == 2
    assert stats.features == 0
    assert all("offline" in s for s in stats.skips)
    assert not routes_path.is_file()


def test_build_routes_identity_hit_requires_matching_cache_key(
    tmp_path: Path,
) -> None:
    """by_route identity alone must not reuse geometry after geocode coords change."""
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows()[:1])
    routes = load_routes_ab(xlsx)
    a_norm = routes[0].addr_a_norm
    b_norm = routes[0].addr_b_norm
    # Адреса те же, координаты новые (геокод обновился)
    address_cache = {
        "version": 1,
        "addresses": {
            a_norm: {
                "lat": 57.800,
                "lon": 41.000,
                "display": "A",
                "query": "A",
                "source": "manual",
                "error": None,
            },
            b_norm: {
                "lat": 57.700,
                "lon": 39.900,
                "display": "B",
                "query": "B",
                "source": "manual",
                "error": None,
            },
        },
    }
    routes_path = tmp_path / "routes.geojson"
    stale_key = route_osrm_cache_key(40.927, 57.766, 39.87, 57.63)
    current_key = route_osrm_cache_key(41.000, 57.800, 39.900, 57.700)
    assert stale_key != current_key
    stale_coords = [[40.927, 57.766], [39.87, 57.63]]
    save_routes_geojson(
        routes_path,
        {
            "type": "FeatureCollection",
            "features": [
                {
                    "type": "Feature",
                    "properties": {
                        "cache_key": stale_key,
                        "машина": routes[0].vehicle,
                        "адрес_A": routes[0].addr_a,
                        "адрес_B": routes[0].addr_b,
                        "км": routes[0].km,
                        "частота": routes[0].frequency,
                        "address_a_norm": a_norm,
                        "address_b_norm": b_norm,
                        "osrm_distance_m": 1.0,
                    },
                    "geometry": {
                        "type": "LineString",
                        "coordinates": stale_coords,
                    },
                }
            ],
        },
    )

    calls: list[tuple[float, float, float, float]] = []
    fresh_coords = [[41.0, 57.8], [39.9, 57.7]]

    def fake_osrm(
        lon1: float, lat1: float, lon2: float, lat2: float
    ) -> tuple[list[list[float]], float] | None:
        calls.append((lon1, lat1, lon2, lat2))
        return fresh_coords, 99000.0

    stats = build_routes_geojson(
        routes,
        address_cache,
        routes_path,
        offline=False,
        sleep_sec=0,
        osrm_fn=fake_osrm,
    )
    assert stats.from_cache == 0
    assert stats.routed == 1
    assert len(calls) == 1
    assert calls[0] == (41.0, 57.8, 39.9, 57.7)

    features = load_routes_geojson(routes_path)["features"]
    # Новый трек с актуальным cache_key; stale identity не должен «закрыть» OSRM
    fresh = [
        f
        for f in features
        if (f.get("properties") or {}).get("cache_key") == current_key
    ]
    assert len(fresh) == 1
    assert fresh[0]["geometry"]["coordinates"] == fresh_coords
    assert fresh[0]["properties"]["osrm_distance_m"] == 99000.0


def test_build_routes_offline_uses_existing_cache(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows())
    routes = load_routes_ab(xlsx)
    address_cache = _addr_cache_for_sample(routes)
    routes_path = tmp_path / "routes.geojson"
    key = route_osrm_cache_key(40.927, 57.766, 39.87, 57.63)
    save_routes_geojson(
        routes_path,
        {
            "type": "FeatureCollection",
            "features": [
                {
                    "type": "Feature",
                    "properties": {
                        "cache_key": key,
                        "машина": routes[0].vehicle,
                        "адрес_A": routes[0].addr_a,
                        "адрес_B": routes[0].addr_b,
                        "км": routes[0].km,
                        "частота": routes[0].frequency,
                        "address_a_norm": routes[0].addr_a_norm,
                        "address_b_norm": routes[0].addr_b_norm,
                        "osrm_distance_m": 1.0,
                    },
                    "geometry": {
                        "type": "LineString",
                        "coordinates": [[40.927, 57.766], [39.87, 57.63]],
                    },
                }
            ],
        },
    )

    stats = build_routes_geojson(
        routes,
        address_cache,
        routes_path,
        offline=True,
        osrm_fn=_no_osrm,
    )
    # Первый — точное совпадение identity; второй — reuse geometry по cache_key
    assert stats.from_cache == 2
    assert stats.skipped_offline == 0
    assert stats.routed == 0
    assert stats.features == 2


def test_build_routes_skips_missing_coords(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows())
    routes = load_routes_ab(xlsx)
    address_cache = {"version": 1, "addresses": {}}
    stats = build_routes_geojson(
        routes,
        address_cache,
        tmp_path / "routes.geojson",
        offline=False,
        osrm_fn=_no_osrm,
    )
    assert stats.skipped_missing_coords == 2
    assert stats.routed == 0
    assert all("missing coords" in s for s in stats.skips)


def test_build_routes_osrm_error_logged(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows()[:1])
    routes = load_routes_ab(xlsx)
    address_cache = _addr_cache_for_sample(routes)

    def fail_osrm(
        *_a: float,
    ) -> tuple[list[list[float]], float] | None:
        raise requests.HTTPError("429")

    stats = build_routes_geojson(
        routes,
        address_cache,
        tmp_path / "routes.geojson",
        sleep_sec=0,
        osrm_fn=fail_osrm,
    )
    assert stats.skipped_osrm_error == 1
    assert any("429" in s for s in stats.skips)


def test_run_osrm_with_mocks(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows())
    cache_path = tmp_path / "addresses.json"
    routes_path = tmp_path / "routes.geojson"
    routes = load_routes_ab(xlsx)
    save_address_cache(cache_path, _addr_cache_for_sample(routes))

    def boom_geo(_q: str) -> tuple[float, float, str] | None:
        raise AssertionError("geocode cache should hit")

    calls: list[tuple] = []

    def fake_osrm(
        lon1: float, lat1: float, lon2: float, lat2: float
    ) -> tuple[list[list[float]], float] | None:
        calls.append((lon1, lat1, lon2, lat2))
        return [[lon1, lat1], [lon2, lat2]], 100.0

    report = run(
        xlsx,
        tmp_path / "map.html",
        offline=False,
        geo_cache_path=cache_path,
        routes_cache_path=routes_path,
        geocode_fn=boom_geo,
        osrm_fn=fake_osrm,
        sleep_sec=0,
        osrm_sleep_sec=0,
    )
    assert report["geocoded_addresses"] == 2
    assert report["osrm_routed"] == 1
    assert report["osrm_from_cache"] == 1
    assert report["osrm_features"] == 2
    assert routes_path.is_file()
    assert len(calls) == 1


def test_run_empty_address_cache_message(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows())
    cache_path = tmp_path / "addresses.json"
    routes_path = tmp_path / "routes.geojson"
    out = tmp_path / "map.html"

    report = run(
        xlsx,
        out,
        offline=True,
        geo_cache_path=cache_path,
        routes_cache_path=routes_path,
        geocode_fn=lambda _q: (_ for _ in ()).throw(AssertionError("offline")),
        osrm_fn=_no_osrm,
    )
    assert report["geocoded_addresses"] == 0
    assert report["osrm_features"] == 0
    assert any("addresses.json" in s for s in report["osrm_skips"])
    assert not routes_path.is_file()
    assert out.is_file()
    assert report["html_features"] == 0


def test_haversine_and_point_to_segment() -> None:
    # ~111 км на 1° широты
    d = haversine_km(40.0, 57.0, 40.0, 58.0)
    assert 110 < d < 112
    # Точка на середине горизонтального отрезка → ~0
    mid = point_to_segment_km((40.5, 57.0), (40.0, 57.0), (41.0, 57.0))
    assert mid < 0.05
    # Точка в стороне от отрезка
    side = point_to_segment_km((40.5, 57.1), (40.0, 57.0), (41.0, 57.0))
    assert 10 < side < 12


def test_distance_point_to_linestring_km() -> None:
    coords = [[40.0, 57.0], [40.5, 57.0], [41.0, 57.0]]
    assert distance_point_to_linestring_km((40.5, 57.0), coords) < 0.05
    assert distance_point_to_linestring_km((40.5, 57.0), [[40.0, 57.0]]) is None


def test_nearest_routes_top3_by_distance() -> None:
    features = [
        {
            "type": "Feature",
            "properties": {
                "машина": "Far",
                "адрес_A": "A1",
                "адрес_B": "B1",
                "км": 10,
                "частота": 1,
            },
            # далеко на севере
            "geometry": {
                "type": "LineString",
                "coordinates": [[40.0, 58.5], [41.0, 58.5]],
            },
        },
        {
            "type": "Feature",
            "properties": {
                "машина": "Near",
                "адрес_A": "A2",
                "адрес_B": "B2",
                "км": 20,
                "частота": 2,
            },
            "geometry": {
                "type": "LineString",
                "coordinates": [[40.0, 57.0], [41.0, 57.0]],
            },
        },
        {
            "type": "Feature",
            "properties": {
                "машина": "Mid",
                "адрес_A": "A3",
                "адрес_B": "B3",
                "км": 30,
                "частота": 3,
            },
            "geometry": {
                "type": "LineString",
                "coordinates": [[40.0, 57.2], [41.0, 57.2]],
            },
        },
        {
            "type": "Feature",
            "properties": {
                "машина": "Closer",
                "адрес_A": "A4",
                "адрес_B": "B4",
            },
            "geometry": {
                "type": "LineString",
                "coordinates": [[40.0, 57.05], [41.0, 57.05]],
            },
        },
    ]
    # Точка чуть севернее линии Near
    top = nearest_routes((40.5, 57.01), features, k=3)
    assert len(top) == 3
    assert top[0]["машина"] == "Near"
    assert top[1]["машина"] == "Closer"
    assert top[2]["машина"] == "Mid"
    assert top[0]["distance_km"] <= top[1]["distance_km"] <= top[2]["distance_km"]
    assert "адрес_A" in top[0] and "адрес_B" in top[0]
    assert top[0]["км"] == 20
    # k=1
    assert len(nearest_routes((40.5, 57.01), features, k=1)) == 1
    assert nearest_routes((40.5, 57.01), features, k=0) == []
    # без LineString — пусто
    assert nearest_routes((40.5, 57.0), [{"type": "Feature", "geometry": None}]) == []


def test_known_addresses_from_cache() -> None:
    cache = {
        "version": 1,
        "addresses": {
            "a": {
                "lat": 57.7,
                "lon": 40.9,
                "display": "Кострома",
                "error": None,
            },
            "b": {"lat": None, "lon": None, "display": "fail", "error": "x"},
            "c": {"lat": 57.6, "lon": 39.8, "display": "Ярославль"},
        },
    }
    known = known_addresses_from_cache(cache)
    assert len(known) == 2
    displays = {a["display"] for a in known}
    assert displays == {"Кострома", "Ярославль"}
    assert all("lat" in a and "lon" in a and "norm" in a for a in known)


def test_write_map_html_contains_leaflet_and_feature_count(tmp_path: Path) -> None:
    geojson = {
        "type": "FeatureCollection",
        "features": [
            {
                "type": "Feature",
                "properties": {
                    "машина": "Geely Monjaro",
                    "адрес_A": "Кострома",
                    "адрес_B": "Ярославль",
                    "км": 95.0,
                    "частота": 3,
                },
                "geometry": {
                    "type": "LineString",
                    "coordinates": [[40.9, 57.7], [39.8, 57.6]],
                },
            },
            {
                "type": "Feature",
                "properties": {
                    "машина": "Hyundai Palisade",
                    "адрес_A": "A2",
                    "адрес_B": "B2",
                    "км": 10,
                    "частота": 1,
                },
                "geometry": {
                    "type": "LineString",
                    "coordinates": [[40.5, 57.5], [40.6, 57.55]],
                },
            },
        ],
    }
    known = [
        {
            "norm": "кострома",
            "display": "г.Кострома, ул.Кузнецкая",
            "lat": 57.766,
            "lon": 40.927,
        }
    ]
    out = tmp_path / "карта.html"
    write_map_html(out, geojson, VEHICLE_COLORS, known_addresses=known)
    html = out.read_text(encoding="utf-8")
    assert "leaflet" in html.lower()
    assert "leaflet@1.9" in html
    assert "FEATURE_COUNT = 2" in html
    assert 'id="feature-count">2</span>' in html
    assert "Geely Monjaro" in html
    assert "Hyundai Palisade" in html
    assert VEHICLE_COLORS["Geely Monjaro"] in html
    assert "адрес_A" in html
    assert "chk-all" in html
    assert "tile.openstreetmap.org" in html
    assert 'id="search-input"' in html
    assert "KNOWN_ADDRESSES" in html
    assert "г.Кострома, ул.Кузнецкая" in html
    assert "photon.komoot.io" in html
    assert "nearestRoutes" in html
    assert "search-stub" not in html


XSS_ADDR_A = '<img src=x onerror=alert(1)>'
XSS_VEHICLE = '"><script>alert(1)</script>'


def _malicious_geojson() -> dict:
    return {
        "type": "FeatureCollection",
        "features": [
            {
                "type": "Feature",
                "properties": {
                    "машина": XSS_VEHICLE,
                    "адрес_A": XSS_ADDR_A,
                    "адрес_B": "Safe B",
                    "км": 95.0,
                    "частота": 3,
                },
                "geometry": {
                    "type": "LineString",
                    "coordinates": [[40.9, 57.7], [39.8, 57.6]],
                },
            }
        ],
    }


def test_write_map_html_includes_escape_html_and_uses_in_popup_nearest_search(
    tmp_path: Path,
) -> None:
    out = tmp_path / "xss.html"
    write_map_html(out, _malicious_geojson(), VEHICLE_COLORS)
    html = out.read_text(encoding="utf-8")

    assert "function escapeHtml(s)" in html
    assert 'escapeHtml(p["машина"]' in html
    assert 'escapeHtml(p["адрес_A"]' in html
    assert 'escapeHtml(p["адрес_B"]' in html
    assert 'escapeHtml(r["машина"]' in html
    assert 'escapeHtml(r["адрес_A"]' in html
    assert 'escapeHtml(r["адрес_B"]' in html
    assert "bindPopup(escapeHtml(label))" in html


def test_write_map_html_xss_payload_not_in_unescaped_html_context(
    tmp_path: Path,
) -> None:
    """Payload in GeoJSON JSON is OK; innerHTML/popup paths must escape."""
    out = tmp_path / "xss.html"
    write_map_html(out, _malicious_geojson(), VEHICLE_COLORS)
    html = out.read_text(encoding="utf-8")

    popup_section = html.split("function popupHtml", 1)[1].split("const layersByVehicle", 1)[0]
    nearest_section = html.split("function renderNearest", 1)[1].split("function showPoint", 1)[0]
    show_point_section = html.split("function showPoint", 1)[1].split("function normalizeQuery", 1)[0]

    for section in (popup_section, nearest_section, show_point_section):
        assert "onerror=" not in section
        assert "<script>" not in section

    assert "&lt;img" in html or "&lt;script" in html or "escapeHtml" in html
    assert XSS_ADDR_A not in popup_section
    assert XSS_VEHICLE not in popup_section
    assert XSS_ADDR_A not in nearest_section
    assert XSS_VEHICLE not in nearest_section


def test_write_map_html_empty_feature_collection(tmp_path: Path) -> None:
    out = tmp_path / "empty.html"
    write_map_html(
        out,
        {"type": "FeatureCollection", "features": []},
        VEHICLE_COLORS,
    )
    html = out.read_text(encoding="utf-8")
    assert out.is_file()
    assert "leaflet" in html.lower()
    assert "FEATURE_COUNT = 0" in html
    assert 'id="feature-count">0</span>' in html
    assert 'id="search-input"' in html
    for name in VEHICLE_COLORS:
        assert name in html  # палитра в JS даже без features


def test_run_writes_html(tmp_path: Path) -> None:
    xlsx = tmp_path / "pool.xlsx"
    _write_routes_ab(xlsx, _sample_rows())
    cache_path = tmp_path / "addresses.json"
    routes_path = tmp_path / "routes.geojson"
    routes = load_routes_ab(xlsx)
    save_address_cache(cache_path, _addr_cache_for_sample(routes))
    out = tmp_path / "map.html"

    report = run(
        xlsx,
        out,
        offline=False,
        geo_cache_path=cache_path,
        routes_cache_path=routes_path,
        geocode_fn=lambda _q: (_ for _ in ()).throw(AssertionError("cache")),
        osrm_fn=lambda lon1, lat1, lon2, lat2: (
            [[lon1, lat1], [lon2, lat2]],
            100.0,
        ),
        sleep_sec=0,
        osrm_sleep_sec=0,
    )
    assert out.is_file()
    assert report["html_features"] == 2
    html = out.read_text(encoding="utf-8")
    assert "leaflet" in html.lower()
    assert "FEATURE_COUNT = 2" in html
    assert "KNOWN_ADDRESSES" in html
    assert "nearestRoutes" in html
