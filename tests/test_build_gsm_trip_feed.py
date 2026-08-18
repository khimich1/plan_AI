"""Unit-тесты ленты «поездки+заправки» и экспорта станций на карту."""

from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

from scripts.build_gsm_trip_feed import (
    FuelTx,
    TripLeg,
    _station_group_key,
    build_stations_geojson,
    clean_station_display,
    hm_to_min,
    load_manual_stations,
    norm_card,
    norm_consumption,
    save_manual_template,
    split_cards,
    tx_position,
)
from scripts.build_gsm_trip_feed import _extract_plus_code


def _tx(
    dt: datetime,
    card: str = "3005454266",
    service: str = "АИ-95",
    qty: float | None = 40.0,
    amount: float | None = 2000.0,
    brand: str = "Газпромнефть",
    city: str = "Кострома",
    address: str = "г.Кострома, ул.Ленина, д.1",
    addr_norm: str | None = None,
) -> FuelTx:
    return FuelTx(
        dt=dt,
        card=card,
        service=service,
        qty=qty,
        unit="л",
        price=50.0,
        amount=amount,
        vendor="",
        brand=brand,
        city=city,
        address=address,
        addr_norm=addr_norm if addr_norm is not None else address.lower(),
        source="test.xls",
    )


def _leg(
    vehicle: str = "МАЗ",
    day: date = date(2026, 1, 12),
    km: float | None = 100.0,
    norm_summer: float | None = 30.0,
    norm_winter: float | None = 33.0,
) -> TripLeg:
    return TripLeg(
        vehicle=vehicle,
        plate="А001АА44",
        day=day,
        seq=1,
        addr_from="Кострома",
        addr_to="Ярославль",
        time_dep="08:00",
        time_ret="17:00",
        km=km,
        driver="",
        cards="3005454266",
        norm_summer=norm_summer,
        norm_winter=norm_winter,
        source="facts",
    )


def test_norm_card_strips_float_tail() -> None:
    assert norm_card(3005454266.0) == "3005454266"
    assert norm_card(" 3005454266 ") == "3005454266"
    assert norm_card(None) == ""


def test_split_cards() -> None:
    assert split_cards("3005454266, 3005454268") == ["3005454266", "3005454268"]
    assert split_cards("") == []


def test_hm_to_min() -> None:
    assert hm_to_min("07:10") == 430
    assert hm_to_min("") is None
    assert hm_to_min("мусор") is None


def test_tx_position() -> None:
    interval = (8 * 60, 17 * 60)
    assert tx_position(_tx(datetime(2026, 1, 12, 7, 0)), interval) == "до_выезда"
    assert tx_position(_tx(datetime(2026, 1, 12, 12, 0)), interval) == "в_рейсе"
    assert (
        tx_position(_tx(datetime(2026, 1, 12, 19, 30)), interval)
        == "после_возвращения"
    )
    assert tx_position(_tx(datetime(2026, 1, 12, 12, 0)), None) == ""


def test_norm_consumption_winter_summer() -> None:
    winter_leg = _leg(day=date(2026, 1, 12), km=100.0)
    assert norm_consumption(winter_leg) == 33.0
    summer_leg = _leg(day=date(2026, 7, 12), km=100.0)
    assert norm_consumption(summer_leg) == 30.0
    assert norm_consumption(_leg(km=None)) is None


def test_clean_station_display_drops_junk_and_appends_city() -> None:
    out = clean_station_display(
        "АЗС №2156, д.Кузнечиха, ул.Луговая, д.12, Территория", "Кострома"
    )
    assert "АЗС" not in out
    assert "Территория" not in out
    assert "ул.Луговая" in out
    assert "Кострома" in out


def test_extract_plus_code() -> None:
    assert _extract_plus_code("PVGW+3P Кострома, Россия") == (
        "PVGW+3P",
        "Кострома",
        "Кострома, Россия",
    )
    assert _extract_plus_code("ул. Ленина, 1") is None


def test_plus_code_reference_uses_full_tail(monkeypatch) -> None:
    """Референс для recoverNearest — весь хвост адреса с регионом.

    Иначе «Сусанино, Костромская обл.» разрешается в одноимённый посёлок
    Ленинградской области (было: крюк 1596 км на карте).
    """
    from scripts.build_gsm_trip_feed import resolve_plus_codes

    tx = _tx(
        datetime(2026, 1, 12),
        address="5J49+28 Сусанино, Костромская область, Россия",
        addr_norm="5j49+28 сусанино костромская область россия",
    )
    addresses: dict = {}
    queries: list[str] = []

    def fake_geocode(query: str):
        queries.append(query)
        # правильный ответ возможен только при полном хвосте с регионом
        if "костромская" in query.lower():
            return (57.84, 41.17, "Сусанино, Костромская обл.")
        return (59.15, 30.62, "Сусанино, Ленинградская обл.")

    monkeypatch.setattr(
        "scripts.build_gsm_trip_feed.nominatim_geocode", fake_geocode
    )
    monkeypatch.setattr("scripts.build_gsm_trip_feed.time.sleep", lambda _s: None)
    resolved = resolve_plus_codes([tx], addresses, offline=False)
    assert resolved == 1
    assert queries and "костромская область" in queries[0].lower()
    lat = addresses[tx.addr_norm]["lat"]
    assert 57.0 < lat < 59.0  # Костромская обл., не Ленинградская


def test_station_group_key_prefers_coords() -> None:
    tx = _tx(datetime(2026, 1, 12), addr_norm="г кострома ул ленина 1")
    addresses = {"г кострома ул ленина 1": {"lat": 57.7679, "lon": 40.9269}}
    assert _station_group_key(tx, addresses) == "57.7679,40.9269"


def test_station_group_key_normalizes_gorod_token() -> None:
    tx = _tx(datetime(2026, 1, 12), addr_norm="г кострома ул ленина 1")
    assert _station_group_key(tx, {}) == "кострома ул ленина 1"


def test_build_stations_geojson_aggregates_and_skips_nocoords() -> None:
    norm = "г кострома ул ленина 1"
    addresses = {norm: {"lat": 57.7679, "lon": 40.9269, "source": "nominatim"}}
    fuel1 = _tx(datetime(2026, 1, 12, 9, 0), qty=40.0, addr_norm=norm)
    fuel2 = _tx(datetime(2026, 1, 20, 9, 0), qty=60.0, addr_norm=norm)
    wash = _tx(
        datetime(2026, 1, 13, 10, 0),
        service="Мойка",
        qty=None,
        amount=500.0,
        addr_norm=norm,
    )
    no_coords = _tx(datetime(2026, 1, 14, 9, 0), addr_norm="километр 11 м-8")
    matched = {("МАЗ", date(2026, 1, 12)): [fuel1]}
    fc = build_stations_geojson(
        [fuel1, fuel2, wash, no_coords], matched, [fuel2, wash, no_coords],
        addresses, {"3005454266": "МАЗ"},
    )
    assert len(fc["features"]) == 1
    props = fc["features"][0]["properties"]
    assert fc["features"][0]["geometry"]["coordinates"] == [40.9269, 57.7679]
    assert props["тип"] == "азс+мойка"
    assert props["заправок"] == 2
    assert props["моек"] == 1
    assert props["литров"] == 100.0  # литры мойки не считаем
    assert props["сумма_руб"] == 4500.0
    assert props["вне_поездок"] == 2
    assert props["машины"] == "МАЗ"
    assert props["первая_дата"] == "2026-01-12"
    assert props["последняя_дата"] == "2026-01-20"


def test_manual_stations_roundtrip(tmp_path: Path) -> None:
    manual_path = tmp_path / "stations_manual.json"
    added = save_manual_template(
        manual_path, ["километр 11 м-8"], {"километр 11 м-8": "М-8, 11-й километр"}
    )
    assert added == 1
    # Заготовка без координат не подхватывается
    assert load_manual_stations(manual_path) == {}

    data = json.loads(manual_path.read_text(encoding="utf-8"))
    data["stations"]["километр 11 м-8"]["lat"] = 57.8
    data["stations"]["километр 11 м-8"]["lon"] = 40.9
    manual_path.write_text(json.dumps(data), encoding="utf-8")

    loaded = load_manual_stations(manual_path)
    assert loaded["километр 11 м-8"]["lat"] == 57.8
    assert loaded["километр 11 м-8"]["source"] == "manual"
    # Повторный прогон не дублирует и не затирает заполненное
    assert save_manual_template(manual_path, ["километр 11 м-8"], {}) == 0
    assert load_manual_stations(manual_path)["километр 11 м-8"]["lat"] == 57.8
