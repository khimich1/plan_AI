"""TDD RED for Task T4: scripts/import_gsm_history.py (одноразовый импорт истории ГСМ).

Ожидаемый API (для worker) — см. блок Expected CLI/API в конце файла.
Оригиналы ``ГСМ/**`` только на чтение; запись — во временную БД / tmp-фикстуры.
"""

from __future__ import annotations

import importlib
import json
import sqlite3
from pathlib import Path

import pytest
from openpyxl import Workbook

from core import kp_db_schema

PROJECT_ROOT = Path(__file__).resolve().parents[1]
GSM_DIR = PROJECT_ROOT / "ГСМ"
ROMANU = GSM_DIR / "Роману.xlsx"
PUL = GSM_DIR / "пул_поездок.xlsx"

pytest.importorskip("openpyxl")
pytest.importorskip("xlrd")


def _import_mod():
    """Load scripts.import_gsm_history (fails RED until worker implements it)."""
    return importlib.import_module("scripts.import_gsm_history")


def _fresh_db(tmp_path: Path, name: str = "gsm_import.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _count(db_path: str, table: str, where: str = "1=1", params: tuple = ()) -> int:
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(f"SELECT COUNT(*) FROM {table} WHERE {where}", params).fetchone()
    return int(row[0])


def _conflicts_blob(conflicts) -> str:
    """Serialize conflict report for substring assertions."""
    try:
        return json.dumps(conflicts, ensure_ascii=False, default=str)
    except TypeError:
        return str(conflicts)


def _write_minimal_gsm_fixture(root: Path) -> Path:
    """Мини-дерево ГСМ в tmp (не трогает реальный ``ГСМ/``).

    Достаточно для идемпотентности справочников/маршрутов; ПЛ .xls не кладём —
    confirmed waybills проверяются интеграционным тестом на реальных данных.
    """
    root.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"
    for _ in range(6):
        ws.append([None] * 8)
    ws.append(
        [
            None,
            " Марка а/м",
            "Гос.номер",
            "Водитель",
            "Номер топливной карты",
            "Объем бака",
            "Нормы расхода (лето)",
            "Нормы расхода (зима)",
        ]
    )
    ws.append(
        [
            None,
            "Geely Monjaro",
            "О 165 ХУ 44",
            "Кулигин Н.В.",
            3005454263,
            "60 л",
            9.5,
            10.5,
        ]
    )
    ws.append(
        [
            None,
            "Geely Tugella",
            "О 848 ХР 44",
            "Cкрябин А.А., Скрябин А.А.",
            "3005454266, 3005454268",
            "55 л",
            9.4,
            10.3,
        ]
    )
    wb.save(root / "Роману.xlsx")
    wb.close()

    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "routes_ab"
    ws2.append(
        [
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
            "топливная_карта",
            "объем_бака",
            "норма_лето",
            "норма_зима",
            "водители_реестр",
            "даты_примеры",
            "дат_всего",
        ]
    )
    ws2.append(
        [
            "Geely Monjaro",
            "Geely Monjaro",
            "О 165 ХУ 44",
            "г.Кострома, ул. Кузнецкая, д.18Б",
            "г.Ярославль, пр-д Домостроителей, д.1",
            "г кострома ул кузнецкая д 18б",
            "г ярославль проезд домостроителей д 1",
            95,
            10,
            "07:10",
            "Кулигин Никита Валерьевич",
            "АИ-95",
            "3005454263",
            "60 л",
            "9.5",
            "10.5",
            "Кулигин Н.В.",
            "2025-01-10",
            10,
        ]
    )
    ws2.append(
        [
            "Geely Tugella 848",
            "Geely Tugella",
            "О 848 ХР 44",
            "г.Кострома, ул. Кузнецкая, д.18Б",
            "г.Иваново, ул. Станкостроителей, д.1",
            "г кострома ул кузнецкая д 18б",
            "г иваново ул станкостроителей д 1",
            120,
            5,
            "08:00",
            "Cкрябин Александр Анатольевич; Скрябин Александр Анатольевич",
            "АИ-95",
            "3005454266",
            "55 л",
            "9.4",
            "10.3",
            "Скрябин А.А.",
            "2025-02-01",
            5,
        ]
    )
    wb2.save(root / "пул_поездок.xlsx")
    wb2.close()

    geo = root / "geo_cache"
    geo.mkdir(exist_ok=True)
    (geo / "stations.geojson").write_text(
        json.dumps(
            {
                "type": "FeatureCollection",
                "features": [
                    {
                        "type": "Feature",
                        "geometry": {"type": "Point", "coordinates": [40.92, 57.74]},
                        "properties": {
                            "адрес": "Кострома, ул. Магистральная, д. 8",
                            "бренд": "КТК",
                            "источник_координат": "nominatim",
                        },
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    return root


# =============================================================================
# 1. Module / entrypoint
# =============================================================================


def test_import_gsm_history_module_importable():
    """scripts/import_gsm_history.py должен импортироваться; run_import — callable."""
    mod = _import_mod()
    assert callable(getattr(mod, "run_import", None)), (
        "ожидается run_import(db_path, gsm_dir) -> report"
    )
    assert callable(getattr(mod, "normalize_driver_name", None)), (
        "ожидается normalize_driver_name(name) для дедупа Скрябин"
    )


def test_run_import_signature_accepts_db_and_gsm_dir(tmp_path: Path):
    mod = _import_mod()
    gsm = _write_minimal_gsm_fixture(tmp_path / "gsm_mini")
    db_path = _fresh_db(tmp_path)
    report = mod.run_import(db_path, gsm)
    assert report is not None


# =============================================================================
# 2. Name normalization (pure helper)
# =============================================================================


def test_normalize_driver_name_latin_c_vs_cyrillic_s_skryabin():
    """Latin C (U+0043) и Cyrillic С (U+0421) в «Скрябин» → один ключ."""
    normalize_driver_name = _import_mod().normalize_driver_name

    latin = "Cкрябин Александр Анатольевич"  # C = U+0043
    cyrillic = "Скрябин Александр Анатольевич"  # С = U+0421
    assert latin[0] == "C" and ord(latin[0]) == 0x43
    assert cyrillic[0] == "С" and ord(cyrillic[0]) == 0x421

    n_lat = normalize_driver_name(latin)
    n_cyr = normalize_driver_name(cyrillic)
    assert n_lat == n_cyr
    assert n_cyr.startswith("Скрябин")
    assert ord(n_cyr[0]) == 0x421


def test_normalize_driver_name_short_and_spaced_skryabin():
    normalize_driver_name = _import_mod().normalize_driver_name
    full = normalize_driver_name("Скрябин Александр Анатольевич")
    assert normalize_driver_name("Cкрябин Александр Анатольевич") == full
    assert normalize_driver_name("  Cкрябин   Александр  Анатольевич ") == full
    short = normalize_driver_name("Cкрябин А.А.")
    assert short.startswith("Скрябин")
    assert ord(short[0]) == 0x421


# =============================================================================
# 3. Idempotency (tmp fixture — no writes into ГСМ/**)
# =============================================================================


def test_run_import_idempotent_on_tmp_fixture(tmp_path: Path):
    run_import = _import_mod().run_import
    gsm = _write_minimal_gsm_fixture(tmp_path / "gsm_mini")
    db_path = _fresh_db(tmp_path)

    run_import(db_path, gsm)

    vehicles = _count(db_path, "gsm_vehicle")
    cards = _count(db_path, "gsm_fuel_card")
    drivers = _count(db_path, "gsm_driver")
    routes = _count(db_path, "gsm_route")
    waybills = _count(db_path, "gsm_waybill")
    stations = _count(db_path, "gsm_station")

    # Фикстура без ПЛ .xls: машины/карты/маршруты/станции из xlsx+geojson.
    # Водители в основном из ПЛ — полный счётчик (8) в integration-тесте.
    assert vehicles == 2
    assert cards == 3
    assert routes == 2
    assert stations >= 1

    run_import(db_path, gsm)
    assert _count(db_path, "gsm_vehicle") == vehicles
    assert _count(db_path, "gsm_fuel_card") == cards
    assert _count(db_path, "gsm_driver") == drivers
    assert _count(db_path, "gsm_route") == routes
    assert _count(db_path, "gsm_waybill") == waybills
    assert _count(db_path, "gsm_station") == stations


# =============================================================================
# 4. Conflict report (Lonshakova) — pure helper
# =============================================================================


def test_license_conflict_helper_includes_lonshakova():
    """Два ВУ Лоншаковой → запись в отчёте конфликтов.

    Worker экспортирует ``build_driver_conflicts`` или ``collect_driver_conflicts``.
    """
    mod = _import_mod()
    builder = getattr(mod, "build_driver_conflicts", None) or getattr(
        mod, "collect_driver_conflicts", None
    )
    assert builder is not None, (
        "ожидается build_driver_conflicts(...) или collect_driver_conflicts(...)"
    )

    records = [
        {
            "full_name": "Лоншакова Надежда Евгеньевна",
            "license_number": "99 43 283502",
            "license_issued_at": "16.05.2025",
        },
        {
            "full_name": "Лоншакова Надежда Евгеньевна",
            "license_number": "99 43 2835052",
            "license_issued_at": "16.05.2025",
        },
        {
            "full_name": "Кулигин Никита Валерьевич",
            "license_number": "44 21 846315",
            "license_issued_at": "30.07.2015",
        },
    ]
    conflicts = builder(records)
    blob = _conflicts_blob(conflicts)
    assert "Лоншакова" in blob
    assert "2835052" in blob
    assert "283502" in blob


# =============================================================================
# 5. Integration against real ГСМ (read-only)
# =============================================================================


requires_real_gsm = pytest.mark.skipif(
    not ROMANU.is_file() or not PUL.is_file(),
    reason="реальный каталог ГСМ отсутствует",
)


@requires_real_gsm
@pytest.mark.integration
def test_run_import_real_gsm_counts_and_idempotent(tmp_path: Path):
    """4 машины, 6 карт, 8 водителей, 610 маршрутов, 1 confirmed imported ПЛ/машина;
    повтор без дублей; конфликт ВУ Лоншаковой в report.conflicts.
    """
    run_import = _import_mod().run_import
    db_path = _fresh_db(tmp_path, "gsm_history_real.db")

    report = run_import(db_path, GSM_DIR)

    assert _count(db_path, "gsm_vehicle") == 4
    assert _count(db_path, "gsm_fuel_card") == 6
    assert _count(db_path, "gsm_driver") == 8
    assert _count(db_path, "gsm_route") == 610

    confirmed = _count(
        db_path,
        "gsm_waybill",
        "status = ? AND source = ?",
        ("confirmed", "imported"),
    )
    assert confirmed == 4

    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            """
            SELECT vehicle_id, COUNT(*)
            FROM gsm_waybill
            WHERE status = 'confirmed' AND source = 'imported'
            GROUP BY vehicle_id
            """
        ).fetchall()
    assert len(rows) == 4
    assert all(cnt == 1 for _, cnt in rows)

    with sqlite3.connect(db_path) as conn:
        skryabin = conn.execute(
            """
            SELECT full_name FROM gsm_driver
            WHERE lower(full_name) LIKE '%крябин%'
            """
        ).fetchall()
    assert len(skryabin) == 1

    blob = _conflicts_blob(getattr(report, "conflicts", []))
    assert "Лоншакова" in blob
    assert "2835052" in blob
    assert "283502" in blob

    run_import(db_path, GSM_DIR)
    assert _count(db_path, "gsm_vehicle") == 4
    assert _count(db_path, "gsm_fuel_card") == 6
    assert _count(db_path, "gsm_driver") == 8
    assert _count(db_path, "gsm_route") == 610
    assert (
        _count(
            db_path,
            "gsm_waybill",
            "status = ? AND source = ?",
            ("confirmed", "imported"),
        )
        == 4
    )


# =============================================================================
# Expected CLI / API for worker (contract)
# =============================================================================
#
# CLI (спека gsm-module-putevye-listy.md):
#   python scripts/import_gsm_history.py --db plita.db --gsm-dir "ГСМ"
#
# Public functions in scripts/import_gsm_history.py:
#
#   normalize_driver_name(name: str) -> str
#       Map Latin/Cyrillic lookalikes (esp. C U+0043 → С U+0421), trim,
#       collapse whitespace. Same person → identical string starting with Cyrillic С.
#
#   build_driver_conflicts(records) -> list[dict]
#       Alias collect_driver_conflicts OK.
#       Input: iterable of mapping/DTO with at least full_name, license_number.
#       Lonshakova licenses «99 43 283502» vs «99 43 2835052» → conflict entry
#       whose serialized form contains «Лоншакова», «283502», «2835052».
#
#   run_import(db_path: str | Path, gsm_dir: str | Path) -> report
#       Read-only sources under gsm_dir:
#         • Роману.xlsx → gsm_vehicle, gsm_fuel_card (tank, norms)
#         • пул_поездок.xlsx / sheet routes_ab → gsm_route (610 in prod)
#         • geo_cache/stations.geojson → gsm_station (coords)
#         • vehicle folders /**/ПЛ *.xls (skip «Не заполн*», ~$):
#             drivers (ФИО, ВУ+дата, СНИЛС/табельный where present);
#             latest dated PL per vehicle → gsm_waybill
#             status='confirmed', source='imported' (odometer + fuel from file)
#       Must be idempotent (second run: no duplicate vehicles/cards/drivers/
#       routes/waybills).
#       report.conflicts includes Lonshakova license conflict (missing SNILS
#       notes for ~5 drivers — nice-to-have in same list).
#
# Report minimum surface (attributes or Mapping):
#   conflicts: list[dict]
#   optional counters: vehicles, cards, drivers, routes, waybills, stations
#
# main() / argparse: --db, --gsm-dir; exit 0 on success.
#