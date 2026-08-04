"""SHIP-101: нормализация имён, авто-дедуп при импорте, CarrierService (list/merge)."""

from __future__ import annotations

import sqlite3

import pytest
from openpyxl import Workbook

from app.services.carrier_service import CarrierError, CarrierService
from core.carrier_catalog import import_carriers_from_xlsx, normalize_carrier_name
from tests.helpers import kp_db_fixtures as fx


@pytest.mark.parametrize(
    ("raw", "normalized"),
    [
        ('ООО «ТрансЛайн»', "транслайн"),
        ("ООО ТрансЛайн", "транслайн"),
        ('ООО "ТрансЛайн"', "транслайн"),
        ("  ООО   «ТрансЛайн»  ", "транслайн"),
        ('ИП Иванов И.И.', "иванов и и"),
        ('АО "Вектор"', "вектор"),
        ("ПАО «Совтрансавто»", "совтрансавто"),
        ("ЗАО «Ёлка»", "елка"),
        ("ОАО «Дорога»", "дорога"),
        ("ООО «Транс-Лайн»", "транс лайн"),
        ("ООО «Груз, ЛТД»", "груз лтд"),
        ("Просто Имя", "просто имя"),
    ],
)
def test_normalize_carrier_name(raw: str, normalized: str) -> None:
    assert normalize_carrier_name(raw) == normalized


def _build_registry_xlsx(tmp_path, sheets: dict[str, list[str]]) -> str:
    path = str(tmp_path / "registry.xlsx")
    wb = Workbook()
    first = True
    for sheet_name, names in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = sheet_name
        ws.append(["Организация", "Телефон", "Контактное лицо"])
        for name in names:
            # ПДн-колонки заполнены специально — в БД попасть не должны.
            ws.append([name, "+7 900 000-00-00", "Иван Иванов"])
    wb.save(path)
    return path


def test_import_dedup_and_report(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    xlsx = _build_registry_xlsx(
        tmp_path,
        {
            "Перевозчики": ['ООО «Альфа»', "ООО Альфа", 'ИП Бета'],
            "Транспортные Компании": ["АЛЬФА", 'ООО «Гамма»'],
        },
    )
    report = import_carriers_from_xlsx(xlsx, db_path)
    assert report.inserted == 3  # альфа, бета, гамма
    assert len(report.duplicates) == 2  # «ООО Альфа» и «АЛЬФА»
    assert {dup["name"] for dup in report.duplicates} == {"ООО Альфа", "АЛЬФА"}
    assert report.per_sheet["Перевозчики"] == {"read": 3, "imported": 2}
    assert report.per_sheet["Транспортные Компании"] == {"read": 2, "imported": 1}

    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            "SELECT name, name_normalized, source_sheet FROM carriers ORDER BY id"
        ).fetchall()
        # Только имя и лист-источник: телефонов/контактов в схеме нет.
        columns = {row[1] for row in conn.execute("PRAGMA table_info(carriers)")}
    assert len(rows) == 3
    assert rows[0] == ('ООО «Альфа»', "альфа", "Перевозчики")
    assert "phone" not in columns
    assert "contact_person" not in columns


def test_import_rerun_skips_existing(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    xlsx = _build_registry_xlsx(tmp_path, {"Перевозчики": ['ООО «Альфа»']})
    first = import_carriers_from_xlsx(xlsx, db_path)
    second = import_carriers_from_xlsx(xlsx, db_path)
    assert first.inserted == 1
    assert second.inserted == 0
    assert second.skipped_existing == 1
    with sqlite3.connect(db_path) as conn:
        total = conn.execute("SELECT COUNT(*) FROM carriers").fetchone()[0]
    assert total == 1


def _seed_carrier(db_path: str, name: str, *, active: int = 1) -> int:
    with sqlite3.connect(db_path) as conn:
        cur = conn.execute(
            "INSERT INTO carriers (name, name_normalized, source_sheet, active) "
            "VALUES (?, ?, 'test', ?)",
            (name, normalize_carrier_name(name), active),
        )
        conn.commit()
        return int(cur.lastrowid)


def _seed_shipment(db_path: str, carrier_id: int | None) -> int:
    with sqlite3.connect(db_path) as conn:
        cur = conn.execute(
            "INSERT INTO shipments (shipment_date, delivery_type, carrier_id) "
            "VALUES ('2026-07-31', 'delivery', ?)",
            (carrier_id,),
        )
        conn.commit()
        return int(cur.lastrowid)


def test_carrier_list_search_and_count(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    alpha = _seed_carrier(db_path, 'ООО «Альфа Транс»')
    _seed_carrier(db_path, 'ИП Бета')
    _seed_shipment(db_path, alpha)
    _seed_shipment(db_path, alpha)

    svc = CarrierService(db_path=db_path)
    result = svc.list_carriers()
    assert result.count == 2
    alpha_item = next(item for item in result.items if item.id == alpha)
    assert alpha_item.shipments_count == 2

    found = svc.list_carriers(q="АЛЬФА")  # регистр + кириллица
    assert [item.id for item in found.items] == [alpha]
    assert svc.list_carriers(q="несуществующий").count == 0


def test_carrier_merge_moves_shipments(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    source = _seed_carrier(db_path, 'ООО «Альфа»')
    target = _seed_carrier(db_path, 'ООО «Альфа Транс»')
    shipment_id = _seed_shipment(db_path, source)

    svc = CarrierService(db_path=db_path)
    response = svc.merge(source, target)
    assert response.moved_shipments == 1

    with sqlite3.connect(db_path) as conn:
        carrier_id = conn.execute(
            "SELECT carrier_id FROM shipments WHERE id = ?", (shipment_id,)
        ).fetchone()[0]
        merged = conn.execute(
            "SELECT active, merged_into_id FROM carriers WHERE id = ?", (source,)
        ).fetchone()
    assert carrier_id == target
    assert merged == (0, target)

    # Источник неактивен: дефолтный список его не показывает.
    assert svc.list_carriers().count == 1
    assert svc.list_carriers(active=False).count == 2


def test_carrier_merge_conflicts(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    source = _seed_carrier(db_path, 'ООО «Альфа»')
    target = _seed_carrier(db_path, 'ООО «Бета»')
    inactive = _seed_carrier(db_path, 'ООО «Неактив»', active=0)

    svc = CarrierService(db_path=db_path)
    with pytest.raises(CarrierError) as self_merge:
        svc.merge(source, source)
    assert self_merge.value.code == "carrier_merge_conflict"

    with pytest.raises(CarrierError) as into_inactive:
        svc.merge(source, inactive)
    assert into_inactive.value.code == "carrier_merge_conflict"

    with pytest.raises(CarrierError) as missing_source:
        svc.merge(9999, target)
    assert missing_source.value.code == "carrier_not_found"

    with pytest.raises(CarrierError) as missing_target:
        svc.merge(source, 9999)
    assert missing_target.value.code == "carrier_not_found"

    # После конфликтов ничего не изменилось.
    with sqlite3.connect(db_path) as conn:
        active_count = conn.execute(
            "SELECT COUNT(*) FROM carriers WHERE active = 1"
        ).fetchone()[0]
    assert active_count == 2
