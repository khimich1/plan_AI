"""Проверки search_kp_by_customer_name: LIKE, регистр, экранирование, лимит."""

from __future__ import annotations

import pytest

from core import kp_db


@pytest.fixture()
def iso_db(tmp_path_factory: pytest.TempPathFactory) -> str:
    db_path = str(tmp_path_factory.mktemp("kp_search") / "plita.db")
    kp_db.init_schema(db_path)
    return db_path


def _save_kp(
    iso_db: str,
    customer_name: str,
    *,
    status: str = "в архиве",
) -> int:
    order_data = [
        {
            "name": "ПБ",
            "qty": 1,
            "unit_price": 10.0,
            "length_m": 1.0,
            "width_m": 1.0,
        }
    ]
    return kp_db.save_kp_to_db(
        "01.03.2026",
        order_data,
        customer_name=customer_name,
        db_path=iso_db,
        status=status,
    )


def test_search_by_customer_name_partial_match(iso_db: str) -> None:
    _save_kp(iso_db, "ООО Ромашка")
    _save_kp(iso_db, "ИП Петров")

    rows, total = kp_db.search_kp_by_customer_name("ромаш", db_path=iso_db)

    assert total == 1
    assert len(rows) == 1
    assert rows[0]["customer_name"] == "ООО Ромашка"


def test_search_by_customer_name_case_insensitive(iso_db: str) -> None:
    _save_kp(iso_db, "ООО БЕТОН")

    rows, total = kp_db.search_kp_by_customer_name("бетон", db_path=iso_db)

    assert total == 1
    assert rows[0]["customer_name"] == "ООО БЕТОН"


def test_search_by_customer_name_escapes_like_wildcards(iso_db: str) -> None:
    _save_kp(iso_db, "100% скидка")
    _save_kp(iso_db, "ООО Обычный")

    rows, total = kp_db.search_kp_by_customer_name("100%", db_path=iso_db)

    assert total == 1
    assert rows[0]["customer_name"] == "100% скидка"


def test_search_by_customer_name_orders_by_kp_id_desc(iso_db: str) -> None:
    first = _save_kp(iso_db, "ООО Альфа")
    second = _save_kp(iso_db, "ООО Альфа-2")

    rows, total = kp_db.search_kp_by_customer_name("Альфа", db_path=iso_db)

    assert total == 2
    assert [row["kp_id"] for row in rows] == sorted([first, second], reverse=True)


def test_search_by_customer_name_limit_and_total(iso_db: str) -> None:
    for index in range(55):
        _save_kp(iso_db, f"ООО Массовый-{index}")

    rows, total = kp_db.search_kp_by_customer_name("Массовый", limit=50, db_path=iso_db)

    assert total == 55
    assert len(rows) == 50
