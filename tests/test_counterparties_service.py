"""CTR-002: CounterpartiesService — нормализация, create, дубли."""

from __future__ import annotations

from pathlib import Path

import pytest

import sqlite3

from app.services.counterparties_service import (
    CounterpartiesService,
    CounterpartyValidationError,
    DuplicateCodeError,
)
from tests.helpers import kp_db_fixtures as fx


def _service(tmp_path: Path) -> CounterpartiesService:
    return CounterpartiesService(db_path=fx.make_iso_db(tmp_path))


def test_create_duplicate_code_raises_with_existing(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    first = svc.create(name="РОМАШКА", code_1c="00-00003431", inn="7701000001")
    assert first.item["code_1c"] == "00-00003431"
    assert first.warning is None

    with pytest.raises(DuplicateCodeError) as exc_info:
        svc.create(name="ДРУГАЯ", code_1c=" 00-00003431 ", inn="9900000000")
    assert "уже есть" in str(exc_info.value)
    assert exc_info.value.existing["id"] == first.item["id"]
    assert exc_info.value.existing["name"] == "РОМАШКА"


def test_create_duplicate_inn_ok_with_warning(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    svc.create(name="АЛЬФА", code_1c="00-1", inn="7701000001")
    result = svc.create(name="БЕТА", code_1c="00-2", inn="7701000001")
    assert result.item["code_1c"] == "00-2"
    assert result.warning is not None
    assert "ИНН" in result.warning


def test_create_empty_inn_has_no_warning(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    svc.create(name="АЛЬФА", code_1c="00-1")
    result = svc.create(name="БЕТА", code_1c="00-2")
    assert result.warning is None


def test_create_normalizes_trim_and_is_immediately_searchable(
    tmp_path: Path,
) -> None:
    svc = _service(tmp_path)
    created = svc.create(
        name="  СТОЛИЦА ООО  ",
        code_1c="  Б00000002  ",
        inn="  7701000001  ",
        kpp=" 770101001 ",
    )
    assert created.item["name"] == "СТОЛИЦА ООО"
    assert created.item["code_1c"] == "Б00000002"
    assert created.item["inn"] == "7701000001"
    assert created.item["source"] == "manual"

    items = svc.search("столица")
    assert len(items) == 1
    assert items[0]["id"] == created.item["id"]


def _count(db_path: str) -> int:
    with sqlite3.connect(db_path) as conn:
        return int(conn.execute("SELECT COUNT(*) FROM counterparties").fetchone()[0])


def test_resolve_new_code_inserts_searchable_manual_client(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    row = svc.resolve_active_client_for_bind(
        name="  БАРМАЛЕЙ ООО  ",
        code_1c="  00-NEW  ",
        inn="  7701000001  ",
        kpp=" 770101001 ",
    )
    assert row["name"] == "БАРМАЛЕЙ ООО"
    assert row["code_1c"] == "00-NEW"
    assert row["inn"] == "7701000001"
    assert row["kpp"] == "770101001"
    assert row["source"] == "manual"
    assert int(row["is_client"]) == 1
    assert int(row["is_active"]) == 1
    assert row.get("guid_1c") in (None, "")

    items = svc.search("бармалей")
    assert len(items) == 1
    assert items[0]["id"] == row["id"]


def test_resolve_duplicate_active_client_returns_same_id_without_rename(
    tmp_path: Path,
) -> None:
    svc = _service(tmp_path)
    first = svc.create(name="РОМАШКА", code_1c="00-00003431")
    before = _count(svc.db_path)

    row = svc.resolve_active_client_for_bind(
        name="ДРУГОЕ ИМЯ",
        code_1c=" 00-00003431 ",
        inn="9900000000",
    )
    assert row["id"] == first.item["id"]
    assert row["name"] == "РОМАШКА"
    assert _count(svc.db_path) == before


def test_resolve_supplier_or_inactive_raises_validation(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    supplier = svc.repo.insert(
        code_1c="00-sup",
        name="ПОСТАВЩИК ООО",
        is_client=False,
        source="import",
    )
    inactive = svc.repo.insert(
        code_1c="00-off",
        name="АРХИВ ООО",
        is_client=True,
        is_active=False,
        source="import",
    )

    with pytest.raises(CounterpartyValidationError, match="не отмечен как клиент"):
        svc.resolve_active_client_for_bind(name="X", code_1c="00-sup")
    with pytest.raises(CounterpartyValidationError, match="не найден"):
        svc.resolve_active_client_for_bind(name="X", code_1c="00-off")
    assert svc.get_by_id(int(supplier["id"]))["name"] == "ПОСТАВЩИК ООО"
    assert svc.get_by_id(int(inactive["id"]))["name"] == "АРХИВ ООО"


def test_resolve_empty_name_or_code_raises_value_error(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    with pytest.raises(ValueError, match="обязательны"):
        svc.resolve_active_client_for_bind(name="  ", code_1c="00-1")
    with pytest.raises(ValueError, match="обязательны"):
        svc.resolve_active_client_for_bind(name="Клиент", code_1c="  ")
    assert _count(svc.db_path) == 0
