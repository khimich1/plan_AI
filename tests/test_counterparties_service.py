"""CTR-002: CounterpartiesService — нормализация, create, дубли."""

from __future__ import annotations

from pathlib import Path

import pytest

from app.services.counterparties_service import (
    CounterpartiesService,
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
