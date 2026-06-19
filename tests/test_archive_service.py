from __future__ import annotations

import asyncio
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock

import pytest

from app.services.archive_service import (
    ArchiveNotFoundError,
    ArchiveService,
    ArchiveValidationError,
)

ADMIN = {"id": 1, "role": "admin"}


def _make_raw(**overrides: Any) -> dict:
    base = {
        "kp_id": 42,
        "creation_date": "01.03.2026",
        "customer_name": "ООО Тест",
        "manager_name": "Иван Иванов",
        "discount_percent": 5.0,
        "subtotal": 1000.0,
        "vat_amount": 220.0,
        "total_amount": 1220.0,
        "logistics_cost": 0.0,
        "delivery_conditions": "Самовывоз",
        "payment_conditions": "100% предоплата",
        "execution_terms": "",
        "status": "в архиве",
        "owner_user_id": 1,
        "plates": [
            {
                "position_number": 1,
                "plate_name": "ПБ 78-12-8п",
                "length_m": 7.8,
                "width_m": 1.2,
                "qty": 2,
                "load_class": 800,
                "unit_price": 500.0,
                "discounted_price": 475.0,
                "unit_weight": 1000,
                "total_weight": 2000,
            }
        ],
    }
    base.update(overrides)
    return base


def _make_service(repository: MagicMock, tmp_path: Path) -> ArchiveService:
    return ArchiveService(repository=repository, outputs_dir=tmp_path)


def test_list_offers_for_archived_skips_completion(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.list_by_section.return_value = [_make_raw()]
    service = _make_service(repository, tmp_path)

    items = service.list_offers("archived", user=ADMIN)

    assert len(items) == 1
    assert items[0].kp_id == 42
    assert items[0].completion_percentage is None
    repository.get_completion_percentage.assert_not_called()


def test_list_offers_for_production_includes_completion(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.list_by_section.return_value = [_make_raw(status="в работе")]
    repository.get_completion_percentage.return_value = {"percentage": 42.5}
    service = _make_service(repository, tmp_path)

    items = service.list_offers("in_production", user=ADMIN)

    assert items[0].completion_percentage == pytest.approx(42.5)
    repository.get_completion_percentage.assert_called_once_with(42)


def test_get_details_raises_when_missing(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = None
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.get_details(999, user=ADMIN)


def test_update_discount_out_of_range(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw()
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveValidationError):
        service.update_discount(1, 150, user=ADMIN)


def test_update_discount_not_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw()
    repository.update_discount.return_value = False
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.update_discount(1, 10, user=ADMIN)


def test_update_logistics_cost_not_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw()
    repository.update_logistics_cost.return_value = False
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.update_logistics_cost(1, 100.0, user=ADMIN)


def test_update_logistics_cost_calls_repository_and_returns_details(tmp_path: Path) -> None:
    """После успешного апдейта возвращаются детали с logistics_cost из свежего снимка БД (get_by_id)."""
    repository = MagicMock()
    repository.update_logistics_cost.return_value = True
    repository.get_by_id.return_value = _make_raw(
        discount_percent=0.0,
        subtotal=607.0,
        vat_amount=143.0,
        total_amount=750.0,
        logistics_cost=100.0,
        plates=[
            {
                "position_number": 1,
                "plate_name": "ПБ",
                "length_m": 1.0,
                "width_m": 1.0,
                "qty": 65,
                "load_class": 800,
                "unit_price": 10.0,
                "discounted_price": 10.0,
                "unit_weight": 0.0,
                "total_weight": 0.0,
            }
        ],
    )
    service = _make_service(repository, tmp_path)

    details = service.update_logistics_cost(42, 100.0, user=ADMIN)

    repository.update_logistics_cost.assert_called_once_with(42, 100.0)
    assert details.logistics_cost == 100.0
    assert details.finance.total_amount == 750.0
    assert details.delivery_service_total_rub == pytest.approx(100.0)


def test_move_to_production_requires_archived_status(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(status="в работе")
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveValidationError):
        service.move_to_production(42, "5 дней", user=ADMIN)


def test_move_to_production_happy_path(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.side_effect = [
        _make_raw(status="в архиве"),
        _make_raw(status="в работе", execution_terms="01.04.2026"),
    ]
    repository.update_execution_date.return_value = True
    repository.update_status.return_value = True
    service = _make_service(repository, tmp_path)

    details = service.move_to_production(42, "01.04.2026", user=ADMIN)

    assert details.status == "в работе"
    assert details.execution_terms == "01.04.2026"
    repository.update_execution_date.assert_called_once_with(42, "01.04.2026")
    repository.update_status.assert_called_once_with(42, "в работе")


def test_move_to_production_normalizes_iso_date(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.side_effect = [
        _make_raw(status="в архиве"),
        _make_raw(status="в работе", execution_terms="05.06.2026"),
    ]
    repository.update_execution_date.return_value = True
    repository.update_status.return_value = True
    service = _make_service(repository, tmp_path)

    service.move_to_production(42, "2026-06-05", user=ADMIN)

    repository.update_execution_date.assert_called_once_with(42, "05.06.2026")


def test_move_to_production_normalizes_ddmmyyyy(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.side_effect = [
        _make_raw(status="в архиве"),
        _make_raw(status="в работе", execution_terms="05.06.2026"),
    ]
    repository.update_execution_date.return_value = True
    repository.update_status.return_value = True
    service = _make_service(repository, tmp_path)

    service.move_to_production(42, "05.06.2026", user=ADMIN)

    repository.update_execution_date.assert_called_once_with(42, "05.06.2026")


def test_move_to_production_normalizes_five_days(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from datetime import datetime as dt_cls

    base = dt_cls(2026, 5, 31, 12, 0, 0)

    class FixedNowDatetime(dt_cls):
        @classmethod
        def now(cls, tz=None):
            return cls(base.year, base.month, base.day, base.hour, base.minute, base.second)

    monkeypatch.setattr("core.execution_terms.datetime", FixedNowDatetime)
    repository = MagicMock()
    repository.get_by_id.side_effect = [
        _make_raw(status="в архиве"),
        _make_raw(status="в работе", execution_terms="05.06.2026"),
    ]
    repository.update_execution_date.return_value = True
    repository.update_status.return_value = True
    service = _make_service(repository, tmp_path)

    service.move_to_production(42, "5 дней", user=ADMIN)

    repository.update_execution_date.assert_called_once_with(42, "05.06.2026")


def test_parse_execution_terms_formats() -> None:
    assert ArchiveService._parse_execution_terms("01.04.2026") == "01.04.2026"
    assert ArchiveService._parse_execution_terms("2026-04-01") == "01.04.2026"
    result_days = ArchiveService._parse_execution_terms("7 дней")
    assert len(result_days) == 10


def test_parse_execution_terms_invalid() -> None:
    with pytest.raises(ArchiveValidationError):
        ArchiveService._parse_execution_terms("")
    with pytest.raises(ArchiveValidationError):
        ArchiveService._parse_execution_terms("скоро")


def test_estimate_production(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(
        plates=[
            {"length_m": 50.0, "qty": 2},
            {"length_m": 100.0, "qty": 1},
        ]
    )
    service = _make_service(repository, tmp_path)

    estimate = service.estimate_production(42, user=ADMIN)

    assert estimate["total_length_m"] == 200.0
    assert estimate["estimated_tracks"] >= 1
    assert estimate["estimated_days"] >= 1


def test_delete_offer_not_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = None
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.delete_offer(42, user=ADMIN)


def test_generate_document_pdf(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw()
    service = _make_service(repository, tmp_path)

    class FakeBuffer:
        def getvalue(self) -> bytes:
            return b"%PDF-FAKE"

    fake_pdf = MagicMock(return_value=FakeBuffer())
    monkeypatch.setattr(
        "app.services.archive_service.generate_commercial_offer_pdf",
        fake_pdf,
    )

    path = asyncio.run(service.generate_document(42, "pdf", user=ADMIN))

    assert path.exists()
    assert path.name == "КП_42.pdf"
    assert path.read_bytes() == b"%PDF-FAKE"
    fake_pdf.assert_called_once()
    call_kwargs = fake_pdf.call_args.kwargs
    assert call_kwargs["delivery_conditions"] == "Самовывоз"
    assert call_kwargs["payment_conditions"] == "100% предоплата"


def test_generate_document_rejects_empty_plates(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(plates=[])
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveValidationError):
        asyncio.run(service.generate_document(42, "xlsx", user=ADMIN))


def test_search_by_number_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(kp_id=42)
    service = _make_service(repository, tmp_path)

    result = service.search(user=ADMIN, kp_id=42)

    assert result.mode == "number"
    assert result.total == 1
    assert result.items[0].kp_id == 42
    assert result.truncated is False


def test_search_by_number_not_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = None
    service = _make_service(repository, tmp_path)

    result = service.search(user=ADMIN, kp_id=999)

    assert result.mode == "number"
    assert result.total == 0
    assert result.items == []


def test_search_by_customer_delegates_to_repository(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.search_by_customer_name.return_value = (
        [_make_raw(kp_id=10, customer_name="ООО Ромашка")],
        1,
    )
    service = _make_service(repository, tmp_path)

    result = service.search(user=ADMIN, customer="Ромашка")

    assert result.mode == "customer"
    assert result.total == 1
    assert result.items[0].customer_name == "ООО Ромашка"
    assert result.truncated is False
    repository.search_by_customer_name.assert_called_once_with("Ромашка", limit=50)


def test_search_by_customer_truncated_flag(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.search_by_customer_name.return_value = (
        [_make_raw(kp_id=i) for i in range(50)],
        60,
    )
    service = _make_service(repository, tmp_path)

    result = service.search(user=ADMIN, customer="Тест")

    assert result.truncated is True
    assert result.total == 60
    assert len(result.items) == 50
