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
        "delivery_conditions": "Самовывоз",
        "payment_conditions": "100% предоплата",
        "execution_terms": "",
        "status": "в архиве",
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

    items = service.list_offers("archived")

    assert len(items) == 1
    assert items[0].kp_id == 42
    assert items[0].completion_percentage is None
    repository.get_completion_percentage.assert_not_called()


def test_list_offers_for_production_includes_completion(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.list_by_section.return_value = [_make_raw(status="в работе")]
    repository.get_completion_percentage.return_value = {"percentage": 42.5}
    service = _make_service(repository, tmp_path)

    items = service.list_offers("in_production")

    assert items[0].completion_percentage == pytest.approx(42.5)
    repository.get_completion_percentage.assert_called_once_with(42)


def test_get_details_raises_when_missing(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = None
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.get_details(999)


def test_update_discount_out_of_range(tmp_path: Path) -> None:
    service = _make_service(MagicMock(), tmp_path)

    with pytest.raises(ArchiveValidationError):
        service.update_discount(1, 150)


def test_update_discount_not_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.update_discount.return_value = False
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.update_discount(1, 10)


def test_move_to_production_requires_archived_status(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(status="в работе")
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveValidationError):
        service.move_to_production(42, "5 дней")


def test_move_to_production_happy_path(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.side_effect = [
        _make_raw(status="в архиве"),
        _make_raw(status="в работе", execution_terms="01.04.2026"),
    ]
    repository.update_execution_date.return_value = True
    repository.update_status.return_value = True
    service = _make_service(repository, tmp_path)

    details = service.move_to_production(42, "01.04.2026")

    assert details.status == "в работе"
    assert details.execution_terms == "01.04.2026"
    repository.update_execution_date.assert_called_once_with(42, "01.04.2026")
    repository.update_status.assert_called_once_with(42, "в работе")


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

    estimate = service.estimate_production(42)

    assert estimate["total_length_m"] == 200.0
    assert estimate["estimated_tracks"] >= 1
    assert estimate["estimated_days"] >= 1


def test_delete_offer_not_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.delete.return_value = False
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.delete_offer(42)


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

    path = asyncio.run(service.generate_document(42, "pdf"))

    assert path.exists()
    assert path.name == "КП_42.pdf"
    assert path.read_bytes() == b"%PDF-FAKE"
    fake_pdf.assert_called_once()


def test_generate_document_rejects_empty_plates(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(plates=[])
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveValidationError):
        asyncio.run(service.generate_document(42, "xlsx"))
