"""Unit tests for OffersService PDF/XLSX generation (audit Q1 v2)."""

from __future__ import annotations

from io import BytesIO
from unittest.mock import MagicMock

import pytest

from app.services.offers_service import OffersService


def _offer_without_creation_date() -> dict:
    return {
        "kp_id": 42,
        "creation_date": None,
        "customer_name": "Клиент",
        "manager_name": "Менеджер",
        "discount_percent": 0,
        "delivery_conditions": None,
        "payment_conditions": None,
        "plates": [
            {
                "plate_name": "ПБ 60-12-8п",
                "length_m": 6.0,
                "width_m": 1.2,
                "load_class": 800,
                "qty": 1,
                "unit_price": 1000.0,
                "unit_weight": 500.0,
                "total_weight": 500.0,
            }
        ],
    }


@pytest.fixture()
def offers_service() -> OffersService:
    repo = MagicMock()
    repo.get_offer.return_value = _offer_without_creation_date()
    return OffersService(kp_repository=repo)


def test_generate_pdf_without_creation_date_no_nameerror(
    offers_service: OffersService, monkeypatch: pytest.MonkeyPatch
) -> None:
    captured: dict = {}

    def fake_pdf(**kwargs):
        captured.update(kwargs)
        buf = BytesIO(b"%PDF")
        return buf

    monkeypatch.setattr(
        "app.services.offers_service.generate_commercial_offer_pdf", fake_pdf
    )
    user = {"id": 1, "role": "admin", "username": "admin"}
    filename, data = offers_service.generate_pdf(42, user=user)
    assert filename == "KP_42.pdf"
    assert data == b"%PDF"
    assert "offer_date" in captured
    assert isinstance(captured["offer_date"], str)
    assert len(captured["offer_date"]) > 0


def test_generate_xlsx_without_creation_date_no_nameerror(
    offers_service: OffersService, monkeypatch: pytest.MonkeyPatch
) -> None:
    captured: dict = {}

    def fake_xlsx(**kwargs):
        captured.update(kwargs)
        buf = BytesIO(b"PK")
        return buf

    monkeypatch.setattr(
        "app.services.offers_service.generate_commercial_offer_xlsx", fake_xlsx
    )
    user = {"id": 1, "role": "admin", "username": "admin"}
    filename, data = offers_service.generate_xlsx(42, user=user)
    assert filename == "KP_42.xlsx"
    assert data == b"PK"
    assert "offer_date" in captured
    assert isinstance(captured["offer_date"], str)
