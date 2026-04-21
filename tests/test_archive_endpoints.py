from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest
from fastapi.testclient import TestClient

from app.api.v1.endpoints.archive import get_archive_service
from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.schemas.archive import (
    ArchiveOfferDetails,
    ArchiveOfferFinance,
    ArchiveOfferListItem,
)
from app.security.session import create_session_token


@pytest.fixture()
def fake_service() -> MagicMock:
    return MagicMock()


@pytest.fixture()
def client(
    monkeypatch: pytest.MonkeyPatch,
    fake_service: MagicMock,
) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key")
    get_settings.cache_clear()
    monkeypatch.setattr(
        AuthRepository,
        "list_users",
        lambda self: [
            {
                "id": 1,
                "username": "tester",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            }
        ],
    )
    app = create_app()
    app.dependency_overrides[get_archive_service] = lambda: fake_service
    return TestClient(app)


@pytest.fixture()
def auth_cookie() -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": 1, "username": "tester", "role": "admin"},
            ttl_seconds=300,
        ),
    }


def _fake_details(kp_id: int = 42, status: str = "в архиве") -> ArchiveOfferDetails:
    return ArchiveOfferDetails(
        kp_id=kp_id,
        creation_date="01.03.2026",
        customer_name="ООО Тест",
        manager_name="Иван Иванов",
        status=status,
        execution_terms=None,
        delivery_conditions=None,
        payment_conditions=None,
        finance=ArchiveOfferFinance(
            subtotal=1000, vat_amount=220, total_amount=1220, discount_percent=5
        ),
        plates=[],
        completion_percentage=None,
    )


def test_list_requires_auth(client: TestClient) -> None:
    response = client.get("/api/v1/commercial/archive?section=archived")
    assert response.status_code == 401


def test_list_returns_items(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.list_offers.return_value = [
        ArchiveOfferListItem(
            kp_id=42,
            creation_date="01.03.2026",
            customer_name="ООО Тест",
            manager_name="Иван",
            discount_percent=5.0,
            subtotal=1000,
            vat_amount=220,
            total_amount=1220,
            status="в архиве",
        )
    ]

    response = client.get(
        "/api/v1/commercial/archive?section=archived",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert len(payload) == 1
    assert payload[0]["kp_id"] == 42
    fake_service.list_offers.assert_called_once_with("archived")


def test_get_details_404(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    from app.services.archive_service import ArchiveNotFoundError

    fake_service.get_details.side_effect = ArchiveNotFoundError("нет такого")

    response = client.get(
        "/api/v1/commercial/archive/999",
        cookies=auth_cookie,
    )

    assert response.status_code == 404


def test_update_discount_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.update_discount.return_value = _fake_details()

    response = client.patch(
        "/api/v1/commercial/archive/42/discount",
        json={"discount": 10},
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    fake_service.update_discount.assert_called_once_with(42, 10.0)


def test_update_discount_validation(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    response = client.patch(
        "/api/v1/commercial/archive/42/discount",
        json={"discount": 150},
        cookies=auth_cookie,
    )
    assert response.status_code == 422


def test_delete_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    response = client.delete(
        "/api/v1/commercial/archive/42",
        cookies=auth_cookie,
    )

    assert response.status_code == 204
    fake_service.delete_offer.assert_called_once_with(42)


def test_move_to_production_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.move_to_production.return_value = _fake_details(status="в работе")

    response = client.post(
        "/api/v1/commercial/archive/42/move-to-production",
        json={"execution_terms": "5 дней"},
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    assert response.json()["status"] == "в работе"
    fake_service.move_to_production.assert_called_once_with(42, "5 дней")


def test_move_to_production_validation_error(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    from app.services.archive_service import ArchiveValidationError

    fake_service.move_to_production.side_effect = ArchiveValidationError("bad")

    response = client.post(
        "/api/v1/commercial/archive/42/move-to-production",
        json={"execution_terms": "скоро"},
        cookies=auth_cookie,
    )

    assert response.status_code == 400


def test_search_found(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search_by_number.return_value = _fake_details()

    response = client.get(
        "/api/v1/commercial/archive/search?query=42",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["found"] is True
    assert payload["offer"]["kp_id"] == 42


def test_search_not_found(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search_by_number.return_value = None

    response = client.get(
        "/api/v1/commercial/archive/search?query=999",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["found"] is False
    assert payload["offer"] is None


def test_download_file_returns_file(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
    tmp_path: Path,
) -> None:
    target = tmp_path / "КП_42.pdf"
    target.write_bytes(b"%PDF-TEST")

    async def fake_generate(kp_id: int, kind: str) -> Path:
        return target

    fake_service.generate_document = fake_generate  # type: ignore[assignment]

    response = client.get(
        "/api/v1/commercial/archive/42/files/pdf",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    assert response.content == b"%PDF-TEST"
