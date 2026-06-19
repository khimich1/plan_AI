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
    ArchiveSearchResponse,
)
from app.security.session import create_session_token

TESTER_USER = {
    "id": 1,
    "username": "tester",
    "role": "admin",
    "manager_id": None,
    "is_active": 1,
    "created_at": "2026-01-01 00:00:00",
}


@pytest.fixture()
def fake_service() -> MagicMock:
    return MagicMock()


@pytest.fixture()
def client(
    monkeypatch: pytest.MonkeyPatch,
    fake_service: MagicMock,
) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
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


def _fake_details(
    kp_id: int = 42,
    status: str = "в архиве",
    *,
    finance: ArchiveOfferFinance | None = None,
    logistics_cost: float = 0.0,
    total_cargo_weight_kg: float = 0.0,
    delivery_service_total_rub: float = 0.0,
) -> ArchiveOfferDetails:
    return ArchiveOfferDetails(
        kp_id=kp_id,
        creation_date="01.03.2026",
        customer_name="ООО Тест",
        manager_name="Иван Иванов",
        status=status,
        execution_terms=None,
        delivery_conditions=None,
        payment_conditions=None,
        finance=finance
        or ArchiveOfferFinance(
            subtotal=1000, vat_amount=220, total_amount=1220, discount_percent=5
        ),
        logistics_cost=logistics_cost,
        total_cargo_weight_kg=total_cargo_weight_kg,
        delivery_service_total_rub=delivery_service_total_rub,
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
    fake_service.list_offers.assert_called_once_with("archived", user=TESTER_USER)


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
    fake_service.update_discount.assert_called_once_with(42, 10.0, user=TESTER_USER)


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


def test_update_logistics_cost_requires_auth(client: TestClient) -> None:
    response = client.patch(
        "/api/v1/commercial/archive/42/logistics-cost",
        json={"logistics_cost": 100},
    )
    assert response.status_code == 401


def test_update_logistics_cost_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.update_logistics_cost.return_value = _fake_details(
        finance=ArchiveOfferFinance(
            subtotal=607.0,
            vat_amount=143.0,
            total_amount=750.0,
            discount_percent=0.0,
        ),
        logistics_cost=100.0,
        total_cargo_weight_kg=18414.5,
        delivery_service_total_rub=100.0,
    )

    response = client.patch(
        "/api/v1/commercial/archive/42/logistics-cost",
        json={"logistics_cost": 100},
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    fake_service.update_logistics_cost.assert_called_once_with(42, 100.0, user=TESTER_USER)
    payload = response.json()
    assert payload["logistics_cost"] == 100.0
    assert payload["finance"]["total_amount"] == 750.0


def test_update_logistics_cost_validation_negative(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    response = client.patch(
        "/api/v1/commercial/archive/42/logistics-cost",
        json={"logistics_cost": -1},
        cookies=auth_cookie,
    )
    assert response.status_code == 422


def test_update_logistics_cost_404(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    from app.services.archive_service import ArchiveNotFoundError

    fake_service.update_logistics_cost.side_effect = ArchiveNotFoundError("нет такого")

    response = client.patch(
        "/api/v1/commercial/archive/999/logistics-cost",
        json={"logistics_cost": 50},
        cookies=auth_cookie,
    )

    assert response.status_code == 404


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
    fake_service.delete_offer.assert_called_once_with(42, user=TESTER_USER)


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
    fake_service.move_to_production.assert_called_once_with(42, "5 дней", user=TESTER_USER)


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


def _fake_list_item(kp_id: int = 42, customer_name: str = "ООО Тест") -> ArchiveOfferListItem:
    return ArchiveOfferListItem(
        kp_id=kp_id,
        creation_date="01.03.2026",
        customer_name=customer_name,
        manager_name="Иван",
        discount_percent=5.0,
        subtotal=1000,
        vat_amount=220,
        total_amount=1220,
        status="в архиве",
    )


def _fake_search_response(
    *,
    mode: str = "number",
    items: list[ArchiveOfferListItem] | None = None,
    total: int | None = None,
    truncated: bool = False,
) -> ArchiveSearchResponse:
    resolved_items = items if items is not None else [_fake_list_item()]
    resolved_total = total if total is not None else len(resolved_items)
    return ArchiveSearchResponse(
        mode=mode,  # type: ignore[arg-type]
        items=resolved_items,
        total=resolved_total,
        truncated=truncated,
    )


def test_search_by_number_found(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(mode="number")

    response = client.get(
        "/api/v1/commercial/archive/search?kp_id=42",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["mode"] == "number"
    assert payload["total"] == 1
    assert len(payload["items"]) == 1
    assert payload["items"][0]["kp_id"] == 42
    fake_service.search.assert_called_once_with(user=TESTER_USER, kp_id=42)


def test_search_by_number_not_found(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(mode="number", items=[], total=0)

    response = client.get(
        "/api/v1/commercial/archive/search?kp_id=999",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["mode"] == "number"
    assert payload["total"] == 0
    assert payload["items"] == []


def test_search_by_customer_returns_items(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(
        mode="customer",
        items=[_fake_list_item(10, "ООО Ромашка"), _fake_list_item(5, "ООО Ромашка-2")],
        total=2,
    )

    response = client.get(
        "/api/v1/commercial/archive/search?customer=Ромашка",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["mode"] == "customer"
    assert payload["total"] == 2
    assert len(payload["items"]) == 2
    fake_service.search.assert_called_once_with(user=TESTER_USER, customer="Ромашка")


def test_search_by_customer_not_found(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(mode="customer", items=[], total=0)

    response = client.get(
        "/api/v1/commercial/archive/search?customer=Несуществующий",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["total"] == 0
    assert payload["items"] == []


def test_search_customer_too_short_returns_400(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    response = client.get(
        "/api/v1/commercial/archive/search?customer=А",
        cookies=auth_cookie,
    )

    assert response.status_code == 400
    assert "2" in response.json()["detail"]
    fake_service.search.assert_not_called()


def test_search_without_params_returns_422(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    response = client.get(
        "/api/v1/commercial/archive/search",
        cookies=auth_cookie,
    )

    assert response.status_code == 422
    fake_service.search.assert_not_called()


def test_search_both_params_prefers_kp_id(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(mode="number")

    response = client.get(
        "/api/v1/commercial/archive/search?kp_id=42&customer=Ромашка",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    fake_service.search.assert_called_once_with(user=TESTER_USER, kp_id=42)


def test_search_customer_truncated_flag(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(
        mode="customer",
        items=[_fake_list_item(i, "ООО Тест") for i in range(50)],
        total=73,
        truncated=True,
    )

    response = client.get(
        "/api/v1/commercial/archive/search?customer=Тест",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["truncated"] is True
    assert payload["total"] == 73
    assert len(payload["items"]) == 50


def test_download_file_returns_file(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
    tmp_path: Path,
) -> None:
    target = tmp_path / "КП_42.pdf"
    target.write_bytes(b"%PDF-TEST")

    async def fake_generate(kp_id: int, kind: str, **kwargs) -> Path:
        return target

    fake_service.generate_document = fake_generate  # type: ignore[assignment]

    response = client.get(
        "/api/v1/commercial/archive/42/files/pdf",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    assert response.content == b"%PDF-TEST"
