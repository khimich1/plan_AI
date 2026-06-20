from __future__ import annotations

from unittest.mock import MagicMock

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.security.session import create_session_token
from app.services.offers_service import OffersService

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    get_settings.cache_clear()
    monkeypatch.setattr(
        AuthRepository,
        "list_users",
        lambda self: [
            {
                "id": 2,
                "username": "prod",
                "role": "production",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            }
        ],
    )
    return TestClient(create_app())


def test_web_offers_page_uses_offers_service_with_user(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    token = create_session_token({"id": 2, "username": "prod", "role": "production"}, ttl_seconds=60)
    fake_service = MagicMock()
    fake_service.list_offers.return_value = []
    monkeypatch.setattr("app.web.router.OffersService", lambda: fake_service)

    response = client.get("/web/offers", cookies={"app_session": token})

    assert response.status_code == 200
    fake_service.list_offers.assert_called_once()
    call_kwargs = fake_service.list_offers.call_args.kwargs
    assert call_kwargs["status"] == "all"
    assert call_kwargs["limit"] == 100
    assert call_kwargs["user"]["id"] == 2
    assert call_kwargs["user"]["role"] == "production"
