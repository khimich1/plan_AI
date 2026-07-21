from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from tests.helpers.auth_fixtures import patch_auth_users
from app.security.session import create_session_token

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    get_settings.cache_clear()
    patch_auth_users(
        monkeypatch,
        [
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


def test_web_offers_page_redirects_to_spa_archive(
    client: TestClient,
) -> None:
    token = create_session_token({"id": 2, "username": "prod", "role": "production"}, ttl_seconds=60)

    response = client.get("/web/offers", cookies={"app_session": token}, follow_redirects=False)

    assert response.status_code == 303
    assert response.headers["location"] == "/commercial-offer/archive"
    assert response.headers.get("Deprecation") == "true"
