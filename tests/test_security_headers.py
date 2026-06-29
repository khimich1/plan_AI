"""Security response headers on HTTP endpoints (S7)."""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app

from tests.conftest import VALID_APP_SECRET_KEY


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("APP_ENV", "development")
    get_settings.cache_clear()
    return TestClient(create_app())


def test_health_includes_security_headers(client: TestClient) -> None:
    response = client.get("/health")

    assert response.status_code == 200
    assert response.headers["x-frame-options"] == "DENY"
    assert response.headers["x-content-type-options"] == "nosniff"
    assert response.headers["referrer-policy"] == "strict-origin-when-cross-origin"
    csp = response.headers["content-security-policy-report-only"]
    assert "default-src 'self'" in csp
    assert "frame-ancestors 'none'" in csp
    assert "strict-transport-security" not in response.headers


def test_hsts_enabled_outside_development(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "1:admin")
    monkeypatch.setenv("BOT_AUTH_ENABLED", "true")
    get_settings.cache_clear()

    with TestClient(create_app()) as production_client:
        response = production_client.get("/health")

    assert response.status_code == 200
    assert "max-age=31536000" in response.headers["strict-transport-security"]
