from __future__ import annotations

import pytest
from starlette.requests import Request

from app.core.settings import get_settings
from app.security.login_rate_limit import resolve_client_ip

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"


@pytest.fixture(autouse=True)
def _valid_secret_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


def _make_request(
    *,
    client_host: str | None = "203.0.113.99",
    x_forwarded_for: str | None = None,
) -> Request:
    headers: list[tuple[bytes, bytes]] = []
    if x_forwarded_for is not None:
        headers.append((b"x-forwarded-for", x_forwarded_for.encode()))
    client = (client_host, 12345) if client_host is not None else None
    scope = {
        "type": "http",
        "headers": headers,
        "client": client,
    }
    return Request(scope)


def test_resolve_client_ip_ignores_xff_when_no_trusted_proxies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("TRUSTED_PROXY_IPS", raising=False)
    get_settings.cache_clear()

    request = _make_request(client_host="203.0.113.99", x_forwarded_for="198.51.100.1")

    assert resolve_client_ip(request) == "203.0.113.99"


def test_resolve_client_ip_uses_xff_from_trusted_proxy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("TRUSTED_PROXY_IPS", "127.0.0.1,203.0.113.99")
    get_settings.cache_clear()

    request = _make_request(client_host="203.0.113.99", x_forwarded_for="198.51.100.1, 10.0.0.1")

    assert resolve_client_ip(request) == "198.51.100.1"


def test_resolve_client_ip_ignores_xff_from_untrusted_proxy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("TRUSTED_PROXY_IPS", "127.0.0.1")
    get_settings.cache_clear()

    request = _make_request(client_host="203.0.113.99", x_forwarded_for="198.51.100.1")

    assert resolve_client_ip(request) == "203.0.113.99"


def test_resolve_client_ip_returns_unknown_without_client() -> None:
    request = _make_request(client_host=None, x_forwarded_for="198.51.100.1")

    assert resolve_client_ip(request) == "unknown"
