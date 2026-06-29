"""CSRF helpers for integration tests (double-submit cookie pattern)."""

from __future__ import annotations

from fastapi.testclient import TestClient

from app.security.csrf import CSRF_COOKIE_NAME, CSRF_FORM_FIELD, CSRF_HEADER_NAME

_SAFE_METHODS = frozenset({"GET", "HEAD", "OPTIONS", "TRACE"})


def ensure_csrf_cookie(client: TestClient) -> str:
    if not client.cookies.get(CSRF_COOKIE_NAME):
        response = client.get("/health")
        assert response.status_code == 200
    token = client.cookies.get(CSRF_COOKIE_NAME)
    assert token, "CSRF cookie missing after bootstrap GET /health"
    return token


def csrf_headers(client: TestClient) -> dict[str, str]:
    return {CSRF_HEADER_NAME: ensure_csrf_cookie(client)}


class CsrfAwareTestClient(TestClient):
    """TestClient that auto-injects CSRF header/form field on unsafe methods."""

    def request(self, method: str, url: str, **kwargs):  # type: ignore[no-untyped-def]
        if method.upper() not in _SAFE_METHODS:
            token = ensure_csrf_cookie(self)
            headers = dict(kwargs.get("headers") or {})
            if CSRF_HEADER_NAME not in headers:
                headers[CSRF_HEADER_NAME] = token
                kwargs["headers"] = headers
            cookies = dict(kwargs.get("cookies") or {})
            cookies.setdefault(CSRF_COOKIE_NAME, token)
            kwargs["cookies"] = cookies
            data = kwargs.get("data")
            if isinstance(data, dict) and CSRF_FORM_FIELD not in data:
                kwargs["data"] = {**data, CSRF_FORM_FIELD: token}
        return super().request(method, url, **kwargs)
