from __future__ import annotations

import pytest
from pydantic import ValidationError

from app.core.settings import Settings, get_settings

from tests.conftest import VALID_APP_SECRET_KEY

_WEAK_KEY_ERROR = (
    r"APP_SECRET_KEY must be set via environment, must not use a known default, "
    r"and must be at least 32 characters"
)


@pytest.fixture(autouse=True)
def _isolate_settings_constructor_env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Constructor kwargs must win over process env and ``.env`` for direct ``Settings()`` tests."""
    for name in ("APP_SECRET_KEY", "APP_ENV", "COOKIE_SECURE"):
        monkeypatch.delenv(name, raising=False)


def _settings(**kwargs: object) -> Settings:
    return Settings(_env_file=None, **kwargs)


@pytest.mark.parametrize(
    "bad_key",
    [
        "",
        "short-key",
        "change-this-secret-key-in-env",
        "changeme",
        "secret",
        "a" * 31,
    ],
    ids=[
        "empty",
        "too_short",
        "default_placeholder",
        "changeme",
        "secret",
        "31_chars",
    ],
)
def test_app_secret_key_rejects_weak_or_invalid_values(bad_key: str) -> None:
    with pytest.raises(ValidationError, match=_WEAK_KEY_ERROR):
        _settings(app_secret_key=bad_key)


def test_app_secret_key_accepts_valid_value() -> None:
    settings = _settings(app_secret_key=VALID_APP_SECRET_KEY)
    assert settings.app_secret_key == VALID_APP_SECRET_KEY


def test_app_secret_key_strips_surrounding_whitespace() -> None:
    padded = f"  {VALID_APP_SECRET_KEY}  "
    settings = _settings(app_secret_key=padded)
    assert settings.app_secret_key == VALID_APP_SECRET_KEY


def test_get_settings_fails_fast_on_weak_env_key(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", "changeme")
    with pytest.raises(ValidationError, match=_WEAK_KEY_ERROR):
        get_settings()


def test_get_settings_fails_fast_on_missing_env_key(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("APP_SECRET_KEY", raising=False)
    monkeypatch.setenv("APP_SECRET_KEY", "")
    with pytest.raises(ValidationError, match=_WEAK_KEY_ERROR):
        get_settings()


@pytest.mark.parametrize(
    ("app_env", "cookie_secure", "expected"),
    [
        ("development", None, False),
        ("production", None, True),
        ("development", True, True),
        ("production", False, False),
        ("staging", True, True),
    ],
)
def test_cookie_secure_enabled(
    app_env: str,
    cookie_secure: bool | None,
    expected: bool,
) -> None:
    kwargs: dict = {
        "app_secret_key": VALID_APP_SECRET_KEY,
        "app_env": app_env,
    }
    if cookie_secure is not None:
        kwargs["cookie_secure"] = cookie_secure
    if app_env.lower() == "production":
        kwargs["bot_telegram_allowlist_raw"] = "1:admin"
    settings = _settings(**kwargs)
    assert settings.cookie_secure_enabled is expected
