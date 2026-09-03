from __future__ import annotations

import pytest
from pydantic import ValidationError

from app.core.settings import Settings

from tests.conftest import VALID_APP_SECRET_KEY


@pytest.fixture(autouse=True)
def _isolate_settings_constructor_env(monkeypatch: pytest.MonkeyPatch) -> None:
    for name in ("APP_SECRET_KEY", "APP_ENV", "APP_DEBUG", "COOKIE_SECURE"):
        monkeypatch.delenv(name, raising=False)


def _settings(**kwargs: object) -> Settings:
    return Settings(_env_file=None, **kwargs)


def test_production_rejects_app_debug_true() -> None:
    with pytest.raises(
        ValidationError,
        match="APP_DEBUG must be false when APP_ENV=production",
    ):
        _settings(
            app_secret_key=VALID_APP_SECRET_KEY,
            app_env="production",
            app_debug=True,
            bot_telegram_allowlist_raw="123456789:admin",
        )


def test_production_allows_app_debug_false() -> None:
    settings = _settings(
        app_secret_key=VALID_APP_SECRET_KEY,
        app_env="production",
        app_debug=False,
        bot_telegram_allowlist_raw="123456789:admin",
    )
    assert settings.app_env == "production"
    assert settings.app_debug is False


def test_development_allows_app_debug_true() -> None:
    settings = _settings(
        app_secret_key=VALID_APP_SECRET_KEY,
        app_env="development",
        app_debug=True,
    )
    assert settings.app_debug is True


def test_commercial_ocr_uploads_per_hour_allows_zero() -> None:
    settings = _settings(
        app_secret_key=VALID_APP_SECRET_KEY,
        commercial_ocr_uploads_per_hour=0,
    )
    assert settings.commercial_ocr_uploads_per_hour == 0

    defaulted = _settings(app_secret_key=VALID_APP_SECRET_KEY)
    assert defaulted.commercial_ocr_uploads_per_hour == 0

    with pytest.raises(ValidationError):
        _settings(
            app_secret_key=VALID_APP_SECRET_KEY,
            commercial_ocr_uploads_per_hour=-1,
        )
