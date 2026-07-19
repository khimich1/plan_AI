from __future__ import annotations

import asyncio
from unittest.mock import AsyncMock, MagicMock

import pytest
from aiogram.types import CallbackQuery, Chat, Message, User
from pydantic import ValidationError

from app.core.settings import Settings, get_settings
from bot.middleware.auth import BotAuthMiddleware
from bot.middleware.role import RoleMiddleware
from bot.security.users import BotUser, has_role, resolve_bot_user
from tests.conftest import VALID_APP_SECRET_KEY


def _settings(**kwargs: object) -> Settings:
    return Settings(_env_file=None, app_secret_key=VALID_APP_SECRET_KEY, **kwargs)


@pytest.fixture
def _isolate_bot_auth_env(monkeypatch: pytest.MonkeyPatch) -> None:
    for name in (
        "APP_ENV",
        "BOT_AUTH_ENABLED",
        "BOT_TELEGRAM_ALLOWLIST",
        "BOT_AUTH_FAIL_CLOSED",
    ):
        monkeypatch.delenv(name, raising=False)


def test_parse_bot_allowlist_comma_format() -> None:
    settings = _settings(bot_telegram_allowlist_raw="111:admin,222:manager")
    assert settings.bot_telegram_allowlist == {111: "admin", 222: "manager"}


def test_parse_bot_allowlist_json_format() -> None:
    raw = '[{"id": 333, "role": "production"}]'
    settings = _settings(bot_telegram_allowlist_raw=raw)
    assert settings.bot_telegram_allowlist == {333: "production"}


def test_parse_bot_allowlist_rejects_invalid_role() -> None:
    with pytest.raises(ValidationError, match="Invalid bot role"):
        _settings(bot_telegram_allowlist_raw="1:superuser")


def test_production_rejects_empty_allowlist_when_auth_enabled(
    _isolate_bot_auth_env: None,
) -> None:
    with pytest.raises(ValidationError, match="BOT_TELEGRAM_ALLOWLIST"):
        _settings(app_env="production", bot_auth_enabled=True, bot_telegram_allowlist_raw="")


def test_production_rejects_disabled_auth(_isolate_bot_auth_env: None) -> None:
    with pytest.raises(ValidationError, match="BOT_AUTH_ENABLED can only be false"):
        _settings(app_env="production", bot_auth_enabled=False)


def test_staging_rejects_disabled_auth(_isolate_bot_auth_env: None) -> None:
    with pytest.raises(ValidationError, match="BOT_AUTH_ENABLED can only be false"):
        _settings(app_env="staging", bot_auth_enabled=False)


def test_development_allows_disabled_auth(_isolate_bot_auth_env: None) -> None:
    settings = _settings(app_env="development", bot_auth_enabled=False)
    assert settings.bot_auth_enabled is False


def test_development_allows_empty_allowlist(_isolate_bot_auth_env: None) -> None:
    settings = _settings(app_env="development", bot_telegram_allowlist_raw="")
    assert settings.bot_telegram_allowlist == {}


def test_resolve_bot_user_from_allowlist(
    monkeypatch: pytest.MonkeyPatch,
    _isolate_bot_auth_env: None,
) -> None:
    monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "42:manager")
    monkeypatch.setenv("BOT_AUTH_ENABLED", "true")
    monkeypatch.setenv("APP_ENV", "development")
    get_settings.cache_clear()
    user = resolve_bot_user(42)
    assert user is not None
    assert user.role == "manager"
    assert resolve_bot_user(99) is None
    get_settings.cache_clear()


def test_has_role() -> None:
    admin = BotUser(telegram_id=1, role="admin")
    assert has_role(admin, "admin", "manager")
    assert not has_role(admin, "production")
    assert not has_role(None, "admin")


def test_auth_middleware_denies_unknown_user(
    monkeypatch: pytest.MonkeyPatch,
    _isolate_bot_auth_env: None,
) -> None:
    monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "1:admin")
    monkeypatch.setenv("BOT_AUTH_ENABLED", "true")
    monkeypatch.setenv("APP_ENV", "development")
    get_settings.cache_clear()

    handled = False

    async def handler(event: Message, data: dict) -> None:
        nonlocal handled
        handled = True

    async def _run() -> tuple[dict, AsyncMock]:
        middleware = BotAuthMiddleware()
        message = MagicMock(spec=Message)
        message.from_user = User(id=999, is_bot=False, first_name="Test")
        message.answer = AsyncMock()
        payload: dict = {}
        await middleware(handler, message, payload)
        return payload, message.answer

    data, answer_mock = asyncio.run(_run())
    assert handled is False
    assert "bot_user" not in data
    answer_mock.assert_awaited_once()
    get_settings.cache_clear()


def test_auth_middleware_allows_allowlisted_user(
    monkeypatch: pytest.MonkeyPatch,
    _isolate_bot_auth_env: None,
) -> None:
    monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "100:admin")
    monkeypatch.setenv("BOT_AUTH_ENABLED", "true")
    monkeypatch.setenv("APP_ENV", "development")
    get_settings.cache_clear()

    handled = False

    async def handler(event: Message, data: dict) -> None:
        nonlocal handled
        handled = True
        assert data["bot_user"].role == "admin"

    async def _run() -> None:
        middleware = BotAuthMiddleware()
        message = MagicMock(spec=Message)
        message.from_user = User(id=100, is_bot=False, first_name="Admin")
        message.answer = AsyncMock()
        await middleware(handler, message, {})

    asyncio.run(_run())
    assert handled is True
    get_settings.cache_clear()


def test_auth_middleware_dev_disabled_auth_does_not_grant_admin(
    monkeypatch: pytest.MonkeyPatch,
    _isolate_bot_auth_env: None,
) -> None:
    monkeypatch.setenv("BOT_AUTH_ENABLED", "false")
    monkeypatch.setenv("APP_ENV", "development")
    get_settings.cache_clear()

    handled = False

    async def handler(event: Message, data: dict) -> None:
        nonlocal handled
        handled = True
        assert data["bot_user"].role == "production"

    async def _run() -> None:
        middleware = BotAuthMiddleware()
        message = MagicMock(spec=Message)
        message.from_user = User(id=777, is_bot=False, first_name="Dev")
        message.answer = AsyncMock()
        await middleware(handler, message, {})

    asyncio.run(_run())
    assert handled is True
    get_settings.cache_clear()


def test_auth_middleware_dev_disabled_auth_uses_allowlist_role(
    monkeypatch: pytest.MonkeyPatch,
    _isolate_bot_auth_env: None,
) -> None:
    monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "100:admin")
    monkeypatch.setenv("BOT_AUTH_ENABLED", "false")
    monkeypatch.setenv("APP_ENV", "development")
    get_settings.cache_clear()

    handled = False

    async def handler(event: Message, data: dict) -> None:
        nonlocal handled
        handled = True
        assert data["bot_user"].role == "admin"

    async def _run() -> None:
        middleware = BotAuthMiddleware()
        message = MagicMock(spec=Message)
        message.from_user = User(id=100, is_bot=False, first_name="Admin")
        message.answer = AsyncMock()
        await middleware(handler, message, {})

    asyncio.run(_run())
    assert handled is True
    get_settings.cache_clear()


def test_auth_middleware_fail_closed_outside_development(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from unittest.mock import MagicMock as MockSettings

    mock_settings = MockSettings()
    mock_settings.bot_auth_enabled = False
    mock_settings.app_env = "staging"
    monkeypatch.setattr("bot.middleware.auth.get_settings", lambda: mock_settings)

    handled = False

    async def handler(event: Message, data: dict) -> None:
        nonlocal handled
        handled = True

    async def _run() -> dict:
        middleware = BotAuthMiddleware()
        message = MagicMock(spec=Message)
        message.from_user = User(id=1, is_bot=False, first_name="User")
        message.answer = AsyncMock()
        payload: dict = {}
        await middleware(handler, message, payload)
        return payload

    payload = asyncio.run(_run())
    assert handled is False
    assert "bot_user" not in payload


def test_validate_bot_startup_fails_on_production_disabled_auth(
    monkeypatch: pytest.MonkeyPatch,
    _isolate_bot_auth_env: None,
) -> None:
    from bot.bot_main import validate_bot_startup

    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("BOT_AUTH_ENABLED", "false")
    monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "1:admin")
    get_settings.cache_clear()
    assert validate_bot_startup() is False
    get_settings.cache_clear()


def test_role_middleware_blocks_manager_from_admin_action() -> None:
    handled = False

    async def handler(event: CallbackQuery, data: dict) -> None:
        nonlocal handled
        handled = True

    async def _run() -> None:
        middleware = RoleMiddleware("admin")
        callback = MagicMock(spec=CallbackQuery)
        callback.from_user = User(id=2, is_bot=False, first_name="Mgr")
        callback.message = MagicMock(spec=Message)
        callback.message.answer = AsyncMock()
        callback.answer = AsyncMock()
        data = {"bot_user": BotUser(telegram_id=2, role="manager")}
        await middleware(handler, callback, data)

    asyncio.run(_run())
    assert handled is False


def test_manager_blocked_on_db_clear_confirmed_callback_data() -> None:
    """Manager must not pass admin-only RoleMiddleware (db_clear_confirmed path)."""
    handled = False

    async def handler(event: CallbackQuery, data: dict) -> None:
        nonlocal handled
        handled = True

    async def _run() -> None:
        middleware = RoleMiddleware("admin")
        callback = MagicMock(spec=CallbackQuery)
        callback.data = "db_clear_confirmed"
        callback.from_user = User(id=50, is_bot=False, first_name="Manager")
        callback.message = MagicMock(spec=Message)
        callback.message.answer = AsyncMock()
        callback.answer = AsyncMock()
        data = {"bot_user": BotUser(telegram_id=50, role="manager")}
        await middleware(handler, callback, data)

    asyncio.run(_run())
    assert handled is False
