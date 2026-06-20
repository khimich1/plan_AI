"""Shared pytest hooks: valid APP_SECRET_KEY before any project import at collection."""

from __future__ import annotations

import os

# Must run before imports that pull in ``core.db_config`` (import-time ``get_settings()``).
VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"
os.environ.setdefault("APP_SECRET_KEY", VALID_APP_SECRET_KEY)

import pytest

from app.core.settings import get_settings


@pytest.fixture(autouse=True)
def _clear_settings_cache() -> None:
    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


@pytest.fixture(autouse=True)
def _reset_login_rate_limiter() -> None:
    from app.security.login_rate_limit import reset_login_rate_limiter_for_tests

    reset_login_rate_limiter_for_tests()
    yield
    reset_login_rate_limiter_for_tests()


# ---------------------------------------------------------------------------
# Production API integration (WP3 / Q-M9)
# ---------------------------------------------------------------------------


@pytest.fixture()
def production_api_db(tmp_path, monkeypatch: pytest.MonkeyPatch) -> str:
    from tests.helpers import production_api_fixtures as paf

    return paf.configure_production_api_env(tmp_path, monkeypatch)


@pytest.fixture()
def production_api_client(production_api_db: str):
    from fastapi.testclient import TestClient

    from app.main import create_app

    del production_api_db
    return TestClient(create_app())


@pytest.fixture()
def production_admin_cookie() -> dict[str, str]:
    from tests.helpers import production_api_fixtures as paf

    return paf.session_cookie(1, "admin", "admin")


@pytest.fixture()
def production_user_cookie() -> dict[str, str]:
    from tests.helpers import production_api_fixtures as paf

    return paf.session_cookie(2, "production", "prod_user")


@pytest.fixture()
def production_manager_cookie() -> dict[str, str]:
    from tests.helpers import production_api_fixtures as paf

    return paf.session_cookie(3, "manager", "manager_a")


@pytest.fixture()
def production_built_plan(
    production_api_client,
    production_admin_cookie: dict[str, str],
) -> dict:
    from tests.helpers import production_api_fixtures as paf

    return paf.build_plan_via_api(production_api_client, production_admin_cookie)
