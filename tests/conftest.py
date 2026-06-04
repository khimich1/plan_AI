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
