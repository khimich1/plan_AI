from __future__ import annotations

import logging

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.security.login_rate_limit import (
    configured_worker_count,
    rate_limit_deployment_info,
    warn_if_multi_worker_without_shared_store,
)

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"


@pytest.fixture(autouse=True)
def _valid_secret_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


@pytest.fixture()
def client() -> TestClient:
    return TestClient(create_app())


def test_configured_worker_count_reads_uvicorn_workers(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("WEB_CONCURRENCY", raising=False)
    monkeypatch.setenv("UVICORN_WORKERS", "3")
    assert configured_worker_count() == 3


def test_configured_worker_count_reads_web_concurrency(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("UVICORN_WORKERS", raising=False)
    monkeypatch.setenv("WEB_CONCURRENCY", "2")
    assert configured_worker_count() == 2


def test_rate_limit_deployment_info_without_env() -> None:
    info = rate_limit_deployment_info()
    assert info["store"] == "in-process"
    assert info["shared_across_workers"] is False
    assert info["single_worker_required"] is True
    assert info["configured_workers"] is None
    assert "warning" not in info


def test_rate_limit_deployment_info_warns_on_multi_worker(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("UVICORN_WORKERS", "4")
    info = rate_limit_deployment_info()
    assert info["configured_workers"] == 4
    assert "warning" in info


def test_warn_if_multi_worker_without_shared_store_logs(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    monkeypatch.setenv("UVICORN_WORKERS", "2")
    caplog.set_level(logging.WARNING, logger="app.security.login_rate_limit")
    warn_if_multi_worker_without_shared_store()
    assert any("in-process store" in record.message for record in caplog.records)


def test_warn_if_multi_worker_skips_single_worker(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    monkeypatch.setenv("UVICORN_WORKERS", "1")
    caplog.set_level(logging.WARNING, logger="app.security.login_rate_limit")
    warn_if_multi_worker_without_shared_store()
    assert caplog.records == []


def test_health_includes_rate_limiting_metadata(client: TestClient) -> None:
    response = client.get("/health")
    assert response.status_code == 200
    payload = response.json()
    assert payload["rate_limiting"]["store"] == "in-process"
    assert payload["rate_limiting"]["single_worker_required"] is True


def test_api_health_includes_rate_limiting_metadata(client: TestClient) -> None:
    response = client.get("/api/v1/health")
    assert response.status_code == 200
    payload = response.json()
    assert payload["rate_limiting"]["shared_across_workers"] is False
