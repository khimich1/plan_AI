from __future__ import annotations

import logging

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.security.login_rate_limit import (
    configured_worker_count,
    enforce_single_instance_workers,
    rate_limit_deployment_info,
    validate_rate_limit_shared_store_config,
    warn_if_multi_worker_without_shared_store,
)

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"


@pytest.fixture(autouse=True)
def _valid_secret_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("APP_ENV", "development")
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


def test_configured_worker_count_reads_app_workers(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("UVICORN_WORKERS", raising=False)
    monkeypatch.delenv("WEB_CONCURRENCY", raising=False)
    monkeypatch.setenv("APP_WORKERS", "5")
    assert configured_worker_count() == 5


def test_rate_limit_deployment_info_without_env(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("UVICORN_WORKERS", raising=False)
    monkeypatch.delenv("WEB_CONCURRENCY", raising=False)
    monkeypatch.delenv("APP_WORKERS", raising=False)
    info = rate_limit_deployment_info()
    assert info["store"] == "in-process"
    assert info["shared_across_workers"] is False
    assert info["single_worker_required"] is True
    assert info["configured_workers"] is None
    assert info["workers_undeclared"] is True
    assert "undeclared" in info["warning"]


def test_rate_limit_deployment_info_warns_on_multi_worker(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("UVICORN_WORKERS", "4")
    info = rate_limit_deployment_info()
    assert info["configured_workers"] == 4
    assert "warning" in info
    assert "workers_undeclared" not in info


def test_warn_if_multi_worker_without_shared_store_logs(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    monkeypatch.setenv("UVICORN_WORKERS", "2")
    caplog.set_level(logging.WARNING, logger="app.security.login_rate_limit")
    warn_if_multi_worker_without_shared_store()
    assert any("in-process store" in record.message for record in caplog.records)


def test_warn_if_undeclared_workers_logs(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    monkeypatch.delenv("UVICORN_WORKERS", raising=False)
    monkeypatch.delenv("WEB_CONCURRENCY", raising=False)
    monkeypatch.delenv("APP_WORKERS", raising=False)
    caplog.set_level(logging.WARNING, logger="app.security.login_rate_limit")
    warn_if_multi_worker_without_shared_store()
    assert any("undeclared" in record.message.lower() for record in caplog.records)


def test_warn_if_multi_worker_skips_single_worker(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    monkeypatch.setenv("UVICORN_WORKERS", "1")
    caplog.set_level(logging.WARNING, logger="app.security.login_rate_limit")
    warn_if_multi_worker_without_shared_store()
    assert caplog.records == []


def test_enforce_single_instance_workers_raises_in_production(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("UVICORN_WORKERS", "2")
    with pytest.raises(RuntimeError, match="Refusing to start"):
        enforce_single_instance_workers(
            app_env="production", storage_layout="single_instance"
        )


def test_enforce_single_instance_workers_skips_development(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("UVICORN_WORKERS", "2")
    enforce_single_instance_workers(
        app_env="development", storage_layout="single_instance"
    )


def test_enforce_single_instance_workers_skips_undeclared(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("UVICORN_WORKERS", raising=False)
    monkeypatch.delenv("WEB_CONCURRENCY", raising=False)
    monkeypatch.delenv("APP_WORKERS", raising=False)
    enforce_single_instance_workers(
        app_env="production", storage_layout="single_instance"
    )


def test_enforce_single_instance_workers_skips_shared_volume(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("UVICORN_WORKERS", "2")
    enforce_single_instance_workers(
        app_env="production", storage_layout="shared_volume"
    )


def test_validate_rate_limit_shared_store_redis_raises(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("RATE_LIMIT_SHARED_STORE", "redis")
    with pytest.raises(NotImplementedError, match="redis"):
        validate_rate_limit_shared_store_config()


def test_warn_if_multi_worker_skips_when_shared_store_configured(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    monkeypatch.setenv("UVICORN_WORKERS", "4")
    monkeypatch.setenv("RATE_LIMIT_SHARED_STORE", "memory")
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


def test_health_includes_environment_in_development(client: TestClient) -> None:
    response = client.get("/health")
    assert response.status_code == 200
    payload = response.json()
    assert payload["status"] == "ok"
    assert "environment" in payload
    assert "app" in payload


def test_health_redacts_environment_in_production(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "1:admin")
    monkeypatch.setenv("BOT_AUTH_ENABLED", "true")
    get_settings.cache_clear()

    with TestClient(create_app()) as production_client:
        for path in ("/health", "/api/v1/health"):
            response = production_client.get(path)
            assert response.status_code == 200
            payload = response.json()
            assert payload == {"status": "ok"}
            assert "environment" not in payload
            assert "app" not in payload
            assert "rate_limiting" not in payload


def test_openapi_docs_disabled_in_production(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "1:admin")
    monkeypatch.setenv("BOT_AUTH_ENABLED", "true")
    get_settings.cache_clear()

    with TestClient(create_app()) as production_client:
        for path in ("/docs", "/redoc", "/openapi.json"):
            response = production_client.get(path)
            assert response.status_code == 404


def test_openapi_docs_enabled_in_development(client: TestClient) -> None:
    response = client.get("/openapi.json")
    assert response.status_code == 200
    assert "openapi" in response.json()
