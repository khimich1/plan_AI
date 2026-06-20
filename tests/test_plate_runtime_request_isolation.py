"""WP4 / S5: parallel HTTP requests must not leak plate runtime state between requests."""

from __future__ import annotations

import asyncio
from typing import Any

import httpx
import pytest
from fastapi import Depends, FastAPI
from httpx import ASGITransport

from app.core.settings import get_settings
from app.dependencies.plate_context import get_plate_order_context
from app.main import create_app
from app.middleware.plate_runtime_isolation import PlateMutableRuntimeIsolationMiddleware
from app.repositories.auth_repository import AuthRepository
from app.security.session import create_session_token
from core import config_and_data as cfg
from core.plate_order_context import PlateOrderContext
from core.plate_runtime_state import get_plate_mutable_runtime


def _auth_headers() -> dict[str, str]:
    token = create_session_token(
        {"id": 1, "username": "tester", "role": "admin"},
        ttl_seconds=300,
    )
    return {"Cookie": f"app_session={token}"}


@pytest.fixture()
def isolated_app(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    get_settings.cache_clear()
    monkeypatch.setattr(
        AuthRepository,
        "list_users",
        lambda self: [
            {
                "id": 1,
                "username": "tester",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            }
        ],
    )
    return create_app()


def _parallel_probe_app() -> FastAPI:
    app = FastAPI()
    app.add_middleware(PlateMutableRuntimeIsolationMiddleware)

    @app.get("/probe/{marker}")
    async def _probe(
        marker: float,
        ctx: PlateOrderContext = Depends(get_plate_order_context),
    ) -> dict[str, Any]:
        cfg.PLATES_1_2.append(marker)
        await asyncio.sleep(0.03)
        runtime_first = cfg.PLATES_1_2[0] if cfg.PLATES_1_2 else None
        ctx_first = ctx.plates.plates_1_2[0] if ctx.plates.plates_1_2 else None
        return {
            "marker": marker,
            "runtime_first": runtime_first,
            "ctx_first": ctx_first,
            "runtime_matches_ctx": get_plate_mutable_runtime() is ctx.plates,
            "runtime_len": len(cfg.PLATES_1_2),
        }

    return app


async def _gather_probe(marker_a: float, marker_b: float) -> tuple[dict[str, Any], dict[str, Any]]:
    transport = ASGITransport(app=_parallel_probe_app())
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        responses = await asyncio.gather(
            client.get(f"/probe/{marker_a}"),
            client.get(f"/probe/{marker_b}"),
        )
    return responses[0].json(), responses[1].json()


def test_parallel_middleware_probe_requests_do_not_mix_plate_state() -> None:
    left, right = asyncio.run(_gather_probe(1.11, 2.22))

    assert left["marker"] == pytest.approx(1.11)
    assert right["marker"] == pytest.approx(2.22)
    assert left["runtime_first"] == pytest.approx(1.11)
    assert right["runtime_first"] == pytest.approx(2.22)
    assert left["ctx_first"] == pytest.approx(1.11)
    assert right["ctx_first"] == pytest.approx(2.22)
    assert left["runtime_matches_ctx"] is True
    assert right["runtime_matches_ctx"] is True
    assert left["runtime_len"] == 1
    assert right["runtime_len"] == 1


async def _gather_generate_preview(
    app,
    text_a: str,
    text_b: str,
) -> tuple[dict[str, Any], dict[str, Any]]:
    headers = _auth_headers()
    transport = ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        responses = await asyncio.gather(
            client.post(
                "/api/v1/commercial/generate-preview",
                json={"text": text_a},
                headers=headers,
            ),
            client.post(
                "/api/v1/commercial/generate-preview",
                json={"text": text_b},
                headers=headers,
            ),
        )
    assert responses[0].status_code == 200, responses[0].text
    assert responses[1].status_code == 200, responses[1].text
    return responses[0].json(), responses[1].json()


def _preview_signature(payload: dict[str, Any]) -> dict[str, Any]:
    order = payload.get("order") or {}
    plates_1_2 = order.get("plates_1_2") or []
    order_data = payload.get("order_data") or []
    optimization = payload.get("optimization") or {}
    load_details = order.get("plate_load_details") or {}
    if isinstance(load_details, list):
        plate_count = len(plates_1_2)
    else:
        plate_count = sum(load_details.values())
    lengths_in_order_data = [
        float(row.get("length_m") or 0)
        for row in order_data
        if isinstance(row, dict) and row.get("length_m")
    ]
    return {
        "plate_count": plate_count,
        "first_length": plates_1_2[0] if plates_1_2 else None,
        "order_data_len": len(order_data),
        "order_data_lengths": lengths_in_order_data,
        "total_plates": optimization.get("total_plates"),
    }


def test_parallel_generate_preview_requests_keep_distinct_plate_data(
    isolated_app,
) -> None:
    text_a = "ПБ 78-12-8п 2"
    text_b = "ПБ 33-12-8п 3"

    payload_a, payload_b = asyncio.run(
        _gather_generate_preview(isolated_app, text_a, text_b)
    )

    sig_a = _preview_signature(payload_a)
    sig_b = _preview_signature(payload_b)

    assert sig_a["first_length"] == pytest.approx(7.8)
    assert sig_b["first_length"] == pytest.approx(3.3)
    assert sig_a["total_plates"] == 2
    assert sig_b["total_plates"] == 3
    assert sig_a["order_data_len"] >= 1
    assert sig_b["order_data_len"] >= 1
    assert 7.8 in sig_a["order_data_lengths"] or sig_a["first_length"] == pytest.approx(7.8)
    assert 3.3 in sig_b["order_data_lengths"] or sig_b["first_length"] == pytest.approx(3.3)
    assert sig_a["first_length"] != sig_b["first_length"]


def test_parallel_preview_does_not_leave_request_data_in_thread_runtime(
    isolated_app,
) -> None:
    asyncio.run(
        _gather_generate_preview(
            isolated_app,
            "ПБ 78-12-8п 1",
            "ПБ 66-12-8п 1",
        )
    )

    runtime = get_plate_mutable_runtime()
    assert 7.8 not in runtime.plates_1_2
    assert 6.6 not in runtime.plates_1_2
