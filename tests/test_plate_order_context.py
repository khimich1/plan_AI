"""A1-001: PlateOrderContext, dependency, middleware (Phase 1)."""

from __future__ import annotations

import asyncio

import pytest
from fastapi import Depends, FastAPI, HTTPException
from fastapi.testclient import TestClient
from starlette.requests import Request

from app.dependencies.plate_context import get_plate_order_context
from app.middleware.plate_runtime_isolation import PlateMutableRuntimeIsolationMiddleware
from core import config_and_data as cfg
from core.optimization.context import new_optimization_context_state
from core.domain.plate_order import PlateOrder, get_current_plate_order
from core.plate_order_context import PlateOrderContext, run_in_order_context
from core.plate_runtime_state import get_plate_mutable_runtime


def test_fresh_empty_returns_empty_plates_and_fresh_opt() -> None:
    ctx = PlateOrderContext.fresh_empty()

    assert ctx.plates.plates_1_2 == []
    assert ctx.plates.plate_load_details == {}
    assert ctx.optimization == new_optimization_context_state()


def test_bound_nests_plate_and_opt_scopes() -> None:
    import core.optimization as optimization

    ctx = PlateOrderContext.fresh_empty()
    ctx.plates.plates_1_2.append(9.99)
    ctx.optimization["opt_plan"]["marker"] = "bound"

    with ctx.bound():
        assert get_plate_mutable_runtime() is ctx.plates
        cfg.PLATES_1_2.append(1.0)
        assert len(cfg.PLATES_1_2) == 2
        optimization.OPT_PLAN["worker"] = "inside"
        assert optimization.OPT_PLAN["marker"] == "bound"
        assert optimization.OPT_PLAN["worker"] == "inside"

    assert ctx.plates.plates_1_2 == [9.99, 1.0]


def test_bound_nested_scopes_restore_outer() -> None:
    outer = PlateOrderContext.fresh_empty()
    inner = PlateOrderContext.fresh_empty()
    outer.plates.plates_1_2.append(1.0)
    inner.plates.plates_1_2.append(2.0)

    with outer.bound():
        assert get_plate_mutable_runtime() is outer.plates
        with inner.bound():
            assert get_plate_mutable_runtime() is inner.plates
        assert get_plate_mutable_runtime() is outer.plates


def test_run_in_order_context_propagates_to_thread_worker() -> None:
    import core.optimization as optimization

    ctx = PlateOrderContext.fresh_empty()
    ctx.plates.plates_1_2.append(7.77)
    ctx.optimization["opt_plan"]["tid"] = "worker"

    def _read_in_thread() -> tuple[int, float, str]:
        rt = get_plate_mutable_runtime()
        return (
            len(rt.plates_1_2),
            rt.plates_1_2[0],
            optimization.OPT_PLAN.get("tid", ""),
        )

    async def _run() -> tuple[int, float, str]:
        return await run_in_order_context(ctx, _read_in_thread)

    length, first, opt_tid = asyncio.run(_run())
    assert length == 1
    assert first == pytest.approx(7.77)
    assert opt_tid == "worker"


def test_get_plate_order_context_reads_request_state() -> None:
    ctx = PlateOrderContext.fresh_empty()
    request = Request({"type": "http", "method": "GET", "path": "/", "headers": []})
    request.state.plate_order_ctx = ctx

    assert get_plate_order_context(request) is ctx


@pytest.mark.parametrize(
    "state_value",
    [None, object(), "not-a-context"],
)
def test_get_plate_order_context_raises_when_not_initialized(
    state_value: object | None,
) -> None:
    request = Request({"type": "http", "method": "GET", "path": "/", "headers": []})
    if state_value is None:
        assert not hasattr(request.state, "plate_order_ctx")
    else:
        request.state.plate_order_ctx = state_value

    with pytest.raises(HTTPException) as exc_info:
        get_plate_order_context(request)

    assert exc_info.value.status_code == 500
    assert exc_info.value.detail == "Plate order context not initialized"


def _probe_app() -> FastAPI:
    app = FastAPI()
    app.add_middleware(PlateMutableRuntimeIsolationMiddleware)

    @app.get("/probe")
    def _probe(
        ctx: PlateOrderContext = Depends(get_plate_order_context),
    ) -> dict[str, bool | int]:
        cfg.PLATES_1_2.append(1.0)
        return {
            "ctx_is_fresh": len(ctx.plates.plates_1_2) == 1,
            "runtime_matches": get_plate_mutable_runtime() is ctx.plates,
        }

    return app


def test_middleware_wires_request_state_and_bound_scope() -> None:
    client = TestClient(_probe_app())

    r1 = client.get("/probe")
    r2 = client.get("/probe")

    assert r1.status_code == 200
    assert r2.status_code == 200
    assert r1.json() == {"ctx_is_fresh": True, "runtime_matches": True}
    assert r2.json() == {"ctx_is_fresh": True, "runtime_matches": True}


def test_apply_to_globals_emits_deprecation_warning() -> None:
    order = PlateOrder()
    order.plates_1_2 = [3.39]
    ctx = PlateOrderContext.fresh_empty()
    with ctx.bound():
        with pytest.warns(DeprecationWarning, match="apply_to_globals"):
            order.apply_to_globals()


def test_get_current_plate_order_emits_deprecation_warning() -> None:
    ctx = PlateOrderContext.fresh_empty()
    with ctx.bound():
        with pytest.warns(DeprecationWarning, match="get_current_plate_order"):
            get_current_plate_order()


def test_hydrate_from_order_populates_plates() -> None:
    ctx = PlateOrderContext.fresh_empty()
    order = PlateOrder()
    order.plates_1_2 = [3.39, 3.39]
    order.plate_load_details[(3.39, 1.2, 8, "")] = 2

    ctx.hydrate_from_order(order)

    assert ctx.plates.plates_1_2 == [3.39, 3.39]
    assert ctx.plates.plate_load_details[(3.39, 1.2, 8, "")] == 2


def test_load_production_snapshot_binds_opt_in_worker() -> None:
    import core.optimization as optimization

    ctx = PlateOrderContext.fresh_empty()
    orders_2d = [{"length": 3.39, "width": 1200, "qty": 2, "load_code": 8}]
    optimization_result = {
        "_opt_status": "ok",
        "total_plates": 2,
        "primary_cuts": [],
        "secondary_cuts": [],
        "plate_assignments": [],
    }
    ctx.load_production_snapshot(orders_2d, optimization_result)

    def _read_opt() -> int:
        return optimization.OPT_CASCADING_PLAN.get("total_plates", 0)

    total = asyncio.run(run_in_order_context(ctx, _read_opt))
    assert total == 2


def test_middleware_fresh_empty_not_demo_order() -> None:
    seen: list[int] = []

    app = FastAPI()
    app.add_middleware(PlateMutableRuntimeIsolationMiddleware)

    @app.get("/len")
    def _len(ctx: PlateOrderContext = Depends(get_plate_order_context)) -> int:
        seen.append(len(ctx.plates.plates_1_2))
        return seen[-1]

    client = TestClient(app)
    assert client.get("/len").json() == 0
    assert seen == [0]
