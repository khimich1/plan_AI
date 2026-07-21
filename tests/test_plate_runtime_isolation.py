"""WP2/WP3: FastAPI hot paths use request-scoped PlateOrderContext (no orphan fresh_empty)."""

from __future__ import annotations

import asyncio
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from app.services.archive_service import ArchiveService, ArchiveValidationError
from app.services.commercial_workflow_service import CommercialWorkflowService
from app.services.day_documents_service import prepare_visualization_ctx
from core.plate_order_context import PlateOrderContext


def test_prepare_visualization_ctx_reuses_request_context() -> None:
    request_ctx = PlateOrderContext.fresh_empty()
    orders_2d = [{"length": 3.39, "width": 1200, "qty": 2, "load_code": 8}]
    optimization_result = {
        "_opt_status": "ok",
        "total_plates": 2,
        "primary_cuts": [],
        "secondary_cuts": [],
        "plate_assignments": [],
    }

    result = prepare_visualization_ctx(request_ctx, orders_2d, optimization_result)

    assert result is request_ctx
    assert request_ctx.plates.plate_load_details


def test_generate_day_schema_passes_request_context(monkeypatch: pytest.MonkeyPatch) -> None:
    from app.services import day_documents_service

    request_ctx = PlateOrderContext.fresh_empty()
    captured: list[PlateOrderContext] = []

    async def fake_run_in_order_context(ctx, fn, *args, **kwargs):
        captured.append(ctx)
        return (None, "/fake/schema.pdf")

    monkeypatch.setattr(day_documents_service, "run_in_order_context", fake_run_in_order_context)
    monkeypatch.setattr(
        day_documents_service,
        "_load_day_bundle",
        lambda _date: {
            "tracks": [],
            "orders_2d": [],
            "optimization_result": {},
        },
    )
    monkeypatch.setattr("app.services.day_documents_service.os.path.exists", lambda _p: True)

    async def _run() -> tuple[Path, Path]:
        return await day_documents_service.generate_day_schema(
            "2026-06-19",
            plate_order_ctx=request_ctx,
        )

    pdf_path, _cleanup = asyncio.run(_run())

    assert pdf_path == Path("/fake/schema.pdf")
    assert captured == [request_ctx]


def test_archive_schema_requires_plate_order_context(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = {
        "kp_id": 42,
        "owner_user_id": 1,
        "plates": [
            {
                "plate_name": "ПБ 78-12-8п",
                "length_m": 7.8,
                "width_m": 1.2,
                "qty": 2,
                "load_class": 800,
            }
        ],
    }
    service = ArchiveService(repository=repository, outputs_dir=tmp_path)

    with pytest.raises(ArchiveValidationError, match="Plate order context is required"):
        asyncio.run(
            service.generate_document(
                42,
                "schema",
                user={"id": 1, "role": "admin"},
            )
        )


def test_commercial_generate_files_schema_requires_context(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workflow = CommercialWorkflowService()
    monkeypatch.setattr(
        workflow,
        "_load_draft_or_raise",
        lambda _draft_id: {
            "order": MagicMock(),
            "optimization_context": MagicMock(),
            "order_data": [{"name": "ПБ 78-12-8п", "unit_price": 100.0, "qty": 1}],
            "metadata": {"manager_name": "", "client_name": "Клиент"},
        },
    )
    monkeypatch.setattr(
        "app.services.commercial_workflow_service.ensure_order_priced",
        lambda *args, **kwargs: None,
    )

    with pytest.raises(ValueError, match="Plate order context is required"):
        workflow.generate_files("draft-1", ("schema",))


def test_commercial_generate_files_schema_uses_request_context(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workflow = CommercialWorkflowService()
    request_ctx = PlateOrderContext.fresh_empty()
    captured: list[PlateOrderContext] = []

    monkeypatch.setattr(
        workflow,
        "_load_draft_or_raise",
        lambda _draft_id: {
            "order": MagicMock(),
            "optimization_context": MagicMock(),
            "order_data": [{"name": "ПБ 78-12-8п", "unit_price": 100.0, "qty": 1}],
            "metadata": {"manager_name": "", "client_name": "Клиент"},
        },
    )
    monkeypatch.setattr(
        "app.services.commercial_workflow_service.ensure_order_priced",
        lambda *args, **kwargs: None,
    )

    def fake_visualization(*, order, context, ctx, output_dir=None):
        captured.append(ctx)
        return (None, str(workflow.settings.outputs_dir / "schema.pdf"))

    monkeypatch.setattr(
        workflow.file_generation_service,
        "generate_visualization",
        fake_visualization,
    )
    monkeypatch.setattr(
        workflow,
        "_resolve_generated_file",
        lambda filename: workflow.settings.outputs_dir / Path(filename).name,
    )
    schema_path = workflow.settings.outputs_dir / "schema.pdf"
    schema_path.parent.mkdir(parents=True, exist_ok=True)
    schema_path.write_bytes(b"%PDF")

    workflow.generate_files("draft-1", ("schema",), plate_order_ctx=request_ctx)

    assert captured == [request_ctx]


def test_parallel_bound_contexts_do_not_share_plate_state() -> None:
    async def exercise(marker: float) -> float:
        ctx = PlateOrderContext.fresh_empty()
        ctx.plates.plates_1_2.append(marker)
        with ctx.bound():
            from core.plate_runtime_state import get_plate_mutable_runtime

            await asyncio.sleep(0.02)
            return get_plate_mutable_runtime().plates_1_2[0]

    async def _gather() -> tuple[float, float]:
        return await asyncio.gather(exercise(1.11), exercise(2.22))

    a, b = asyncio.run(_gather())
    assert a == pytest.approx(1.11)
    assert b == pytest.approx(2.22)


def test_optimization_service_runs_under_plate_order_ctx_bound(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.domain.models.plate_order import PlateOrder
    from app.services.optimization_service import OptimizationService
    from core.plate_runtime_state import get_plate_mutable_runtime

    request_ctx = PlateOrderContext.fresh_empty()
    runtime_matches: list[bool] = []

    def fake_optimize(*, orders_2d, **kwargs):
        runtime_matches.append(get_plate_mutable_runtime() is request_ctx.plates)
        return {
            "_opt_status": "ok",
            "total_plates": 1,
            "primary_cuts": [],
            "secondary_cuts": [],
            "plate_assignments": [],
        }

    monkeypatch.setattr(
        "app.services.optimization_service.optimize_with_cascading_longitudinal_cuts",
        fake_optimize,
    )

    order = PlateOrder.from_orders_2d(
        [{"length": 7.8, "width": 1200, "qty": 1, "load_code": 8}]
    )
    OptimizationService().optimize(order, plate_order_ctx=request_ctx)

    assert runtime_matches == [True]


def test_run_planning_pipeline_reuses_plate_order_context(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.production_planning_service import ProductionPlanningService
    from core.plate_runtime_state import get_plate_mutable_runtime
    from core.production.dto import LoadResult

    request_ctx = PlateOrderContext.fresh_empty()
    runtime_matches: list[bool] = []

    def fake_optimize_with_cuts(*, orders_2d, **kwargs):
        runtime_matches.append(get_plate_mutable_runtime() is request_ctx.plates)
        return {
            "_opt_status": "ok",
            "total_plates": 0,
            "primary_cuts": [],
            "secondary_cuts": [],
            "plate_assignments": [],
        }

    monkeypatch.setattr(
        "core.production.planning.optimize_with_cascading_longitudinal_cuts",
        fake_optimize_with_cuts,
    )

    load_result = LoadResult(
        kp_list=[],
        selected_plates=[],
        orders_2d=[{"length": 7.8, "width": 1200, "qty": 1, "load_code": 8}],
        plate_lookup_exact={},
        plate_lookup_by_length={},
    )

    with request_ctx.bound():
        ProductionPlanningService().run_planning_pipeline(
            load_result=load_result,
            plate_order_ctx=request_ctx,
        )

    assert runtime_matches == [True]
