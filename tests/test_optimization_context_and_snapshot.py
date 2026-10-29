"""OPT ContextVar isolation (asyncio) и неизменяемость LayoutRuntimeSnapshot (A1/A2)."""
from __future__ import annotations

import asyncio
import sys
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.optimization.context import (  # noqa: E402
    OPT_PLAN,
    optimization_context_scope,
)
from core.optimization.layout_runtime_snapshot import (  # noqa: E402
    build_layout_runtime_snapshot,
)
from core.plate_runtime_state import (  # noqa: E402
    new_plate_mutable_runtime_empty,
    plate_mutable_runtime_scope,
)
from core.optimization import context as opt_context  # noqa: E402
from core.optimization.orchestrator import (  # noqa: E402
    optimize_with_cascading_longitudinal_cuts,
)


def test_optimize_entrypoint_opens_fresh_context_each_call() -> None:
    """A1: публичный вход всегда в отдельном optimization_context_scope (новый state)."""
    created: list[dict] = []
    real_new = opt_context.new_optimization_context_state

    def _capture_state() -> dict:
        s = real_new()
        created.append(s)
        return s

    with patch.object(opt_context, "new_optimization_context_state", side_effect=_capture_state):
        optimize_with_cascading_longitudinal_cuts()
        optimize_with_cascading_longitudinal_cuts()
    assert len(created) == 2
    assert created[0] is not created[1]


def test_optimization_context_scope_isolates_concurrent_asyncio_tasks() -> None:
    async def _gather() -> tuple[str, str]:
        async def exercise(marker: str) -> str:
            with optimization_context_scope():
                OPT_PLAN.clear()
                OPT_PLAN["marker"] = marker
                await asyncio.sleep(0.02)
                return str(OPT_PLAN.get("marker", ""))

        a, b = await asyncio.gather(exercise("A"), exercise("B"))
        return a, b

    assert asyncio.run(_gather()) == ("A", "B")


def test_layout_runtime_snapshot_deep_copies_opt_plan() -> None:
    OPT_PLAN.clear()
    OPT_PLAN["k"] = "before"
    with plate_mutable_runtime_scope(new_plate_mutable_runtime_empty()):
        rt = build_layout_runtime_snapshot()
    OPT_PLAN["k"] = "after"
    assert dict(rt.opt_snapshot.opt_plan)["k"] == "before"


def test_snapshot_from_explicit_order_does_not_read_unbound_runtime(
    monkeypatch,
) -> None:
    """Снимок из PlateOrderContext не создаёт заказ потока, даже вне bound()."""
    monkeypatch.setenv("PLATE_RUNTIME_STRICT", "raise")
    from core.plate_order_context import PlateOrderContext
    from core.optimization.layout_runtime_snapshot import (
        build_layout_runtime_snapshot_from_plate_order_context,
    )

    ctx = PlateOrderContext.fresh_empty()
    ctx.plates.plates_1_2.append(3.3)
    snapshot = build_layout_runtime_snapshot_from_plate_order_context(ctx)
    assert snapshot.layout_cfg.plate_lists.plates_1_2 == (3.3,)
