"""OPT-005: потоколокальные OPT_* не делятся между потоками."""
from __future__ import annotations

import sys
import threading
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


def test_opt_plan_is_thread_local() -> None:
    import core.optimization as optimization

    barrier = threading.Barrier(2)
    out: dict[str, str] = {}

    def worker(tid: str, value: str) -> None:
        optimization.OPT_PLAN.clear()
        optimization.OPT_PLAN["tid"] = value
        barrier.wait()
        out[tid] = optimization.OPT_PLAN.get("tid", "")

    t1 = threading.Thread(target=worker, args=("a", "thread_a"))
    t2 = threading.Thread(target=worker, args=("b", "thread_b"))
    t1.start()
    t2.start()
    t1.join()
    t2.join()

    assert out["a"] == "thread_a"
    assert out["b"] == "thread_b"


def test_setattr_routes_to_thread_local() -> None:
    import core.optimization as optimization

    optimization.OPT_CASCADING_PLAN = {"k": 1}
    assert optimization.OPT_CASCADING_PLAN.get("k") == 1
    optimization.OPT_CASCADING_PLAN = {}
    assert dict(optimization.OPT_CASCADING_PLAN) == {}
