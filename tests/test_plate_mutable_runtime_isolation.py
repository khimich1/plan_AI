"""Изоляция plate mutable runtime (S1 / PLATE-CTX-001)."""

from __future__ import annotations

import asyncio
import sys
import threading
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core import config_and_data as cfg  # noqa: E402 — set_plate_lists_from_text only
from core.plate_order_context import PlateOrderContext  # noqa: E402
from core.plate_runtime_state import (  # noqa: E402
    fresh_plate_mutable_request_scope,
    get_plate_mutable_runtime,
    plate_mutable_runtime_scope,
    new_plate_mutable_runtime_empty,
)


def test_fresh_thread_sees_empty_plate_runtime_not_demo() -> None:
    lengths: list[int] = []

    def _in_thread() -> None:
        lengths.append(len(get_plate_mutable_runtime().plates_1_2))

    t = threading.Thread(target=_in_thread)
    t.start()
    t.join()
    assert lengths == [0]


def test_plate_scope_isolates_concurrent_asyncio_tasks() -> None:
    async def _gather() -> tuple[int, int]:
        async def exercise(append_len: int) -> int:
            with plate_mutable_runtime_scope(new_plate_mutable_runtime_empty()) as rt:
                rt.plates_1_2.append(1.0)
                for _ in range(append_len):
                    rt.plates_1_2.append(1.0)
                await asyncio.sleep(0.02)
                return len(rt.plates_1_2)

        a, b = await asyncio.gather(exercise(2), exercise(5))
        return a, b

    assert asyncio.run(_gather()) == (3, 6)


def test_set_plate_lists_from_text_leaves_data_visible_inside_fresh_request_scope() -> None:
    with fresh_plate_mutable_request_scope():
        cfg.set_plate_lists_from_text("ПБ 66-12-8п 1")
        assert len(get_plate_mutable_runtime().plates_1_2) >= 1


def test_nested_bound_scopes_keep_independent_plate_lists() -> None:
    outer = PlateOrderContext.fresh_empty()
    inner = PlateOrderContext.fresh_empty()

    with outer.bound():
        outer.plates.plates_1_2.append(1.0)
        with inner.bound():
            inner.plates.plates_1_2.append(2.0)
            assert get_plate_mutable_runtime().plates_1_2 == [2.0]
        assert get_plate_mutable_runtime().plates_1_2 == [1.0]

    assert outer.plates.plates_1_2 == [1.0]
    assert inner.plates.plates_1_2 == [2.0]


def test_parallel_nested_bound_async_tasks_do_not_share_runtime() -> None:
    async def exercise(marker: float) -> float:
        ctx = PlateOrderContext.fresh_empty()
        with ctx.bound():
            ctx.plates.plates_1_2.append(marker)
            await asyncio.sleep(0.02)
            return get_plate_mutable_runtime().plates_1_2[0]

    async def _gather() -> tuple[float, float]:
        return await asyncio.gather(exercise(3.33), exercise(4.44))

    a, b = asyncio.run(_gather())
    assert a == pytest.approx(3.33)
    assert b == pytest.approx(4.44)
