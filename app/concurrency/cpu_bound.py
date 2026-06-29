from __future__ import annotations

import asyncio
import os
from collections.abc import Callable
from typing import TypeVar

from core.plate_order_context import PlateOrderContext, run_in_order_context

T = TypeVar("T")

_DEFAULT_MAX_CONCURRENT = 2
_semaphore: asyncio.Semaphore | None = None


def _max_concurrent() -> int:
    raw = os.environ.get("CPU_BOUND_MAX_CONCURRENT", str(_DEFAULT_MAX_CONCURRENT))
    try:
        value = int(raw)
    except ValueError:
        value = _DEFAULT_MAX_CONCURRENT
    return max(1, value)


def _get_semaphore() -> asyncio.Semaphore:
    global _semaphore
    if _semaphore is None:
        _semaphore = asyncio.Semaphore(_max_concurrent())
    return _semaphore


async def run_cpu_bound(
    fn: Callable[[], T],
    *,
    plate_order_ctx: PlateOrderContext | None = None,
) -> T:
    """Run sync CPU work in a thread pool with optional concurrency cap."""
    async with _get_semaphore():
        if plate_order_ctx is not None:
            return await run_in_order_context(plate_order_ctx, fn)
        return await asyncio.to_thread(fn)
