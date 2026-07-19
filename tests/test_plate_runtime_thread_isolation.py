"""PLATE-CTX-001: мутабельное состояние заказа не пересекается между потоками."""

from __future__ import annotations

import threading

import core.config_and_data as cfg


def test_plate_load_details_isolated_across_threads() -> None:
    barrier = threading.Barrier(2)
    ok: dict[int, bool] = {}

    def worker(tid: int, text: str, expected_total: int) -> None:
        cfg.set_plate_lists_from_text(text)
        barrier.wait()
        total = sum(cfg.PLATE_LOAD_DETAILS.values())
        ok[tid] = total == expected_total

    t1 = threading.Thread(target=worker, args=(1, "ПБ 78-12-8п 2", 2))
    t2 = threading.Thread(target=worker, args=(2, "ПБ 78-12-8п 5", 5))
    t1.start()
    t2.start()
    t1.join()
    t2.join()
    assert ok[1] and ok[2], f"expected per-thread totals, got {ok}"
