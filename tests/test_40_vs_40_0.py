#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Проверяет, что ПБ 40-12-8п и ПБ 40,0-12-8п НЕ объединяются."""
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as cfg
from core.plate_order_context import PlateOrderContext
from viz_modules.procurement import build_procurement_items


def test_40_vs_40_0_not_merged():
    # Изоляция от чужого OPT_* из других тестов в том же процессе pytest
    import core.optimization as optimization

    optimization.OPT_CASCADING_PLAN.clear()
    optimization.OPT_CASCADING_PLAN_BY_LOAD.clear()

    test_order = """
Плиты ПБ 40-12-8п 4 шт
Плиты ПБ 40,0-12-8п 6 шт
"""
    ctx = PlateOrderContext.fresh_empty()
    with ctx.bound():
        cfg.set_plate_lists_from_text(test_order)
        load_details = dict(ctx.plates.plate_load_details)
        length_raw = dict(ctx.plates.plate_length_dm_raw)
        items = build_procurement_items()

    print("PLATE_LOAD_DETAILS:")
    for key, qty in sorted(load_details.items()):
        print(f"  key={key}, qty={qty}")

    print("\nPLATE_LENGTH_DM_RAW:")
    for key, raw in length_raw.items():
        print(f"  key={key}, raw={raw!r}")

    assert len(load_details) == 2, (
        f"Ожидалось 2 позиции в PLATE_LOAD_DETAILS, получили {len(load_details)}: "
        f"{dict(load_details)}"
    )

    keys = list(load_details.keys())
    ldrs = sorted(k[3] for k in keys)
    assert "40" in ldrs and "40,0" in ldrs, f"Ожидались ключи с ldr '40' и '40,0', получили: {ldrs}"
    print(f"\nProcurement items ({len(items)}):")
    for it in items:
        print(f"  qty={it['qty']}, length={it['length']}, ldr={it.get('length_dm_raw')!r}")

    items_4m = [it for it in items if abs(it["length"] - 4.0) < 0.01 and abs(it["width"] - 1.2) < 0.01]
    assert len(items_4m) == 2, (
        f"Ожидалось 2 позиции 4.0×1.2 в закупке, получили {len(items_4m)}: {items_4m}"
    )

    qtys = sorted(it["qty"] for it in items_4m)
    assert qtys == [4, 6], f"Ожидались qty [4, 6], получили {qtys}"

    print("\nOK: ПБ 40-12-8п (4 шт) и ПБ 40,0-12-8п (6 шт) — две отдельные позиции!")


if __name__ == "__main__":
    test_40_vs_40_0_not_merged()
    print("SUCCESS!")
