# -*- coding: utf-8 -*-
"""Построение карты армирования для последовательности через get_reinforcement(..., db_path=...)."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import TYPE_CHECKING, Any

from viz_modules.layout_sequence.debug_trace import append_json_line
from viz_modules.layout_sequence.helpers import canonical_load_code

if TYPE_CHECKING:
    from viz_modules.layout_sequence.deps import LayoutSequenceDeps


def build_reinforcement_map_for_sequence(
    *,
    plate_load_details: Mapping | None,
    OPT_CASCADING_PLAN_BY_LOAD: dict[Any, dict] | None,
    OPT_CASCADING_PLAN: dict[str, Any] | None,
    get_reinforcement: Callable[..., float | None],
    db_path: Any,
    normalize_load_code: Callable[..., Any],
    deps: LayoutSequenceDeps,
) -> dict[tuple[Any, Any, Any], float]:
    """
    Глобальная карта армирования {(length, width_mm, lc_canonical): reinforcement}
    из PLATE_LOAD_DETAILS и дополнения из primary_cuts планов.
    """
    reinforcement_map: dict[tuple[Any, Any, Any], float] = {}

    debug_log_path = deps.traces.debug_log
    log = deps.log

    if plate_load_details:
        log.info(f"[VISUAL] Начинаем создание карты армирования из {len(plate_load_details)} записей")
        for key, _qty in plate_load_details.items():
            length, width_m, load_code = key[0], key[1], key[2]
            width_mm = int(round(width_m * 1000))
            reinforcement = get_reinforcement(
                length_m=length,
                load_code=load_code,
                source="series",
                db_path=db_path,
                allow_fallback=True,
            )
            lc_canonical = canonical_load_code(load_code) or 8
            if reinforcement is None and lc_canonical == 12:
                reinforcement = get_reinforcement(
                    length_m=length,
                    load_code=13,
                    source="series",
                    db_path=db_path,
                    allow_fallback=True,
                )
            rkey = (length, width_mm, lc_canonical)
            if reinforcement and reinforcement < 999:
                reinforcement_map[rkey] = reinforcement
                log.info(
                    f"[VISUAL]   Добавлено: ({length}м, {width_mm}мм, нагрузка {lc_canonical}) → армирование {reinforcement:.1f}"
                )

    def _supplement_reinforcement_map_from_plan(rmap: dict, plan: dict) -> int:
        added = 0
        for cut in plan.get("primary_cuts", []):
            length = cut["lengths"][0] if cut.get("lengths") else 6.0
            width_mm = cut["width"]
            load_code = normalize_load_code(cut.get("load_code", 8))
            lc_canonical = canonical_load_code(load_code) or 8
            pmap_key = (length, width_mm, lc_canonical)
            if pmap_key not in rmap:
                reinforcement = get_reinforcement(
                    length_m=length,
                    load_code=load_code,
                    source="series",
                    db_path=db_path,
                    allow_fallback=True,
                )
                if reinforcement is None and lc_canonical == 12:
                    reinforcement = get_reinforcement(
                        length_m=length,
                        load_code=13,
                        source="series",
                        db_path=db_path,
                        allow_fallback=True,
                    )
                if reinforcement is not None and reinforcement < 999:
                    rmap[pmap_key] = reinforcement
                    added += 1
                    log.info(
                        f"[VISUAL]   Дополнено из плана: ({length}м, {width_mm}мм, нагрузка {lc_canonical}) → армирование {reinforcement:.1f}"
                    )
                    if abs(length - 7.1) < 0.05 and lc_canonical in (12, 13):
                        try:
                            append_json_line(
                                debug_log_path,
                                {
                                    "hypothesisId": "H71supp",
                                    "location": "layout_sequence:_supplement",
                                    "message": "71-12п дополнение карты",
                                    "data": {
                                        "length": length,
                                        "width_mm": width_mm,
                                        "lc_canonical": lc_canonical,
                                        "reinforcement": reinforcement,
                                        "raw_load_code": load_code,
                                    },
                                    "timestamp": __import__("time").time() * 1000,
                                },
                            )
                        except Exception:
                            pass
        return added

    if OPT_CASCADING_PLAN_BY_LOAD:
        for plan in OPT_CASCADING_PLAN_BY_LOAD.values():
            _supplement_reinforcement_map_from_plan(reinforcement_map, plan)
    elif OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get("primary_cuts"):
        _supplement_reinforcement_map_from_plan(reinforcement_map, OPT_CASCADING_PLAN)

    log.info(f"[VISUAL] Создана карта армирования: {len(reinforcement_map)} записей")
    return reinforcement_map
