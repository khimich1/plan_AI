from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import Any

import core.config_and_data as cfg
from core.optimization.layout_runtime_snapshot import _make_get_load_code_for_plate

LoadCodeFn = Callable[[float, float, int], int]
PlateLoadDetails = Mapping[tuple[float, float, Any, str], int]


def resolve_procurement_load_context(
    *,
    plate_load_details: PlateLoadDetails | None = None,
    get_load_code: LoadCodeFn | None = None,
) -> tuple[PlateLoadDetails, LoadCodeFn]:
    """Resolve load map and lookup callback for procurement builders.

  Explicit ``plate_load_details`` / ``get_load_code`` take priority over TLS globals.
    """
    if get_load_code is not None:
        details: PlateLoadDetails = (
            plate_load_details if plate_load_details is not None else {}
        )
        return details, get_load_code
    if plate_load_details is not None:
        return plate_load_details, _make_get_load_code_for_plate(plate_load_details)
    return cfg.PLATE_LOAD_DETAILS, cfg.get_load_code_for_plate
