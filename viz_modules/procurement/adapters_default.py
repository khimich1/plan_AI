from __future__ import annotations

import core.config_and_data as cfg
from core.price_db import get_price
from core.raw_material_db import get_raw_material_cost
from core.reinforcement_db import get_reinforcement

from .ports import ProcurementDeps


def default_procurement_deps() -> ProcurementDeps:
    return ProcurementDeps(
        db_path=cfg.PRICE_DB_PATH,
        get_price=get_price,
        get_raw_material_cost=get_raw_material_cost,
        get_reinforcement=get_reinforcement,
    )
