"""GSM waybill domain: pure balance math, transaction parse, and frozen DTOs (no I/O)."""

from core.gsm.balance import (
    BalanceViolation,
    apply_day,
    apply_day_chain,
    burn_for_km,
)
from core.gsm.geo import GeoPoint
from core.gsm.models import (
    Anchor,
    LegPlan,
    RouteRef,
    TankState,
    Transaction,
    WaybillDay,
)
from core.gsm.transactions import (
    ParsedTxFile,
    ParsedTxRow,
    classify_service,
    parse_transactions_xls,
)

__all__ = [
    "Anchor",
    "BalanceViolation",
    "GeoPoint",
    "LegPlan",
    "ParsedTxFile",
    "ParsedTxRow",
    "RouteRef",
    "TankState",
    "Transaction",
    "WaybillDay",
    "apply_day",
    "apply_day_chain",
    "burn_for_km",
    "classify_service",
    "parse_transactions_xls",
]
