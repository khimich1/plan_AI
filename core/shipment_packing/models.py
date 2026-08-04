"""Dataclasses движка укладки."""

from __future__ import annotations

from dataclasses import dataclass, field

from core.shipment_packing.reasons import NotFitReason, WarningCode


@dataclass(frozen=True)
class VehicleLimits:
    max_weight_kg: float = 19_800.0
    body_length_m: float = 13.2
    max_tiers: int = 4


@dataclass(frozen=True)
class PlateCandidate:
    completed_plate_id: int
    kp_id: int
    plate_name: str
    length_m: float | None
    width_m: float | None
    load_class: int | None
    qty: int
    unit_weight_kg: float
    completed_date: str | None = None


@dataclass
class PlateUnit:
    """Одна штука из развёрнутого qty."""

    candidate_key: int
    unit_index: int
    completed_plate_id: int
    kp_id: int
    plate_name: str
    length_m: float | None
    width_m: float | None
    load_class: int | None
    unit_weight_kg: float
    completed_date: str | None
    marking_length_m: float
    marking_fallback: bool = False


@dataclass
class PackedUnit:
    unit: PlateUnit
    suboptimal_piece: bool = False


@dataclass
class RejectedUnit:
    unit: PlateUnit
    reason: NotFitReason
    qty: int = 1


@dataclass
class PackWarning:
    code: WarningCode
    message: str
    kp_ids: list[int] = field(default_factory=list)


@dataclass
class AggregatedLine:
    completed_plate_id: int
    kp_id: int
    plate_name: str
    length_m: float | None
    width_m: float | None
    load_class: int | None
    qty: int
    available_qty: int
    unit_weight_kg: float
    weight_kg: float
    completed_date: str | None = None
    reason_code: str | None = None
    reason_text: str | None = None


@dataclass
class OrderRemainderLine:
    completed_plate_id: int
    kp_id: int
    plate_name: str
    qty_remaining: int


@dataclass(frozen=True)
class LayoutUnit:
    completed_plate_id: int
    kp_id: int
    plate_name: str
    width_m: float | None


@dataclass(frozen=True)
class LayoutTier:
    index: int
    units: list[LayoutUnit] = field(default_factory=list)


@dataclass(frozen=True)
class LayoutStack:
    index: int
    marking_length_m: float
    tiers: list[LayoutTier] = field(default_factory=list)


@dataclass(frozen=True)
class LoadingStep:
    step: int
    stack_index: int
    tier_index: int
    description: str


@dataclass(frozen=True)
class LayoutMetadata:
    body_length_m: float
    body_used_m: float
    stacks: list[LayoutStack] = field(default_factory=list)
    loading_steps: list[LoadingStep] = field(default_factory=list)


@dataclass
class PackResult:
    items: list[AggregatedLine]
    not_fit: list[AggregatedLine]
    order_remainder: list[OrderRemainderLine]
    warnings: list[PackWarning]
    total_weight_kg: float
    layout: LayoutMetadata | None = None
