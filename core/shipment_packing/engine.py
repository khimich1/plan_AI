"""Жадный алгоритм набора рейса по правилам укладки ПБ."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field

from core.shipment_packing.marking import marking_length_m
from core.shipment_packing.models import (
    AggregatedLine,
    LayoutMetadata,
    LayoutStack,
    LayoutTier,
    LayoutUnit,
    LoadingStep,
    OrderRemainderLine,
    PackResult,
    PackWarning,
    PackedUnit,
    PlateCandidate,
    PlateUnit,
    RejectedUnit,
    VehicleLimits,
)
from core.shipment_packing.reasons import (
    REASON_TEXT,
    WARNING_TEXT,
    NotFitReason,
    WarningCode,
)
from core.shipment_packing.rules import (
    body_length_for_stacks,
    gost_stack_count,
    is_piece,
    markings_compatible,
)


@dataclass
class _Tier:
    units: list[PlateUnit] = field(default_factory=list)

    def width_values(self) -> set[float]:
        return {round(float(u.width_m or 0.0), 3) for u in self.units}

    def full_count(self) -> int:
        return sum(1 for u in self.units if not is_piece(u.width_m))

    def piece_count(self) -> int:
        return sum(1 for u in self.units if is_piece(u.width_m))

    def can_add(self, unit: PlateUnit) -> bool:
        widths = self.width_values()
        unit_w = round(float(unit.width_m or 0.0), 3)
        if widths and unit_w not in widths and len(widths) >= 2:
            return False
        if is_piece(unit.width_m):
            if self.full_count() >= 2:
                return False
            return self.piece_count() < 4 and (self.full_count() + self.piece_count()) < 4
        if self.piece_count() > 0:
            return self.full_count() < 1
        return self.full_count() < 2


@dataclass
class _Stack:
    tiers: list[_Tier] = field(default_factory=list)

    @property
    def max_marking(self) -> float:
        if not self.tiers:
            return 0.0
        return max(
            u.marking_length_m for tier in self.tiers for u in tier.units
        )

    def tier_count(self) -> int:
        return len(self.tiers)

    def unit_count(self) -> int:
        return sum(len(t.units) for t in self.tiers)

    def find_tier_for(self, unit: PlateUnit) -> _Tier | None:
        for tier in self.tiers:
            if tier.can_add(unit):
                return tier
        return None

    def add_unit(self, unit: PlateUnit) -> None:
        tier = self.find_tier_for(unit)
        if tier is None:
            tier = _Tier()
            self.tiers.append(tier)
        tier.units.append(unit)

    def can_add(self, unit: PlateUnit, *, max_tiers: int) -> bool:
        tier = self.find_tier_for(unit)
        if tier is not None:
            return True
        return self.tier_count() < max_tiers


@dataclass
class _Layout:
    limits: VehicleLimits
    stacks: list[_Stack] = field(default_factory=list)
    packed: list[PackedUnit] = field(default_factory=list)
    rejected: list[RejectedUnit] = field(default_factory=list)
    total_weight_kg: float = 0.0

    def _stack_markings(self) -> list[float]:
        return [s.max_marking for s in self.stacks if s.max_marking > 0]

    def _body_used(self) -> float:
        return body_length_for_stacks(self._stack_markings())

    def _project_body(self, stack: _Stack, unit: PlateUnit) -> float:
        markings = self._stack_markings()
        if stack in self.stacks and stack.max_marking > 0:
            idx = self.stacks.index(stack)
            new_max = max(stack.max_marking, unit.marking_length_m)
            if idx < len(markings):
                markings[idx] = new_max
            return body_length_for_stacks(markings)
        new_mark = max(stack.max_marking, unit.marking_length_m) if stack.tiers else unit.marking_length_m
        return body_length_for_stacks(markings + [new_mark])

    def _max_stacks_for(self, marking: float) -> int:
        return gost_stack_count(marking)

    def _compatible_stack(self, unit: PlateUnit) -> _Stack | None:
        for stack in self.stacks:
            if not stack.tiers:
                continue
            ref = stack.max_marking
            if markings_compatible(ref, unit.marking_length_m) and stack.can_add(
                unit, max_tiers=self.limits.max_tiers
            ):
                return stack
        return None

    def _reject(self, unit: PlateUnit, reason: NotFitReason) -> None:
        self.rejected.append(RejectedUnit(unit=unit, reason=reason))

    def try_add(self, unit: PlateUnit, *, allow_new_stack: bool = True) -> bool:
        if self.total_weight_kg + unit.unit_weight_kg > self.limits.max_weight_kg + 1e-9:
            self._reject(unit, NotFitReason.WEIGHT_LIMIT)
            return False

        stack = self._compatible_stack(unit)
        if stack is None and allow_new_stack:
            max_stacks = self._max_stacks_for(unit.marking_length_m)
            if len(self.stacks) >= max_stacks:
                self._reject(unit, NotFitReason.BODY_LENGTH)
                return False
            stack = _Stack()
            if self._project_body(stack, unit) > self.limits.body_length_m + 1e-9:
                self._reject(unit, NotFitReason.BODY_LENGTH)
                return False
            self.stacks.append(stack)

        if stack is None:
            self._reject(unit, NotFitReason.LENGTH_MIX)
            return False

        if not stack.can_add(unit, max_tiers=self.limits.max_tiers):
            self._reject(unit, NotFitReason.TIER_LIMIT)
            return False

        if self._project_body(stack, unit) > self.limits.body_length_m + 1e-9:
            self._reject(unit, NotFitReason.BODY_LENGTH)
            return False

        stack.add_unit(unit)
        self.total_weight_kg += unit.unit_weight_kg
        self.packed.append(PackedUnit(unit=unit))
        return True


def _expand_candidates(candidates: list[PlateCandidate]) -> tuple[list[PlateUnit], list[PackWarning]]:
    units: list[PlateUnit] = []
    warnings: list[PackWarning] = []
    fallback_ids: set[int] = set()

    for cand in candidates:
        marking, used_fallback = marking_length_m(cand.plate_name, cand.length_m)
        if used_fallback:
            fallback_ids.add(cand.completed_plate_id)
        for idx in range(int(cand.qty)):
            units.append(
                PlateUnit(
                    candidate_key=cand.completed_plate_id,
                    unit_index=idx,
                    completed_plate_id=cand.completed_plate_id,
                    kp_id=cand.kp_id,
                    plate_name=cand.plate_name,
                    length_m=cand.length_m,
                    width_m=cand.width_m,
                    load_class=cand.load_class,
                    unit_weight_kg=cand.unit_weight_kg,
                    completed_date=cand.completed_date,
                    marking_length_m=marking,
                    marking_fallback=used_fallback,
                )
            )

    if fallback_ids:
        warnings.append(
            PackWarning(
                code=WarningCode.MARKING_FALLBACK,
                message=WARNING_TEXT[WarningCode.MARKING_FALLBACK],
            )
        )
    return units, warnings


def _sort_full_units(units: list[PlateUnit]) -> list[PlateUnit]:
    return sorted(
        units,
        key=lambda u: (
            -u.marking_length_m,
            u.completed_date or "",
            u.completed_plate_id,
            u.unit_index,
        ),
    )


def _piece_priority(unit: PlateUnit, layout: _Layout) -> tuple[int, str, int, int]:
    """Меньше — выше приоритет. Возвращает tie-breakers для сортировки."""
    if not layout.packed:
        return (99, unit.completed_date or "", unit.completed_plate_id, unit.unit_index)

    packed_markings = {p.unit.marking_length_m for p in layout.packed}
    if unit.marking_length_m in packed_markings:
        return (0, unit.completed_date or "", unit.completed_plate_id, unit.unit_index)

    for m in packed_markings:
        if markings_compatible(m, unit.marking_length_m):
            return (1, unit.completed_date or "", unit.completed_plate_id, unit.unit_index)

    # Коротыши в остаток длины кузова
    used = layout._body_used()
    remaining = layout.limits.body_length_m - used
    if unit.marking_length_m <= remaining + 1e-9:
        return (2, unit.completed_date or "", unit.completed_plate_id, unit.unit_index)

    return (3, unit.completed_date or "", unit.completed_plate_id, unit.unit_index)


def _aggregate_lines(
    grouped: dict[int, list[PlateUnit]],
    candidates_by_id: dict[int, PlateCandidate],
    *,
    reason_by_id: dict[int, NotFitReason] | None = None,
) -> list[AggregatedLine]:
    lines: list[AggregatedLine] = []
    reason_by_id = reason_by_id or {}
    for cp_id in sorted(grouped):
        units = grouped[cp_id]
        cand = candidates_by_id[cp_id]
        qty = len(units)
        unit_w = cand.unit_weight_kg
        reason = reason_by_id.get(cp_id)
        lines.append(
            AggregatedLine(
                completed_plate_id=cp_id,
                kp_id=cand.kp_id,
                plate_name=cand.plate_name,
                length_m=cand.length_m,
                width_m=cand.width_m,
                load_class=cand.load_class,
                qty=qty,
                available_qty=cand.qty,
                unit_weight_kg=unit_w,
                weight_kg=unit_w * qty,
                completed_date=cand.completed_date,
                reason_code=reason.value if reason else None,
                reason_text=REASON_TEXT[reason] if reason else None,
            )
        )
    return lines


def _tier_description(units: list[PlateUnit]) -> str:
    counts: dict[str, int] = {}
    for unit in units:
        counts[unit.plate_name] = counts.get(unit.plate_name, 0) + 1
    parts = [f"{name} ×{qty}" if qty > 1 else name for name, qty in counts.items()]
    return " + ".join(parts)


def _build_layout_metadata(layout: _Layout, limits: VehicleLimits) -> LayoutMetadata:
    stacks: list[LayoutStack] = []
    steps: list[LoadingStep] = []
    for stack_index, stack in enumerate(layout.stacks, start=1):
        tiers: list[LayoutTier] = []
        for tier_index, tier in enumerate(stack.tiers, start=1):
            units = [
                LayoutUnit(
                    completed_plate_id=u.completed_plate_id,
                    kp_id=u.kp_id,
                    plate_name=u.plate_name,
                    width_m=u.width_m,
                )
                for u in tier.units
            ]
            tiers.append(LayoutTier(index=tier_index, units=units))
            if units:
                steps.append(
                    LoadingStep(
                        step=len(steps) + 1,
                        stack_index=stack_index,
                        tier_index=tier_index,
                        description=_tier_description(tier.units),
                    )
                )
        stacks.append(
            LayoutStack(
                index=stack_index,
                marking_length_m=stack.max_marking,
                tiers=tiers,
            )
        )
    return LayoutMetadata(
        body_length_m=limits.body_length_m,
        body_used_m=round(layout._body_used(), 3),
        stacks=stacks,
        loading_steps=steps,
    )


def pack_shipment(
    candidates: list[PlateCandidate],
    *,
    limits: VehicleLimits | None = None,
) -> PackResult:
    """Pure: без I/O. Все qty разложены: items + not_fit + remainder."""
    effective = limits or VehicleLimits()
    candidates_by_id = {c.completed_plate_id: c for c in candidates}
    all_units, warnings = _expand_candidates(candidates)

    full_units = [u for u in all_units if not is_piece(u.width_m)]
    piece_units = [u for u in all_units if is_piece(u.width_m)]

    layout = _Layout(limits=effective)

    for unit in _sort_full_units(full_units):
        layout.try_add(unit)

    has_full_plates = bool(layout.packed)

    pending_pieces = sorted(piece_units, key=lambda u: _piece_priority(u, layout))
    for unit in pending_pieces:
        if not has_full_plates:
            layout._reject(unit, NotFitReason.PIECE_PRIORITY)
            continue
        priority = _piece_priority(unit, layout)[0]
        suboptimal = priority > 0
        if priority >= 3:
            layout._reject(unit, NotFitReason.PIECE_PRIORITY)
            continue
        if layout.try_add(unit):
            if suboptimal:
                layout.packed[-1].suboptimal_piece = True
        # try_add already records rejection

    fitted_by_id: dict[int, list[PlateUnit]] = defaultdict(list)
    for packed in layout.packed:
        fitted_by_id[packed.unit.completed_plate_id].append(packed.unit)

    rejected_by_id: dict[int, list[RejectedUnit]] = defaultdict(list)
    for rej in layout.rejected:
        rejected_by_id[rej.unit.completed_plate_id].append(rej)

    # Primary reason per rejected line (most common)
    reason_by_id: dict[int, NotFitReason] = {}
    for cp_id, rejs in rejected_by_id.items():
        reason_by_id[cp_id] = rejs[0].reason

    items = _aggregate_lines(fitted_by_id, candidates_by_id)
    not_fit = _aggregate_lines(
        {cp_id: [r.unit for r in rejs] for cp_id, rejs in rejected_by_id.items()},
        candidates_by_id,
        reason_by_id=reason_by_id,
    )

    fitted_qty = {cp_id: len(units) for cp_id, units in fitted_by_id.items()}
    order_remainder = [
        OrderRemainderLine(
            completed_plate_id=c.completed_plate_id,
            kp_id=c.kp_id,
            plate_name=c.plate_name,
            qty_remaining=c.qty - fitted_qty.get(c.completed_plate_id, 0),
        )
        for c in candidates
        if c.qty - fitted_qty.get(c.completed_plate_id, 0) > 0
    ]

    kp_ids_in_items = {line.kp_id for line in items}
    if len(kp_ids_in_items) >= 2:
        warnings.append(
            PackWarning(
                code=WarningCode.KP_MIX,
                message=WARNING_TEXT[WarningCode.KP_MIX],
                kp_ids=sorted(kp_ids_in_items),
            )
        )

    if any(p.suboptimal_piece for p in layout.packed):
        warnings.append(
            PackWarning(
                code=WarningCode.PIECE_SUBOPTIMAL,
                message=WARNING_TEXT[WarningCode.PIECE_SUBOPTIMAL],
            )
        )

    return PackResult(
        items=items,
        not_fit=not_fit,
        order_remainder=order_remainder,
        warnings=warnings,
        total_weight_kg=layout.total_weight_kg,
        layout=_build_layout_metadata(layout, effective),
    )
