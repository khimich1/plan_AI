from __future__ import annotations

import math

from core.config.constants import LONG_CUT_PRICE_PER_M, MIN_BILLABLE_TRIM_MM

from ..price_utils import find_price_for_plate
from .plan_lookup import _is_same_length
from .ports import ProcurementDeps, resolve_procurement_deps


def _cut_load_key(cut: dict) -> int | None:
    """Нормализованная нагрузка на строке реза; None — легаси-план без load_code."""
    lc = cut.get('load_code')
    if lc is None or lc == '':
        return None
    return int(math.floor(float(lc)))


def _cut_matches_load(cut: dict, load_key: int) -> bool:
    """True, если рез относится к той же нагрузке, что строка заказа (или load_code не задан)."""
    cut_lc = _cut_load_key(cut)
    if cut_lc is None:
        return True
    return cut_lc == load_key


def _cut_length_from_lengths(lengths: list | None, default_length: float) -> float:
    if lengths:
        try:
            return float(lengths[0])
        except (TypeError, ValueError, IndexError):
            pass
    return default_length


def _width_matches_cut(width_mm: int, sec_cuts: list) -> bool:
    return any(abs(width_mm - int(cut_width)) <= 20 for cut_width in sec_cuts)


def _secondary_output_width_mm(sec_cut: dict) -> int:
    cuts = sec_cut.get('cuts', []) or []
    if not cuts:
        return 0
    return int(cuts[0])


def _used_rest_strip_qty_legacy(
    current_plan: dict | None,
    rest_mm: int,
    prim_len: float,
) -> int:
    """Сколько полос rest_mm использовано (по qty secondary), легаси без parent_instance_id."""
    if not current_plan:
        return 0
    used = 0
    for sec_cut in current_plan.get('secondary_cuts') or []:
        if int(sec_cut.get('source', 0) or 0) != rest_mm:
            continue
        if not _is_same_length(sec_cut.get('source_lengths', []), prim_len):
            continue
        used += int(sec_cut.get('qty', 0) or 0)
    return used


def _consumed_width_mm_on_rest_strip(
    current_plan: dict | None,
    rest_mm: int,
    prim_len: float,
    *,
    parent_instance_id: str | None = None,
) -> int:
    """Суммарная ширина продукции с полосы rest_mm (все нагрузки — кросс-каскад).

    Если задан ``parent_instance_id``, учитываются только secondary, привязанные к
    этому первичному слэбу (иначе при нескольких слэбах с одинаковым rest все
    secondary ошибочно суммируются).
    """
    if not current_plan:
        return 0
    total = 0
    for sec_cut in current_plan.get('secondary_cuts') or []:
        if int(sec_cut.get('source', 0) or 0) != rest_mm:
            continue
        if not _is_same_length(sec_cut.get('source_lengths', []), prim_len):
            continue
        if parent_instance_id:
            sec_parent = sec_cut.get('parent_instance_id')
            if sec_parent and sec_parent != parent_instance_id:
                continue
        out_w = _secondary_output_width_mm(sec_cut)
        total += out_w * int(sec_cut.get('qty', 0) or 0)
    return total


def _cascade_qty_on_own_rest(
    current_plan: dict | None,
    rest_groups: dict[tuple[int, float], int],
    load_key: int,
    width_mm: int,
    length: float,
) -> int:
    """Кол-во плит, полученных каскадом с полос остатка своего primary."""
    if not current_plan or not rest_groups:
        return 0
    total = 0
    for sec_cut in current_plan.get('secondary_cuts') or []:
        if not _cut_matches_load(sec_cut, load_key):
            continue
        if not _secondary_matches_primary_rest(sec_cut, rest_groups):
            continue
        if not _is_same_length(sec_cut.get('lengths', []), length):
            continue
        sec_cuts = sec_cut.get('cuts', []) or []
        if not _width_matches_cut(width_mm, sec_cuts):
            continue
        total += int(sec_cut.get('qty', 0) or 0)
    return total


def _has_same_load_cascade_on_rest(
    current_plan: dict | None,
    rest_groups: dict[tuple[int, float], int],
    load_key: int,
    width_mm: int,
    length: float,
) -> bool:
    return _cascade_qty_on_own_rest(
        current_plan, rest_groups, load_key, width_mm, length
    ) > 0


def _longitudinal_cuts_for_rest_secondary(
    sec_cut: dict,
    *,
    min_one_cut_per_op: bool = False,
) -> int:
    """Резы на полосе остатка."""
    sec_qty = int(sec_cut.get('qty', 0) or 0)
    if sec_qty <= 0:
        return 0
    sec_pieces = int(sec_cut.get('pieces', 1) or 1)
    sec_cuts_list = sec_cut.get('cuts', []) or []
    kept_pieces = sec_pieces + max(0, len(sec_cuts_list) - 1)
    waste_w_mm = float(sec_cut.get('waste', 0) or 0)
    internal_cuts = (kept_pieces - 1) + (
        1 if waste_w_mm > MIN_BILLABLE_TRIM_MM else 0
    )
    if min_one_cut_per_op and int(sec_cut.get('source', 0) or 0) > MIN_BILLABLE_TRIM_MM:
        return sec_qty * max(1, internal_cuts)
    return sec_qty * internal_cuts


def _transverse_remainder_unit_cost(
    *,
    base_price_1_2m: float,
    width_mm: int,
    remainder_m: float,
    product_length_m: float,
) -> float:
    if remainder_m <= 0.01 or product_length_m <= 0 or base_price_1_2m <= 0:
        return 0.0
    return base_price_1_2m * (width_mm / 1200.0) * (remainder_m / product_length_m)


def _append_transverse_remainder_term(
    terms: list[tuple[float, int]],
    remainder_m: float,
    count: int,
) -> None:
    if remainder_m <= 0.01 or count <= 0:
        return
    rem_rounded = round(remainder_m, 2)
    for idx, (existing_rem, existing_n) in enumerate(terms):
        if abs(existing_rem - rem_rounded) < 0.01:
            terms[idx] = (existing_rem, existing_n + count)
            return
    terms.append((rem_rounded, count))


def _apply_transverse_remainder_from_cut(
    *,
    src_len: float,
    product_length: float,
    sec_qty: int,
    qty: int,
    base_price_1_2m: float,
    width_mm: int,
    transverse_remainder_cost: float,
    transverse_remainder_terms: list[tuple[float, int]],
    trans_cuts: float,
    count_transverse_op: bool,
) -> tuple[float, list[tuple[float, int]], float]:
    if abs(src_len - product_length) <= 0.05:
        return transverse_remainder_cost, transverse_remainder_terms, trans_cuts

    if count_transverse_op and qty > 0:
        trans_cuts += (1.0 * sec_qty) / qty

    remainder_m = src_len - product_length
    if remainder_m > 0.01 and qty > 0:
        piece_cost = _transverse_remainder_unit_cost(
            base_price_1_2m=base_price_1_2m,
            width_mm=width_mm,
            remainder_m=remainder_m,
            product_length_m=product_length,
        )
        if piece_cost > 0:
            transverse_remainder_cost += piece_cost * (sec_qty / qty)
            _append_transverse_remainder_term(transverse_remainder_terms, remainder_m, sec_qty)

    return transverse_remainder_cost, transverse_remainder_terms, trans_cuts


def _apply_secondary_cut(
    sec_cut: dict,
    *,
    length: float,
    width_mm: int,
    qty: int,
    base_price_1_2m: float,
    base_price: float,
    load_code: int,
    price_table: dict,
    deps: ProcurementDeps,
    total_cuts_for_this_size: int,
    total_plates_from_cuts: int,
    long_cut_meterage: float,
    waste_cost: float,
    waste_terms: list,
    trans_cuts: float,
    transverse_remainder_cost: float,
    transverse_remainder_terms: list[tuple[float, int]],
    charge_strip_waste: bool = True,
    min_one_cut_per_op: bool | None = None,
) -> tuple[int, int, float, float, list, float, float, list[tuple[float, int]]]:
    del base_price, load_code, price_table, deps  # kept for call-site compatibility

    sec_qty = int(sec_cut.get('qty', 0) or 0)
    sec_pieces = int(sec_cut.get('pieces', 1) or 1)
    sec_cuts_list = sec_cut.get('cuts', []) or []
    kept_pieces = sec_pieces + max(0, len(sec_cuts_list) - 1)

    src_lens = sec_cut.get('source_lengths', []) or []
    src_len = float(src_lens[0]) if src_lens else length

    waste_w_mm = float(sec_cut.get('waste', 0) or 0)
    if min_one_cut_per_op is None:
        min_one_cut_per_op = not charge_strip_waste
    cut_count = _longitudinal_cuts_for_rest_secondary(
        sec_cut,
        min_one_cut_per_op=min_one_cut_per_op,
    )

    total_plates_from_cuts += sec_qty * kept_pieces
    total_cuts_for_this_size += cut_count
    long_cut_meterage += cut_count * src_len

    if (
        charge_strip_waste
        and waste_w_mm > MIN_BILLABLE_TRIM_MM
        and base_price_1_2m > 0
        and qty > 0
    ):
        cost_of_waste_piece = (waste_w_mm / 1200.0) * base_price_1_2m
        waste_cost += (cost_of_waste_piece * sec_qty) / qty
        waste_terms.append((waste_w_mm, sec_qty))

    if src_lens:
        is_transverse = sec_cut.get('type') == 'transverse' or abs(src_len - length) > 0.05
        if is_transverse:
            transverse_remainder_cost, transverse_remainder_terms, trans_cuts = (
                _apply_transverse_remainder_from_cut(
                    src_len=src_len,
                    product_length=length,
                    sec_qty=sec_qty,
                    qty=qty,
                    base_price_1_2m=base_price_1_2m,
                    width_mm=width_mm,
                    transverse_remainder_cost=transverse_remainder_cost,
                    transverse_remainder_terms=transverse_remainder_terms,
                    trans_cuts=trans_cuts,
                    count_transverse_op=True,
                )
            )

    return (
        total_cuts_for_this_size,
        total_plates_from_cuts,
        long_cut_meterage,
        waste_cost,
        waste_terms,
        trans_cuts,
        transverse_remainder_cost,
        transverse_remainder_terms,
    )


def _secondary_matches_primary_rest(
    sec_cut: dict,
    rest_groups: dict[tuple[int, float], int],
) -> bool:
    source = int(sec_cut.get('source', 0) or 0)
    src_lens = sec_cut.get('source_lengths', []) or []
    for (rest_mm, prim_len) in rest_groups:
        if source != rest_mm:
            continue
        if _is_same_length(src_lens, prim_len):
            return True
    return False


def _rest_groups_from_plan(
    current_plan: dict | None,
    load_key: int,
) -> dict[tuple[int, float], int]:
    """Собирает rest_groups из всех primary_cuts плана (не только matched по width)."""
    rest_groups: dict[tuple[int, float], int] = {}
    if not current_plan:
        return rest_groups
    for prim_cut in current_plan.get('primary_cuts') or []:
        if not _cut_matches_load(prim_cut, load_key):
            continue
        rest_mm = int(prim_cut.get('rest', 0) or 0)
        if rest_mm <= MIN_BILLABLE_TRIM_MM:
            continue
        prim_len = _cut_length_from_lengths(prim_cut.get('lengths', []), 0.0)
        prim_qty = int(prim_cut.get('qty', 0) or 0)
        key = (rest_mm, prim_len)
        rest_groups[key] = rest_groups.get(key, 0) + prim_qty
    return rest_groups


def _apply_cascade_secondary_for_primary(
    *,
    current_plan: dict,
    rest_groups: dict[tuple[int, float], int],
    matched_primary: list[tuple[int, float, int]],
    length: float,
    width_mm: int,
    qty: int,
    base_price_1_2m: float,
    base_price: float,
    load_code: int,
    price_table: dict,
    deps: ProcurementDeps,
    total_cuts_for_this_size: int,
    total_plates_from_cuts: int,
    long_cut_meterage: float,
    long_cut_length_display: float,
    waste_cost: float,
    waste_terms: list,
    trans_cuts: float,
    transverse_remainder_cost: float,
    transverse_remainder_terms: list[tuple[float, int]],
) -> tuple[int, int, float, float, float, list, float, float, list[tuple[float, int]]]:
    """
    Secondary-резы из остатков matched primary той же ширины, что строка заказа.

    Нужно, когда primary и secondary дают одну марку в одной позиции (напр. 60-5,3 × 2).
    """
    primary_plate_qty = sum(prim_qty for prim_qty, _, _ in matched_primary)
    remaining_slots = max(0, qty - primary_plate_qty)
    if remaining_slots <= 0 or not rest_groups:
        return (
            total_cuts_for_this_size,
            total_plates_from_cuts,
            long_cut_meterage,
            long_cut_length_display,
            waste_cost,
            waste_terms,
            trans_cuts,
            transverse_remainder_cost,
            transverse_remainder_terms,
        )

    load_key = int(math.floor(float(load_code)))

    for sec_cut in current_plan.get('secondary_cuts') or []:
        if remaining_slots <= 0:
            break
        if not _cut_matches_load(sec_cut, load_key):
            continue
        if not _secondary_matches_primary_rest(sec_cut, rest_groups):
            continue
        if not _is_same_length(sec_cut.get('lengths', []), length):
            continue
        sec_cuts = sec_cut.get('cuts', []) or []
        if not _width_matches_cut(width_mm, sec_cuts):
            continue

        sec_qty_available = int(sec_cut.get('qty', 0) or 0)
        effective_qty = min(sec_qty_available, remaining_slots)
        if effective_qty <= 0:
            continue

        src_lens = sec_cut.get('source_lengths', []) or []
        if src_lens:
            long_cut_length_display = float(src_lens[0])

        sec_cut_effective = {**sec_cut, 'qty': effective_qty}
        (
            total_cuts_for_this_size,
            total_plates_from_cuts,
            long_cut_meterage,
            waste_cost,
            waste_terms,
            trans_cuts,
            transverse_remainder_cost,
            transverse_remainder_terms,
        ) = _apply_secondary_cut(
            sec_cut_effective,
            length=length,
            width_mm=width_mm,
            qty=qty,
            base_price_1_2m=base_price_1_2m,
            base_price=base_price,
            load_code=load_code,
            price_table=price_table,
            deps=deps,
            total_cuts_for_this_size=total_cuts_for_this_size,
            total_plates_from_cuts=total_plates_from_cuts,
            long_cut_meterage=long_cut_meterage,
            waste_cost=waste_cost,
            waste_terms=waste_terms,
            trans_cuts=trans_cuts,
            transverse_remainder_cost=transverse_remainder_cost,
            transverse_remainder_terms=transverse_remainder_terms,
            charge_strip_waste=False,
        )
        remaining_slots -= effective_qty

    return (
        total_cuts_for_this_size,
        total_plates_from_cuts,
        long_cut_meterage,
        long_cut_length_display,
        waste_cost,
        waste_terms,
        trans_cuts,
        transverse_remainder_cost,
        transverse_remainder_terms,
    )


def _is_crossload_rest_secondary(
    sec_cut: dict,
    rest_groups: dict[tuple[int, float], int],
    load_key: int,
    current_plan: dict,
) -> bool:
    """Secondary с полосы остатка primary другой нагрузки (10п → 8п)."""
    if _secondary_matches_primary_rest(sec_cut, rest_groups):
        return False
    source = int(sec_cut.get('source', 0) or 0)
    if source <= MIN_BILLABLE_TRIM_MM:
        return False
    src_lens = sec_cut.get('source_lengths', []) or []
    for prim_cut in current_plan.get('primary_cuts') or []:
        prim_lc = _cut_load_key(prim_cut)
        if prim_lc is None or prim_lc == load_key:
            continue
        if int(prim_cut.get('rest', 0) or 0) != source:
            continue
        prim_len = _cut_length_from_lengths(prim_cut.get('lengths', []), 0.0)
        if not _is_same_length(src_lens, prim_len):
            continue
        if not _is_same_length(sec_cut.get('lengths', []), prim_len):
            continue
        return True
    return False


def _apply_crossload_rest_secondaries(
    *,
    current_plan: dict,
    rest_groups: dict[tuple[int, float], int],
    length: float,
    width_mm: int,
    qty: int,
    base_price_1_2m: float,
    base_price: float,
    load_code: int,
    price_table: dict,
    deps: ProcurementDeps,
    total_cuts_for_this_size: int,
    total_plates_from_cuts: int,
    long_cut_meterage: float,
    long_cut_length_display: float,
    waste_cost: float,
    waste_terms: list,
    trans_cuts: float,
    transverse_remainder_cost: float,
    transverse_remainder_terms: list[tuple[float, int]],
) -> tuple[int, int, float, float, float, list, float, float, list[tuple[float, int]]]:
    """
    Secondary с полосы остатка другой нагрузки (кросс-каскад): только резы/метраж,
    без отхода полосы (он на владельце-primary).
    """
    load_key = int(math.floor(float(load_code)))

    for sec_cut in current_plan.get('secondary_cuts') or []:
        if not _cut_matches_load(sec_cut, load_key):
            continue
        if not _is_crossload_rest_secondary(
            sec_cut, rest_groups, load_key, current_plan
        ):
            continue
        if sec_cut.get('type') == 'transverse':
            continue
        if not _is_same_length(sec_cut.get('lengths', []), length):
            continue
        sec_cuts = sec_cut.get('cuts', []) or []
        if not _width_matches_cut(width_mm, sec_cuts):
            continue

        src_lens = sec_cut.get('source_lengths', []) or []
        if src_lens:
            long_cut_length_display = float(src_lens[0])

        (
            total_cuts_for_this_size,
            total_plates_from_cuts,
            long_cut_meterage,
            waste_cost,
            waste_terms,
            trans_cuts,
            transverse_remainder_cost,
            transverse_remainder_terms,
        ) = _apply_secondary_cut(
            sec_cut,
            length=length,
            width_mm=width_mm,
            qty=qty,
            base_price_1_2m=base_price_1_2m,
            base_price=base_price,
            load_code=load_code,
            price_table=price_table,
            deps=deps,
            total_cuts_for_this_size=total_cuts_for_this_size,
            total_plates_from_cuts=total_plates_from_cuts,
            long_cut_meterage=long_cut_meterage,
            waste_cost=waste_cost,
            waste_terms=waste_terms,
            trans_cuts=trans_cuts,
            transverse_remainder_cost=transverse_remainder_cost,
            transverse_remainder_terms=transverse_remainder_terms,
            charge_strip_waste=False,
        )

    return (
        total_cuts_for_this_size,
        total_plates_from_cuts,
        long_cut_meterage,
        long_cut_length_display,
        waste_cost,
        waste_terms,
        trans_cuts,
        transverse_remainder_cost,
        transverse_remainder_terms,
    )


def _apply_transverse_on_primary_strip(
    *,
    current_plan: dict,
    rest_groups: dict[tuple[int, float], int],
    length: float,
    width_mm: int,
    qty: int,
    base_price_1_2m: float,
    load_code: int,
    trans_cuts: float,
    transverse_remainder_cost: float,
    transverse_remainder_terms: list[tuple[float, int]],
    total_cuts_for_this_size: int,
    total_plates_from_cuts: int,
    long_cut_meterage: float,
    long_cut_length_display: float,
) -> tuple[int, int, float, float, float, float, list[tuple[float, int]]]:
    """Поперечный рез / secondary с чужого rest той же нагрузки (не own-rest).

    Также начисляет продольный рез на rest-полосе чужой primary той же нагрузки
    (кейс ПБ 43-7,25 из ленты 8,6 м). Кросс-нагрузочные secondary уже обработаны
    в ``_apply_crossload_rest_secondaries`` (кроме pure ``transverse``).
    """
    load_key = int(math.floor(float(load_code)))

    for sec_cut in current_plan.get('secondary_cuts') or []:
        if not _cut_matches_load(sec_cut, load_key):
            continue
        if _secondary_matches_primary_rest(sec_cut, rest_groups):
            continue
        if not _is_same_length(sec_cut.get('lengths', []), length):
            continue
        sec_cuts = sec_cut.get('cuts', []) or []
        if not _width_matches_cut(width_mm, sec_cuts):
            continue

        src_lens = sec_cut.get('source_lengths', []) or []
        if not src_lens:
            continue
        src_len = float(src_lens[0])
        sec_qty = int(sec_cut.get('qty', 0) or 0)
        if sec_qty <= 0:
            continue

        # Secondary с rest чужой нагрузки уже полностью обработан в
        # _apply_crossload_rest_secondaries, кроме type=='transverse'
        # (тот проход его намеренно пропускает).
        handled_by_crossload = (
            _is_crossload_rest_secondary(sec_cut, rest_groups, load_key, current_plan)
            and sec_cut.get('type') != 'transverse'
        )
        if handled_by_crossload:
            continue

        transverse_remainder_cost, transverse_remainder_terms, trans_cuts = (
            _apply_transverse_remainder_from_cut(
                src_len=src_len,
                product_length=length,
                sec_qty=sec_qty,
                qty=qty,
                base_price_1_2m=base_price_1_2m,
                width_mm=width_mm,
                transverse_remainder_cost=transverse_remainder_cost,
                transverse_remainder_terms=transverse_remainder_terms,
                trans_cuts=trans_cuts,
                count_transverse_op=True,
            )
        )

        # Продольный рез на rest-полосе чужой primary той же нагрузки.
        # Чистый transverse (waste=0, pieces=1) даёт 0 резов автоматически.
        if int(sec_cut.get('source', 0) or 0) > MIN_BILLABLE_TRIM_MM:
            cut_count = _longitudinal_cuts_for_rest_secondary(
                sec_cut,
                min_one_cut_per_op=False,
            )
            if cut_count > 0:
                sec_pieces = int(sec_cut.get('pieces', 1) or 1)
                kept_pieces = sec_pieces + max(0, len(sec_cuts) - 1)
                total_cuts_for_this_size += cut_count
                total_plates_from_cuts += sec_qty * kept_pieces
                long_cut_meterage += cut_count * src_len
                long_cut_length_display = src_len

    return (
        total_cuts_for_this_size,
        total_plates_from_cuts,
        long_cut_meterage,
        long_cut_length_display,
        trans_cuts,
        transverse_remainder_cost,
        transverse_remainder_terms,
    )


def _calc_trim_components(
    current_plan: dict | None,
    *,
    length: float,
    width_mm: int,
    qty: int,
    base_price_1_2m: float,
    base_price: float,
    load_code: int,
    price_table: dict,
    deps: ProcurementDeps | None = None,
):
    """
    Единый расчёт резов/остатков/отходов.

    Primary-продукт: первичные резы + учёт остатка полосы + same-width cascade secondary.
    Secondary-продукт (другая ширина): матч по cuts/lengths без primary width.
    """
    _deps = resolve_procurement_deps(deps)
    rest_cost = 0.0
    rest_width_mm = 0
    rest_used = False
    waste_cost = 0.0
    waste_terms: list[tuple[float, int]] = []
    trans_cuts = 0.0
    transverse_remainder_cost = 0.0
    transverse_remainder_terms: list[tuple[float, int]] = []
    total_cuts_for_this_size = 0
    total_plates_from_cuts = 0
    long_cut_meterage = 0.0
    long_cut_length_display = length
    primary_matched = False
    rest_groups: dict[tuple[int, float], int] = {}
    load_key = int(math.floor(float(load_code)))

    if current_plan and current_plan.get('primary_cuts'):
        matched_primary: list[tuple[int, float, int]] = []
        for prim_cut in current_plan['primary_cuts']:
            if prim_cut.get('width') != width_mm:
                continue
            if not _is_same_length(prim_cut.get('lengths', []), length):
                continue
            if not _cut_matches_load(prim_cut, load_key):
                continue

            primary_matched = True
            prim_qty = int(prim_cut.get('qty', 0) or 0)
            prim_len = _cut_length_from_lengths(prim_cut.get('lengths', []), length)
            total_plates_from_cuts += prim_qty
            long_cut_length_display = prim_len

            primary_rest_width_mm = int(prim_cut.get('rest', 0) or 0)
            matched_primary.append((prim_qty, prim_len, primary_rest_width_mm))

        if matched_primary:
            for prim_qty, prim_len, rest_mm in matched_primary:
                if rest_mm > MIN_BILLABLE_TRIM_MM:
                    key = (rest_mm, prim_len)
                    rest_groups[key] = rest_groups.get(key, 0) + prim_qty

            primary_plate_qty = sum(prim_qty for prim_qty, _, _ in matched_primary)
            cascade_qty_own = _cascade_qty_on_own_rest(
                current_plan, rest_groups, load_key, width_mm, length
            )
            skip_primary_rest_cut = (
                cascade_qty_own > 0
                and primary_plate_qty < cascade_qty_own
            )
            for prim_qty, prim_len, rest_mm in matched_primary:
                if rest_mm <= MIN_BILLABLE_TRIM_MM:
                    continue
                if skip_primary_rest_cut:
                    continue
                total_cuts_for_this_size += prim_qty
                long_cut_meterage += prim_qty * prim_len

            has_secondary_ops = bool(current_plan.get('secondary_cuts'))
            unused_strip_total_mm = 0
            strip_partially_used = False
            all_rests_used = bool(rest_groups)
            assigned_used_from_rest: dict[tuple[int, float], int] = {}
            for prim_cut in current_plan['primary_cuts']:
                if prim_cut.get('width') != width_mm:
                    continue
                if not _is_same_length(prim_cut.get('lengths', []), length):
                    continue
                if not _cut_matches_load(prim_cut, load_key):
                    continue
                rest_mm = int(prim_cut.get('rest', 0) or 0)
                if rest_mm <= MIN_BILLABLE_TRIM_MM:
                    continue
                prim_len = _cut_length_from_lengths(prim_cut.get('lengths', []), length)
                prim_qty = int(prim_cut.get('qty', 0) or 0)
                parent_id = prim_cut.get('primary_instance_id')
                charge_as_strip_waste = False
                strip_unused_mm = 0
                unused_per_strip = 0
                if parent_id:
                    consumed_per_strip = _consumed_width_mm_on_rest_strip(
                        current_plan,
                        rest_mm,
                        prim_len,
                        parent_instance_id=parent_id,
                    )
                    unused_per_strip = max(0, rest_mm - consumed_per_strip)
                else:
                    rest_key = (rest_mm, round(prim_len, 3))
                    total_used = _used_rest_strip_qty_legacy(
                        current_plan, rest_mm, prim_len
                    )
                    already_assigned = assigned_used_from_rest.get(rest_key, 0)
                    used_qty = min(
                        prim_qty, max(0, total_used - already_assigned)
                    )
                    assigned_used_from_rest[rest_key] = already_assigned + used_qty
                    cascade_same_width = False
                    for sec_cut in current_plan.get('secondary_cuts') or []:
                        if int(sec_cut.get('source', 0) or 0) != rest_mm:
                            continue
                        if not _is_same_length(sec_cut.get('source_lengths', []), prim_len):
                            continue
                        if _width_matches_cut(width_mm, sec_cut.get('cuts', []) or []):
                            cascade_same_width = True
                            break
                    if cascade_same_width:
                        width_consumed = _consumed_width_mm_on_rest_strip(
                            current_plan, rest_mm, prim_len, parent_instance_id=None
                        )
                        unused_per_strip = max(0, rest_mm - width_consumed)
                        strip_unused_mm = unused_per_strip
                        charge_as_strip_waste = unused_per_strip > MIN_BILLABLE_TRIM_MM
                        if width_consumed > 0 and charge_as_strip_waste:
                            strip_partially_used = True
                    else:
                        width_consumed = _consumed_width_mm_on_rest_strip(
                            current_plan, rest_mm, prim_len, parent_instance_id=None
                        )
                        if 0 < width_consumed < rest_mm:
                            unused_per_strip = rest_mm - width_consumed
                            strip_unused_mm = unused_per_strip * min(used_qty, prim_qty)
                            charge_as_strip_waste = unused_per_strip > MIN_BILLABLE_TRIM_MM
                            if charge_as_strip_waste:
                                strip_partially_used = True
                            fully_unused_qty = max(0, prim_qty - used_qty)
                            if fully_unused_qty > 0:
                                strip_unused_mm += fully_unused_qty * rest_mm
                        else:
                            unused_strip_qty = max(0, prim_qty - used_qty)
                            strip_unused_mm = unused_strip_qty * rest_mm
                            charge_as_strip_waste = False
                if parent_id:
                    if consumed_per_strip > 0 and unused_per_strip > 0:
                        strip_partially_used = True
                    strip_unused_mm = unused_per_strip * prim_qty
                    charge_as_strip_waste = unused_per_strip > MIN_BILLABLE_TRIM_MM
                if strip_unused_mm <= MIN_BILLABLE_TRIM_MM:
                    continue
                all_rests_used = False
                if charge_as_strip_waste and base_price_1_2m > 0 and qty > 0:
                    unused_strip_total_mm += strip_unused_mm
                    waste_terms.append((unused_per_strip, prim_qty))
                elif base_price_1_2m > 0 and qty > 0:
                    rest_cost += (strip_unused_mm / 1200.0) * base_price_1_2m / qty
                    rest_width_mm = max(rest_width_mm, strip_unused_mm)

            if unused_strip_total_mm > 0 and base_price_1_2m > 0 and qty > 0:
                waste_cost += (unused_strip_total_mm / 1200.0) * base_price_1_2m / qty
                if strip_partially_used:
                    rest_used = True
            elif all_rests_used or strip_partially_used:
                rest_used = True
                rest_width_mm = 0

            (
                total_cuts_for_this_size,
                total_plates_from_cuts,
                long_cut_meterage,
                long_cut_length_display,
                waste_cost,
                waste_terms,
                trans_cuts,
                transverse_remainder_cost,
                transverse_remainder_terms,
            ) = _apply_cascade_secondary_for_primary(
                current_plan=current_plan,
                rest_groups=rest_groups,
                matched_primary=matched_primary,
                length=length,
                width_mm=width_mm,
                qty=qty,
                base_price_1_2m=base_price_1_2m,
                base_price=base_price,
                load_code=load_code,
                price_table=price_table,
                deps=_deps,
                total_cuts_for_this_size=total_cuts_for_this_size,
                total_plates_from_cuts=total_plates_from_cuts,
                long_cut_meterage=long_cut_meterage,
                long_cut_length_display=long_cut_length_display,
                waste_cost=waste_cost,
                waste_terms=waste_terms,
                trans_cuts=trans_cuts,
                transverse_remainder_cost=transverse_remainder_cost,
                transverse_remainder_terms=transverse_remainder_terms,
            )

            (
                total_cuts_for_this_size,
                total_plates_from_cuts,
                long_cut_meterage,
                long_cut_length_display,
                waste_cost,
                waste_terms,
                trans_cuts,
                transverse_remainder_cost,
                transverse_remainder_terms,
            ) = _apply_crossload_rest_secondaries(
                current_plan=current_plan,
                rest_groups=rest_groups,
                length=length,
                width_mm=width_mm,
                qty=qty,
                base_price_1_2m=base_price_1_2m,
                base_price=base_price,
                load_code=load_code,
                price_table=price_table,
                deps=_deps,
                total_cuts_for_this_size=total_cuts_for_this_size,
                total_plates_from_cuts=total_plates_from_cuts,
                long_cut_meterage=long_cut_meterage,
                long_cut_length_display=long_cut_length_display,
                waste_cost=waste_cost,
                waste_terms=waste_terms,
                trans_cuts=trans_cuts,
                transverse_remainder_cost=transverse_remainder_cost,
                transverse_remainder_terms=transverse_remainder_terms,
            )

            (
                total_cuts_for_this_size,
                total_plates_from_cuts,
                long_cut_meterage,
                long_cut_length_display,
                trans_cuts,
                transverse_remainder_cost,
                transverse_remainder_terms,
            ) = _apply_transverse_on_primary_strip(
                current_plan=current_plan,
                rest_groups=rest_groups,
                length=length,
                width_mm=width_mm,
                qty=qty,
                base_price_1_2m=base_price_1_2m,
                load_code=load_code,
                trans_cuts=trans_cuts,
                transverse_remainder_cost=transverse_remainder_cost,
                transverse_remainder_terms=transverse_remainder_terms,
                total_cuts_for_this_size=total_cuts_for_this_size,
                total_plates_from_cuts=total_plates_from_cuts,
                long_cut_meterage=long_cut_meterage,
                long_cut_length_display=long_cut_length_display,
            )

    if not primary_matched and current_plan and current_plan.get('secondary_cuts'):
        plan_rest_groups = _rest_groups_from_plan(current_plan, load_key)
        sec_skip_length = 0
        sec_skip_width = 0
        secondary_matches = 0
        for sec_cut in current_plan['secondary_cuts']:
            if not _cut_matches_load(sec_cut, load_key):
                continue
            if not _is_same_length(sec_cut.get('lengths', []), length):
                sec_skip_length += 1
                continue

            sec_cuts = sec_cut.get('cuts', []) or []
            if not _width_matches_cut(width_mm, sec_cuts):
                sec_skip_width += 1
                continue

            secondary_matches += 1
            src_lens = sec_cut.get('source_lengths', []) or []
            if src_lens:
                long_cut_length_display = float(src_lens[0])

            from_rest = _secondary_matches_primary_rest(sec_cut, plan_rest_groups)
            from_crossload = _is_crossload_rest_secondary(
                sec_cut, plan_rest_groups, load_key, current_plan
            )
            charge_strip_waste = not (from_rest or from_crossload)

            (
                total_cuts_for_this_size,
                total_plates_from_cuts,
                long_cut_meterage,
                waste_cost,
                waste_terms,
                trans_cuts,
                transverse_remainder_cost,
                transverse_remainder_terms,
            ) = _apply_secondary_cut(
                sec_cut,
                length=length,
                width_mm=width_mm,
                qty=qty,
                base_price_1_2m=base_price_1_2m,
                base_price=base_price,
                load_code=load_code,
                price_table=price_table,
                deps=_deps,
                total_cuts_for_this_size=total_cuts_for_this_size,
                total_plates_from_cuts=total_plates_from_cuts,
                long_cut_meterage=long_cut_meterage,
                waste_cost=waste_cost,
                waste_terms=waste_terms,
                trans_cuts=trans_cuts,
                transverse_remainder_cost=transverse_remainder_cost,
                transverse_remainder_terms=transverse_remainder_terms,
                charge_strip_waste=charge_strip_waste,
                # Кросс-нагрузка: минимум 1 продольный рез на операцию; с rest primary
                # той же нагрузки — только если waste > порога (см. _longitudinal_cuts_*).
                min_one_cut_per_op=from_crossload,
            )

        if secondary_matches == 0 and (sec_skip_length > 0 or sec_skip_width > 0):
            print(
                f'[DEBUG] trim_match: {length}x{width_mm}мм '
                f'secondary_cuts не сматчены (skip_length={sec_skip_length}, skip_width={sec_skip_width})'
            )

    long_cut_cost = (
        (long_cut_meterage * LONG_CUT_PRICE_PER_M) / qty if qty > 0 else 0.0
    )

    return {
        'rest_cost': rest_cost,
        'rest_width_mm': rest_width_mm,
        'rest_used': rest_used,
        'waste_cost': waste_cost,
        'waste_terms': waste_terms,
        'trans_cuts': trans_cuts,
        'transverse_remainder_cost': transverse_remainder_cost,
        'transverse_remainder_terms': transverse_remainder_terms,
        'total_cuts_for_this_size': total_cuts_for_this_size,
        'total_plates_from_cuts': total_plates_from_cuts,
        'long_cut_meterage': long_cut_meterage,
        'long_cut_cost': long_cut_cost,
        'long_cut_length_display': long_cut_length_display,
        'primary_matched': primary_matched,
    }


def resolve_long_cut_pricing(
    trim: dict,
    *,
    qty: int,
    length: float,
    width_m: float,
    current_plan: dict | None,
    fallback_long_cuts: int = 0,
    plate_name: str = '',
) -> tuple[float, float, int]:
    """
    Определяет long_cut_cost и long_cuts для позиции после trim.

    Returns: (long_cut_cost, long_cuts, total_cuts_count)
    """
    has_plan = bool(current_plan and current_plan.get('primary_cuts'))
    # Только фактический метраж реза; total_plates_from_cuts>0 при rest=0 (1080 factory strip)
    # не означает отсутствие продольного реза — иначе ранний return обнуляет стоимость.
    has_trim_cuts = trim.get('long_cut_meterage', 0) > 0
    width_mm = int(round(width_m * 1000))

    if has_trim_cuts and qty > 0:
        return (
            trim['long_cut_cost'],
            trim['total_cuts_for_this_size'],
            trim['total_cuts_for_this_size'],
        )

    if abs(width_m - 1.2) < 0.01:
        return 0.0, 0, 0

    # Плиты 10,2–10,8 м: factory strip + 1 продольный рез на плиту (R5).
    if 1020 <= width_mm <= 1080:
        cost = LONG_CUT_PRICE_PER_M * length
        return cost, 1, 0

    if not has_plan:
        long_cuts = fallback_long_cuts if fallback_long_cuts else (1 if width_m < 1.15 else 0)
        cost = long_cuts * (LONG_CUT_PRICE_PER_M * length)
        return cost, long_cuts, 0

    if width_m < 1.15:
        # Secondary с полосы rest primary уже сматчен trim'ом — нулевой метраж
        # намеренный (один физический рез на primary), не подменять fallback'ом.
        trim_matched_secondary = (
            not trim.get('primary_matched', False)
            and int(trim.get('total_plates_from_cuts', 0) or 0) > 0
        )
        if trim_matched_secondary:
            return 0.0, 0, 0

        long_cuts = fallback_long_cuts if fallback_long_cuts else 1
        cost = long_cuts * (LONG_CUT_PRICE_PER_M * length)
        if plate_name:
            print(
                f'[WARNING] План есть, но trim не нашёл резов для {plate_name} '
                f'({length}м × {width_mm}мм), fallback long_cuts={long_cuts}'
            )
        return cost, long_cuts, 0

    if plate_name:
        print(
            f'[WARNING] План есть, но trim не нашёл резов для {plate_name} '
            f'({length}м × {int(round(width_m * 1000))}мм)'
        )
    return 0.0, 0, 0


def apply_factory_strip_waste(
    *,
    width_mm: int,
    base_price_1_2m: float,
    rest_cost: float,
    rest_used: bool,
    waste_cost: float,
    waste_terms: list,
    qty: int,
) -> tuple[float, list]:
    """
    Таблица завода: для 1020–1080 мм обрезок идёт в утилизацию.

    Fallback, если стоимость ещё не учтена в rest_cost.
    Не дублирует остаток, уже начисленный trim или ручным fallback.
    """
    if rest_cost > 0 or rest_used:
        return waste_cost, waste_terms
    if not (1020 <= width_mm <= 1080 and base_price_1_2m > 0):
        return waste_cost, waste_terms
    extra_waste_mm = 1200 - width_mm
    if extra_waste_mm <= 0:
        return waste_cost, waste_terms
    waste_terms.append((extra_waste_mm, qty))
    waste_cost += (extra_waste_mm / 1200.0) * base_price_1_2m
    return waste_cost, waste_terms


def format_long_cut_calculation(trim: dict, qty: int) -> str | None:
    """Формула продольного реза для breakdown (с учётом source_length)."""
    meterage = float(trim.get('long_cut_meterage', 0) or 0)
    total = float(trim.get('total_cuts_for_this_size', 0) or 0)
    if meterage <= 0 or total <= 0:
        return None
    avg_len = meterage / total
    price = LONG_CUT_PRICE_PER_M
    if qty > 1:
        return f"{price:.0f} × {avg_len:.1f} × {total:.0f} / {qty}".replace('.', ',')
    return f"{price:.0f} × {avg_len:.1f} × {total:.0f}".replace('.', ',')


def format_transverse_remainder_calculation(
    trim: dict,
    qty: int,
    *,
    base_price_1_2m: float,
    width_m: float,
    length_m: float,
) -> str | None:
    """Формула остатка после поперечного реза для breakdown."""
    terms = trim.get('transverse_remainder_terms') or []
    if not terms or base_price_1_2m <= 0 or length_m <= 0:
        return None

    rem_m, _ = terms[0]
    base_str = f"{base_price_1_2m:,.2f}".replace(',', ' ').replace('.', ',')
    width_str = f"{width_m:.2f}".replace('.', ',')
    rem_str = f"{rem_m:.2f}".replace('.', ',')
    length_str = f"{length_m:.2f}".replace('.', ',')
    expr = f"{base_str} × ({width_str} / 1,2) × ({rem_str} / {length_str})"
    if qty > 1:
        return f"{expr} / {qty}"
    return expr
