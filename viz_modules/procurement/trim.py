from __future__ import annotations

import core.config_and_data as cfg
from ..price_utils import find_price_for_plate
from .plan_lookup import _is_same_length
from .ports import ProcurementDeps, resolve_procurement_deps


def _cut_length_from_lengths(lengths: list | None, default_length: float) -> float:
    if lengths:
        try:
            return float(lengths[0])
        except (TypeError, ValueError, IndexError):
            pass
    return default_length


def _width_matches_cut(width_mm: int, sec_cuts: list) -> bool:
    return any(abs(width_mm - int(cut_width)) <= 20 for cut_width in sec_cuts)


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
) -> tuple[int, int, float, float, list, float, float, list[tuple[float, int]]]:
    del base_price, load_code, price_table, deps  # kept for call-site compatibility

    sec_qty = int(sec_cut.get('qty', 0) or 0)
    sec_pieces = int(sec_cut.get('pieces', 1) or 1)
    current_cuts = sec_qty * sec_pieces
    total_cuts_for_this_size += current_cuts
    total_plates_from_cuts += current_cuts

    src_lens = sec_cut.get('source_lengths', []) or []
    src_len = float(src_lens[0]) if src_lens else length
    long_cut_meterage += current_cuts * src_len

    waste_w_mm = float(sec_cut.get('waste', 0) or 0)
    if waste_w_mm > 0 and base_price_1_2m > 0 and qty > 0:
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

    for sec_cut in current_plan.get('secondary_cuts') or []:
        if remaining_slots <= 0:
            break
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


def _apply_transverse_on_primary_strip(
    *,
    current_plan: dict,
    rest_groups: dict[tuple[int, float], int],
    length: float,
    width_mm: int,
    qty: int,
    base_price_1_2m: float,
    trans_cuts: float,
    transverse_remainder_cost: float,
    transverse_remainder_terms: list[tuple[float, int]],
) -> tuple[float, float, list[tuple[float, int]]]:
    """Поперечный рез на основной полосе primary (не из rest по ширине)."""
    for sec_cut in current_plan.get('secondary_cuts') or []:
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

    return trans_cuts, transverse_remainder_cost, transverse_remainder_terms


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

    if current_plan and current_plan.get('primary_cuts'):
        matched_primary: list[tuple[int, float, int]] = []
        for prim_cut in current_plan['primary_cuts']:
            if prim_cut.get('width') != width_mm:
                continue
            if not _is_same_length(prim_cut.get('lengths', []), length):
                continue

            primary_matched = True
            prim_qty = int(prim_cut.get('qty', 0) or 0)
            prim_len = _cut_length_from_lengths(prim_cut.get('lengths', []), length)
            total_cuts_for_this_size += prim_qty
            total_plates_from_cuts += prim_qty
            long_cut_meterage += prim_qty * prim_len
            long_cut_length_display = prim_len

            primary_rest_width_mm = int(prim_cut.get('rest', 0) or 0)
            matched_primary.append((prim_qty, prim_len, primary_rest_width_mm))

        if matched_primary:
            for prim_qty, prim_len, rest_mm in matched_primary:
                if rest_mm > 0:
                    key = (rest_mm, prim_len)
                    rest_groups[key] = rest_groups.get(key, 0) + prim_qty

            unused_rest_total_mm = 0
            all_rests_used = bool(rest_groups)
            for (rest_mm, prim_len), produced in rest_groups.items():
                used_rests = 0
                for sec_cut in (current_plan.get('secondary_cuts') or []):
                    if int(sec_cut.get('source', 0) or 0) != rest_mm:
                        continue
                    if not _is_same_length(sec_cut.get('source_lengths', []), prim_len):
                        continue
                    used_rests += int(sec_cut.get('qty', 0) or 0)

                unused_rests = max(0, produced - used_rests)
                if unused_rests > 0:
                    all_rests_used = False
                    unused_rest_total_mm += unused_rests * rest_mm

            if unused_rest_total_mm > 0 and base_price_1_2m > 0 and qty > 0:
                rest_cost = (unused_rest_total_mm / 1200.0) * base_price_1_2m / qty
                rest_width_mm = unused_rest_total_mm
            elif all_rests_used:
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

            trans_cuts, transverse_remainder_cost, transverse_remainder_terms = (
                _apply_transverse_on_primary_strip(
                    current_plan=current_plan,
                    rest_groups=rest_groups,
                    length=length,
                    width_mm=width_mm,
                    qty=qty,
                    base_price_1_2m=base_price_1_2m,
                    trans_cuts=trans_cuts,
                    transverse_remainder_cost=transverse_remainder_cost,
                    transverse_remainder_terms=transverse_remainder_terms,
                )
            )

    if not primary_matched and current_plan and current_plan.get('secondary_cuts'):
        sec_skip_length = 0
        sec_skip_width = 0
        secondary_matches = 0
        for sec_cut in current_plan['secondary_cuts']:
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
            )

        if secondary_matches == 0 and (sec_skip_length > 0 or sec_skip_width > 0):
            print(
                f'[DEBUG] trim_match: {length}x{width_mm}мм '
                f'secondary_cuts не сматчены (skip_length={sec_skip_length}, skip_width={sec_skip_width})'
            )

    long_cut_cost = (
        (long_cut_meterage * cfg.LONG_CUT_PRICE_PER_M) / qty if qty > 0 else 0.0
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
    has_trim_cuts = trim.get('long_cut_meterage', 0) > 0 or trim.get('total_plates_from_cuts', 0) > 0

    if has_trim_cuts and qty > 0:
        return (
            trim['long_cut_cost'],
            trim['total_cuts_for_this_size'],
            trim['total_cuts_for_this_size'],
        )

    if not has_plan:
        if abs(width_m - 1.2) < 0.01:
            return 0.0, 0, 0
        long_cuts = fallback_long_cuts if fallback_long_cuts else (1 if width_m < 1.15 else 0)
        cost = long_cuts * (cfg.LONG_CUT_PRICE_PER_M * length)
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
    price = cfg.LONG_CUT_PRICE_PER_M
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
