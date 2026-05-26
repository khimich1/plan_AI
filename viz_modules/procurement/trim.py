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
) -> tuple[int, int, float, float, list, float]:
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
        if sec_cut.get('type') == 'transverse' or abs(src_len - length) > 0.05:
            if qty > 0:
                trans_cuts += (1.0 * sec_qty) / qty

            len_waste = src_len - length
            if len_waste > 0.01:
                src_price_full = deps.get_price(src_len, load_code, deps.db_path)
                if src_price_full is None:
                    src_price_full = find_price_for_plate(price_table, src_len, load_code) or 0.0
                src_price_width = src_price_full * (width_mm / 1200.0)
                cost_len_waste = src_price_width - base_price
                if cost_len_waste > 0 and qty > 0:
                    waste_cost += cost_len_waste * (sec_qty / qty)

    return (
        total_cuts_for_this_size,
        total_plates_from_cuts,
        long_cut_meterage,
        waste_cost,
        waste_terms,
        trans_cuts,
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

    Primary-продукт: только первичные резы + учёт остатка полосы.
    Secondary-продукт: матч по cuts/lengths без привязки к primary width.
    """
    _deps = resolve_procurement_deps(deps)
    rest_cost = 0.0
    rest_width_mm = 0
    rest_used = False
    waste_cost = 0.0
    waste_terms: list[tuple[float, int]] = []
    trans_cuts = 0.0
    total_cuts_for_this_size = 0
    total_plates_from_cuts = 0
    long_cut_meterage = 0.0
    long_cut_length_display = length
    primary_matched = False

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
            rest_groups: dict[tuple[int, float], int] = {}
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
