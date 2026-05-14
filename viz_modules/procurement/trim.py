from __future__ import annotations

from ..price_utils import find_price_for_plate
from .plan_lookup import _is_same_length
from .ports import ProcurementDeps, resolve_procurement_deps


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
    Правило: непереиспользованный первичный обрезок идёт в `rest_cost`,
    вторичные отходы (`waste`/`length_waste`) — в `waste_cost`.
    """
    _deps = resolve_procurement_deps(deps)
    rest_cost = 0.0
    rest_width_mm = 1200 - width_mm if width_mm < 1150 else 0
    rest_used = False
    waste_cost = 0.0
    waste_terms = []
    trans_cuts = 0.0
    total_cuts_for_this_size = 0
    total_plates_from_cuts = 0

    if current_plan and current_plan.get('primary_cuts'):
        primary_rest_width_mm = 0

        for prim_cut in current_plan['primary_cuts']:
            if prim_cut.get('width') != width_mm:
                continue
            if not _is_same_length(prim_cut.get('lengths', []), length):
                continue

            prim_qty = int(prim_cut.get('qty', 0) or 0)
            total_cuts_for_this_size += prim_qty
            total_plates_from_cuts += prim_qty
            primary_rest_width_mm = int(prim_cut.get('rest', 0) or 0)

            if primary_rest_width_mm > 0:
                produced_rests = prim_qty
                used_rests = 0
                for sec_cut in (current_plan.get('secondary_cuts') or []):
                    if int(sec_cut.get('source', 0) or 0) != primary_rest_width_mm:
                        continue
                    if not _is_same_length(sec_cut.get('source_lengths', []), length):
                        continue
                    used_rests += int(sec_cut.get('qty', 0) or 0)

                unused_rests = max(0, produced_rests - used_rests)
                unused_rest_total_mm = unused_rests * primary_rest_width_mm
                if unused_rest_total_mm > 0 and base_price_1_2m > 0 and qty > 0:
                    rest_cost = (unused_rest_total_mm / 1200.0) * base_price_1_2m / qty
                    rest_width_mm = unused_rest_total_mm
                elif produced_rests > 0:
                    rest_used = True
                    rest_width_mm = 0

            # secondary_cuts: основной матч по целевой ширине + fallback по source/rest
            secondary_matches = 0
            sec_skip_length = 0
            sec_skip_width = 0
            for sec_cut in (current_plan.get('secondary_cuts') or []):
                if not _is_same_length(sec_cut.get('lengths', []), length):
                    sec_skip_length += 1
                    continue

                sec_cuts = sec_cut.get('cuts', []) or []
                width_match = any(abs(width_mm - int(cut_width)) <= 20 for cut_width in sec_cuts)
                source_match = (
                    primary_rest_width_mm > 0 and
                    int(sec_cut.get('source', 0) or 0) == primary_rest_width_mm and
                    _is_same_length(sec_cut.get('source_lengths', []), length)
                )
                if not (width_match or source_match):
                    sec_skip_width += 1
                    continue

                secondary_matches += 1
                sec_qty = int(sec_cut.get('qty', 0) or 0)
                sec_pieces = int(sec_cut.get('pieces', 1) or 1)
                current_cuts = sec_qty * sec_pieces
                total_cuts_for_this_size += current_cuts
                total_plates_from_cuts += current_cuts

                waste_w_mm = float(sec_cut.get('waste', 0) or 0)
                if waste_w_mm > 0 and base_price_1_2m > 0 and qty > 0:
                    cost_of_waste_piece = (waste_w_mm / 1200.0) * base_price_1_2m
                    waste_cost += (cost_of_waste_piece * sec_qty) / qty
                    waste_terms.append((waste_w_mm, sec_qty))

                src_lens = sec_cut.get('source_lengths', [])
                if src_lens:
                    src_len = float(src_lens[0])
                    if sec_cut.get('type') == 'transverse' or abs(src_len - length) > 0.05:
                        if qty > 0:
                            trans_cuts += (1.0 * sec_qty) / qty

                        len_waste = src_len - length
                        if len_waste > 0.01:
                            src_price_full = _deps.get_price(src_len, load_code, _deps.db_path)
                            if src_price_full is None:
                                src_price_full = find_price_for_plate(price_table, src_len, load_code) or 0.0
                            src_price_width = src_price_full * (width_mm / 1200.0)
                            cost_len_waste = src_price_width - base_price
                            if cost_len_waste > 0 and qty > 0:
                                waste_cost += cost_len_waste * (sec_qty / qty)

            if secondary_matches == 0 and (sec_skip_length > 0 or sec_skip_width > 0):
                print(
                    f'[DEBUG] trim_match: {length}x{width_mm}мм '
                    f'secondary_cuts не сматчены (skip_length={sec_skip_length}, skip_width={sec_skip_width})'
                )
            break

    return {
        'rest_cost': rest_cost,
        'rest_width_mm': rest_width_mm,
        'rest_used': rest_used,
        'waste_cost': waste_cost,
        'waste_terms': waste_terms,
        'trans_cuts': trans_cuts,
        'total_cuts_for_this_size': total_cuts_for_this_size,
        'total_plates_from_cuts': total_plates_from_cuts,
    }
