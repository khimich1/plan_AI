from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import Any, TypeAlias

from core import config_and_data as cfg


DemandKey: TypeAlias = tuple[float, int, float | int]
CutOption: TypeAlias = dict[str, Any]
OrderInfoGetter: TypeAlias = Callable[[dict, tuple], dict]


# Kerf is not applied directly. The legacy model encodes it through narrowing
# waste values below; keep the constant for backward compatibility.
KERF_WIDTH_MM = 0

# С одной исходной плиты по ширине дорожки (базовая ширина, напр. 1200 мм) нельзя
# получить больше этого числа готовых полос одной искомой ширины: первичный «main»
# плюс вторичные из остатка (см. direct 300+900 + 3×300 из 900 → 4, недопустимо).
MAX_PRODUCT_SLABS_PER_BASE_WIDTH = 3

NARROWING_TABLE: list[tuple[int, int, int]] = [
    (480, 460, 20),
    (500, 460, 40),
    (495, 460, 35),
    (740, 720, 20),
    (690, 660, 30),
    (890, 860, 30),
    (495, 480, 15),
]


@dataclass(frozen=True)
class GeometryConfig:
    plate_width: int = 1200
    min_useful_width: int = 200
    tolerance_width: int = 20


@dataclass(frozen=True)
class PrimaryCutOptionsResult:
    raw_options: list[CutOption]
    options: list[CutOption]
    next_option_id: int
    solid_widths: list[int]
    target_to_sources: dict[int, list[tuple[int, int, int]]]


def _canonical_length(length: float | int | str | None) -> float:
    """
    Canonical length in meters rounded to 2 decimals.
    """
    try:
        return round(float(length), 2)
    except (TypeError, ValueError):
        return 0.0


def _primary_main_equals_target_for_rest(
    primary_options: list[CutOption],
    source_length: float,
    source_rest_mm: int,
    target_width_mm: int,
) -> bool:
    """
    True, если среди первичных опций есть рез с тем же остатком source_rest_mm и длиной,
    дающий уже одну полосу шириной target_width_mm (поле main).
    """
    sl = _canonical_length(source_length)
    for o in primary_options:
        if _canonical_length(o.get("length")) != sl:
            continue
        if int(o.get("rest") or 0) != int(source_rest_mm):
            continue
        if int(o.get("main") or 0) == int(target_width_mm):
            return True
    return False


def _max_secondary_equal_width_pieces(
    primary_options: list[CutOption],
    source_length: float,
    source_rest_mm: int,
    target_width_mm: int,
) -> int:
    """
    Верхняя граница числа полос ширины target_width_mm, нарезаемых из остатка
    source_rest_mm, с учётом лимита MAX_PRODUCT_SLABS_PER_BASE_WIDTH на одну базовую плиту.
    """
    cap = MAX_PRODUCT_SLABS_PER_BASE_WIDTH
    if _primary_main_equals_target_for_rest(
        primary_options, source_length, source_rest_mm, target_width_mm
    ):
        cap = max(1, cap - 1)
    return cap


def build_narrowing_source_index(
    plate_width: int = 1200,
    min_main_width: int = 200,
    max_main_width: int = 1000,
) -> dict[int, list[tuple[int, int, int]]]:
    """
    Build target width -> [(main_width, source_rest, waste), ...].
    """
    target_to_sources: dict[int, list[tuple[int, int, int]]] = {}
    for source_rest, target_w, waste in NARROWING_TABLE:
        target_to_sources.setdefault(target_w, [])
        main_w = plate_width - source_rest
        if min_main_width <= main_w <= max_main_width:
            target_to_sources[target_w].append((main_w, source_rest, waste))
    return target_to_sources


def generate_primary_cut_options_2d(
    demand_2d: Mapping[DemandKey, int],
    order_info_list: dict,
    order_info_getter: OrderInfoGetter,
    config: GeometryConfig | None = None,
    start_option_id: int = 0,
) -> PrimaryCutOptionsResult:
    """
    Generate and filter primary 2D cut options for the ILP model.
    """
    geometry_config = config or GeometryConfig()
    plate_width = geometry_config.plate_width
    min_useful_width = geometry_config.min_useful_width

    target_to_sources = build_narrowing_source_index(plate_width=plate_width)
    solid_widths = sorted(set([plate_width, 1080]))

    primary_options: list[CutOption] = []
    option_id = start_option_id

    for (length, width, load_code), _qty in demand_2d.items():
        order_info = order_info_getter(order_info_list, (length, width, load_code))
        option_load_code = cfg.normalize_load_code(
            order_info.get("load_code", load_code) if order_info else load_code,
            default=8,
        )

        if width in solid_widths:
            primary_options.append(
                {
                    "id": option_id,
                    "length": length,
                    "main": width,
                    "rest": 0,
                    "type": "solid",
                    "load_code": option_load_code,
                    "kp_id": order_info.get("kp_id"),
                    "customer": order_info.get("customer"),
                    "kp_date": order_info.get("kp_date"),
                    "plate_name": order_info.get("plate_name"),
                }
            )
            option_id += 1
        elif width < plate_width:
            rest = plate_width - width
            primary_options.append(
                {
                    "id": option_id,
                    "length": length,
                    "main": width,
                    "rest": rest,
                    "type": "direct",
                    "load_code": option_load_code,
                    "kp_id": order_info.get("kp_id"),
                    "customer": order_info.get("customer"),
                    "kp_date": order_info.get("kp_date"),
                    "plate_name": order_info.get("plate_name"),
                }
            )
            option_id += 1

            if width in target_to_sources:
                for main_w, rest_w, waste in target_to_sources[width]:
                    if main_w != width and rest_w >= min_useful_width:
                        primary_options.append(
                            {
                                "id": option_id,
                                "length": length,
                                "main": main_w,
                                "rest": rest_w,
                                "type": "indirect",
                                "target_width": width,
                                "narrowing_waste": waste,
                                "load_code": option_load_code,
                                "kp_id": order_info.get("kp_id"),
                                "customer": order_info.get("customer"),
                                "kp_date": order_info.get("kp_date"),
                                "plate_name": order_info.get("plate_name"),
                            }
                        )
                        option_id += 1

    filtered_primary: list[CutOption] = []
    for opt in primary_options:
        if opt.get("type") == "indirect":
            target_w = opt.get("target_width")
            has_direct = any(
                o["type"] == "direct"
                and o["main"] == target_w
                and _canonical_length(o["length"]) == _canonical_length(opt["length"])
                for o in primary_options
            )
            if has_direct:
                continue
        filtered_primary.append(opt)

    return PrimaryCutOptionsResult(
        raw_options=primary_options,
        options=filtered_primary,
        next_option_id=option_id,
        solid_widths=solid_widths,
        target_to_sources=target_to_sources,
    )


def generate_secondary_cut_options_2d(
    primary_options: list[CutOption],
    demand_2d: Mapping[DemandKey, int],
    config: GeometryConfig | None = None,
) -> list[CutOption]:
    """
    Generate and filter secondary 2D cut options from primary residuals.
    """
    return filter_secondary_cut_options_2d(
        generate_raw_secondary_cut_options_2d(primary_options, demand_2d, config)
    )


def generate_raw_secondary_cut_options_2d(
    primary_options: list[CutOption],
    demand_2d: Mapping[DemandKey, int],
    config: GeometryConfig | None = None,
) -> list[CutOption]:
    """
    Generate unfiltered secondary 2D cut options from primary residuals.
    """
    geometry_config = config or GeometryConfig()
    min_useful_width = geometry_config.min_useful_width
    tolerance_width = geometry_config.tolerance_width

    secondary_options: list[CutOption] = []
    possible_rests: dict[tuple[float, int], list[int]] = {}
    for opt in primary_options:
        key = (opt["length"], opt["rest"])
        possible_rests.setdefault(key, [])
        possible_rests[key].append(opt["id"])

    sec_id = 0
    for (source_length, source_width), source_ids in possible_rests.items():
        if source_width < min_useful_width:
            continue

        for (target_length, target_width, target_load_code), _qty in demand_2d.items():
            if _canonical_length(target_length) == _canonical_length(source_length) and target_width <= source_width:
                max_pieces = source_width // target_width
                _cap = _max_secondary_equal_width_pieces(
                    primary_options, source_length, source_width, target_width
                )
                max_pieces = min(max_pieces, _cap)
                for pieces in range(1, max_pieces + 1):
                    waste = source_width - (pieces * target_width)
                    max_waste_fraction = 0.8 if pieces == 1 else 0.5
                    if waste <= source_width * max_waste_fraction:
                        secondary_options.append(
                            {
                                "id": sec_id,
                                "source_length": source_length,
                                "source_rest": source_width,
                                "output_length": target_length,
                                "output_width": target_width,
                                "pieces": pieces,
                                "waste": waste,
                                "type": "multiple",
                                "source_ids": source_ids,
                                "target_order_key": (target_length, target_width, target_load_code),
                            }
                        )
                        sec_id += 1

            if target_length < source_length - 0.1 and target_width <= source_width:
                _cap = _max_secondary_equal_width_pieces(
                    primary_options, source_length, source_width, target_width
                )
                pieces = min(source_width // target_width, _cap)
                if pieces >= 1:
                    waste_width = source_width - (pieces * target_width)
                    waste_length = (source_length - target_length) * 1000
                    if waste_width < source_width * 0.5:
                        secondary_options.append(
                            {
                                "id": sec_id,
                                "source_length": source_length,
                                "source_rest": source_width,
                                "output_length": target_length,
                                "output_width": target_width,
                                "pieces": pieces,
                                "waste": waste_width,
                                "length_waste": waste_length,
                                "type": "multiple_transverse",
                                "source_ids": source_ids,
                                "target_order_key": (target_length, target_width, target_load_code),
                            }
                        )
                        sec_id += 1

            if (
                _canonical_length(target_length) == _canonical_length(source_length)
                and target_width < source_width <= target_width + 100
            ):
                waste = source_width - target_width
                if waste <= 100:
                    secondary_options.append(
                        {
                            "id": sec_id,
                            "source_length": source_length,
                            "source_rest": source_width,
                            "output_length": target_length,
                            "output_width": target_width,
                            "pieces": 1,
                            "waste": waste,
                            "type": "narrowing",
                            "source_ids": source_ids,
                            "target_order_key": (target_length, target_width, target_load_code),
                        }
                    )
                    sec_id += 1

            if (
                target_length < source_length - 0.1
                and target_width <= source_width
                and source_width - target_width <= tolerance_width
            ):
                length_waste = (source_length - target_length) * 1000
                secondary_options.append(
                    {
                        "id": sec_id,
                        "source_length": source_length,
                        "source_rest": source_width,
                        "output_length": target_length,
                        "output_width": target_width,
                        "pieces": 1,
                        "waste": 0,
                        "length_waste": length_waste,
                        "type": "transverse",
                        "source_ids": source_ids,
                        "target_order_key": (target_length, target_width, target_load_code),
                    }
                )
                sec_id += 1

    return secondary_options


def filter_secondary_cut_options_2d(secondary_options: list[CutOption]) -> list[CutOption]:
    """
    Remove duplicate and high-waste secondary 2D options.
    """
    filtered_secondary: list[CutOption] = []
    seen_combinations: set[tuple[Any, ...]] = set()

    for opt in secondary_options:
        key = (
            opt["source_length"],
            opt["source_rest"],
            opt["output_length"],
            opt["output_width"],
            opt["type"],
            opt.get("pieces", 1),
            opt.get("target_order_key"),
        )
        if key in seen_combinations:
            continue
        seen_combinations.add(key)

        waste_width = opt.get("waste", 0)
        waste_length = opt.get("length_waste", 0)
        source_area = opt["source_length"] * opt["source_rest"]
        waste_area = (waste_width * opt["source_length"]) + (
            waste_length * opt["source_rest"] / 1000.0
        )

        max_waste_fraction_area = 0.8 if opt.get("pieces", 1) == 1 else 0.3
        if opt["type"] != "multiple_transverse" and waste_area > source_area * max_waste_fraction_area:
            continue

        if opt["type"] == "transverse":
            waste_fraction = waste_length / (opt["source_length"] * 1000) if opt["source_length"] > 0 else 0
            if waste_fraction > 0.5:
                continue

        filtered_secondary.append(opt)

    return filtered_secondary


def generate_primary_cut_options_1d(
    target_widths: list[int],
    plate_width: int = 1200,
    min_useful_width: int = 200,
) -> list[CutOption]:
    """
    Generate legacy 1D primary width cut options.
    """
    primary_cut_options: list[CutOption] = []
    solid_widths = sorted(set([plate_width, 1080]))
    for target_w in target_widths:
        if target_w in solid_widths:
            primary_cut_options.append(
                {
                    "id": f"prim_{target_w}",
                    "main": target_w,
                    "rest": 0,
                }
            )
            continue

        rest_w = plate_width - target_w
        if rest_w >= min_useful_width:
            primary_cut_options.append(
                {
                    "id": f"prim_{target_w}",
                    "main": target_w,
                    "rest": rest_w,
                }
            )
    return primary_cut_options


def generate_secondary_cut_options_1d(
    primary_cut_options: list[CutOption],
    target_widths: list[int],
    tolerance: int = 20,
) -> list[CutOption]:
    """
    Generate legacy 1D secondary width cut options.
    """
    secondary_cut_options: list[CutOption] = []
    possible_rests = set(opt["rest"] for opt in primary_cut_options if opt["rest"] > 0)

    for rest_w in possible_rests:
        for target_w1 in target_widths:
            target_w2 = rest_w - target_w1
            for target_w2_candidate in target_widths:
                if abs(target_w2 - target_w2_candidate) <= tolerance:
                    secondary_cut_options.append(
                        {
                            "id": f"sec_{rest_w}_to_{target_w1}_{target_w2_candidate}",
                            "source_rest": rest_w,
                            "output1": target_w1,
                            "output2": target_w2_candidate,
                            "waste": abs(rest_w - target_w1 - target_w2_candidate),
                        }
                    )
                    break

            for target_w_candidate in target_widths:
                max_geom = rest_w // target_w_candidate
                _cap = MAX_PRODUCT_SLABS_PER_BASE_WIDTH
                if any(
                    int(o.get("main") or 0) == int(target_w_candidate)
                    and int(o.get("rest") or 0) == int(rest_w)
                    for o in primary_cut_options
                ):
                    _cap = max(1, MAX_PRODUCT_SLABS_PER_BASE_WIDTH - 1)
                num_pieces = min(max_geom, _cap)
                if num_pieces >= 2:
                    waste = rest_w - (target_w_candidate * num_pieces)
                    if waste < rest_w * 0.5:
                        secondary_cut_options.append(
                            {
                                "id": f"sec_{rest_w}_to_{num_pieces}x{target_w_candidate}",
                                "source_rest": rest_w,
                                "output1": target_w_candidate,
                                "output2": 0,
                                "pieces": num_pieces,
                                "waste": waste,
                            }
                        )

            for target_w_candidate in target_widths:
                if target_w_candidate < rest_w <= target_w_candidate + 100:
                    waste = rest_w - target_w_candidate
                    if waste <= 100:
                        secondary_cut_options.append(
                            {
                                "id": f"sec_{rest_w}_narrow_to_{target_w_candidate}",
                                "source_rest": rest_w,
                                "output1": target_w_candidate,
                                "output2": 0,
                                "pieces": 1,
                                "waste": waste,
                            }
                        )

    return secondary_cut_options
