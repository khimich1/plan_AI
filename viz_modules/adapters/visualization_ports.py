"""Default viz_modules implementations for ``core.ports.visualization``."""

from __future__ import annotations


def register_default_visualization_ports() -> None:
    from core.ports.visualization import (
        register_build_component_breakdown,
        register_build_component_breakdown_production,
        register_build_layout_sequence,
        register_build_price_rows,
        register_build_price_rows_production,
        register_build_procurement_items,
        register_draw_segment,
        register_draw_split_plate,
        register_draw_transverse_cut,
        register_get_orders_from_opt_plan,
        register_load_price_table_from_xlsx,
        register_visualize_plan,
    )
    from core.visualization import visualize_plan
    from viz_modules.layout_sequence import build_layout_sequence
    from viz_modules.price_utils import load_price_table_from_xlsx
    from viz_modules.procurement import (
        build_component_breakdown,
        build_component_breakdown_production,
        build_price_rows,
        build_price_rows_production,
        build_procurement_items,
        get_orders_from_opt_plan,
    )
    from viz_modules.visualization_drawing import (
        _draw_segment,
        _draw_split_plate,
        _draw_transverse_cut,
    )

    register_build_layout_sequence(build_layout_sequence)
    register_load_price_table_from_xlsx(load_price_table_from_xlsx)
    register_build_price_rows(build_price_rows)
    register_build_component_breakdown(build_component_breakdown)
    register_build_procurement_items(build_procurement_items)
    register_get_orders_from_opt_plan(get_orders_from_opt_plan)
    register_build_price_rows_production(build_price_rows_production)
    register_build_component_breakdown_production(build_component_breakdown_production)
    register_draw_segment(_draw_segment)
    register_draw_split_plate(_draw_split_plate)
    register_draw_transverse_cut(_draw_transverse_cut)
    register_visualize_plan(visualize_plan)
