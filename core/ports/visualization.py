"""Visualization ports — core defines the contract; viz_modules/app wire adapters.

Grep gate: ``core/`` must not ``import viz_modules`` on hot paths.
"""

from __future__ import annotations

from typing import Any, Protocol

_build_layout_sequence: BuildLayoutSequenceFn | None = None
_load_price_table: LoadPriceTableFn | None = None
_build_price_rows: BuildPriceRowsFn | None = None
_build_component_breakdown: BuildComponentBreakdownFn | None = None
_build_procurement_items: BuildProcurementItemsFn | None = None
_get_orders_from_opt_plan: GetOrdersFromOptPlanFn | None = None
_build_price_rows_production: BuildPriceRowsProductionFn | None = None
_build_component_breakdown_production: BuildComponentBreakdownProductionFn | None = None
_draw_segment: DrawSegmentFn | None = None
_draw_split_plate: DrawSplitPlateFn | None = None
_draw_transverse_cut: DrawTransverseCutFn | None = None
_visualize_plan: VisualizePlanFn | None = None


class BuildLayoutSequenceFn(Protocol):
    def __call__(
        self,
        *,
        runtime: Any = None,
        pb_db_path: Any = None,
        log: Any = None,
        traces: Any = None,
        **kwargs: Any,
    ) -> Any: ...


class LoadPriceTableFn(Protocol):
    def __call__(self, xlsx_path: str) -> dict[str, Any]: ...


class BuildPriceRowsFn(Protocol):
    def __call__(self, price_table: dict[str, Any], **kwargs: Any) -> tuple[list, float]: ...


class BuildComponentBreakdownFn(Protocol):
    def __call__(
        self,
        price_table: dict[str, Any],
        price_rows: list | None = None,
        **kwargs: Any,
    ) -> list: ...


class BuildProcurementItemsFn(Protocol):
    def __call__(self, **kwargs: Any) -> list: ...


class GetOrdersFromOptPlanFn(Protocol):
    def __call__(self, opt_snapshot: Any = None) -> list | None: ...


class BuildPriceRowsProductionFn(Protocol):
    def __call__(self, price_table: dict[str, Any], **kwargs: Any) -> tuple[list, float]: ...


class BuildComponentBreakdownProductionFn(Protocol):
    def __call__(
        self,
        price_table: dict[str, Any],
        price_rows: list | None = None,
        **kwargs: Any,
    ) -> list: ...


class DrawSegmentFn(Protocol):
    def __call__(self, ax: Any, x0: float, length: float, color: str, label: str, **kwargs: Any) -> None: ...


class DrawSplitPlateFn(Protocol):
    def __call__(
        self,
        ax: Any,
        x0: float,
        length: float,
        main_w: float,
        rest_w: float,
        label_main: str,
        **kwargs: Any,
    ) -> None: ...


class DrawTransverseCutFn(Protocol):
    def __call__(
        self,
        ax: Any,
        x0: float,
        total_length: float,
        target_length: float,
        width: float,
        label_target: str,
        remainder_length: float,
        **kwargs: Any,
    ) -> None: ...


class VisualizePlanFn(Protocol):
    def __call__(
        self,
        output_dir: str = "Визуализация_Раскладки",
        tracks_per_file: int | None = None,
        start_track_index: int = 0,
        use_production_pricing: bool = False,
        auto_import_price_to_db: bool = True,
        existing_tracks: list | None = None,
        plate_order_ctx: Any = None,
        **kwargs: Any,
    ) -> Any: ...


def register_build_layout_sequence(fn: BuildLayoutSequenceFn) -> None:
    global _build_layout_sequence
    _build_layout_sequence = fn


def register_load_price_table_from_xlsx(fn: LoadPriceTableFn) -> None:
    global _load_price_table
    _load_price_table = fn


def register_build_price_rows(fn: BuildPriceRowsFn) -> None:
    global _build_price_rows
    _build_price_rows = fn


def register_build_component_breakdown(fn: BuildComponentBreakdownFn) -> None:
    global _build_component_breakdown
    _build_component_breakdown = fn


def register_build_procurement_items(fn: BuildProcurementItemsFn) -> None:
    global _build_procurement_items
    _build_procurement_items = fn


def register_get_orders_from_opt_plan(fn: GetOrdersFromOptPlanFn) -> None:
    global _get_orders_from_opt_plan
    _get_orders_from_opt_plan = fn


def register_build_price_rows_production(fn: BuildPriceRowsProductionFn) -> None:
    global _build_price_rows_production
    _build_price_rows_production = fn


def register_build_component_breakdown_production(fn: BuildComponentBreakdownProductionFn) -> None:
    global _build_component_breakdown_production
    _build_component_breakdown_production = fn


def register_draw_segment(fn: DrawSegmentFn) -> None:
    global _draw_segment
    _draw_segment = fn


def register_draw_split_plate(fn: DrawSplitPlateFn) -> None:
    global _draw_split_plate
    _draw_split_plate = fn


def register_draw_transverse_cut(fn: DrawTransverseCutFn) -> None:
    global _draw_transverse_cut
    _draw_transverse_cut = fn


def register_visualize_plan(fn: VisualizePlanFn) -> None:
    global _visualize_plan
    _visualize_plan = fn


def reset_visualization_ports() -> None:
    """Clear registered implementations (tests)."""
    global _build_layout_sequence
    global _load_price_table
    global _build_price_rows
    global _build_component_breakdown
    global _build_procurement_items
    global _get_orders_from_opt_plan
    global _build_price_rows_production
    global _build_component_breakdown_production
    global _draw_segment
    global _draw_split_plate
    global _draw_transverse_cut
    global _visualize_plan
    _build_layout_sequence = None
    _load_price_table = None
    _build_price_rows = None
    _build_component_breakdown = None
    _build_procurement_items = None
    _get_orders_from_opt_plan = None
    _build_price_rows_production = None
    _build_component_breakdown_production = None
    _draw_segment = None
    _draw_split_plate = None
    _draw_transverse_cut = None
    _visualize_plan = None


def get_visualize_plan() -> VisualizePlanFn:
    if _visualize_plan is None:
        raise _port_not_registered("visualize_plan")
    return _visualize_plan


def build_layout_sequence(*, runtime: Any = None, **kwargs: Any) -> Any:
    if _build_layout_sequence is None:
        raise RuntimeError(
            "build_layout_sequence port is not registered. "
            "Call app.adapters.visualization.wire_visualization_ports() at startup."
        )
    return _build_layout_sequence(runtime=runtime, **kwargs)


def load_price_table_from_xlsx(xlsx_path: str) -> dict[str, Any]:
    if _load_price_table is None:
        raise RuntimeError(
            "load_price_table_from_xlsx port is not registered. "
            "Call app.adapters.visualization.wire_visualization_ports() at startup."
        )
    return _load_price_table(xlsx_path)


def _port_not_registered(name: str) -> RuntimeError:
    return RuntimeError(
        f"{name} port is not registered. "
        "Call app.adapters.visualization.wire_visualization_ports() at startup."
    )


def build_price_rows(price_table: dict[str, Any], **kwargs: Any) -> tuple[list, float]:
    if _build_price_rows is None:
        raise _port_not_registered("build_price_rows")
    return _build_price_rows(price_table, **kwargs)


def build_component_breakdown(
    price_table: dict[str, Any],
    price_rows: list | None = None,
    **kwargs: Any,
) -> list:
    if _build_component_breakdown is None:
        raise _port_not_registered("build_component_breakdown")
    return _build_component_breakdown(price_table, price_rows, **kwargs)


def build_procurement_items(**kwargs: Any) -> list:
    if _build_procurement_items is None:
        raise _port_not_registered("build_procurement_items")
    return _build_procurement_items(**kwargs)


def get_orders_from_opt_plan(opt_snapshot: Any = None) -> list | None:
    if _get_orders_from_opt_plan is None:
        raise _port_not_registered("get_orders_from_opt_plan")
    return _get_orders_from_opt_plan(opt_snapshot)


def build_price_rows_production(price_table: dict[str, Any], **kwargs: Any) -> tuple[list, float]:
    if _build_price_rows_production is None:
        raise _port_not_registered("build_price_rows_production")
    return _build_price_rows_production(price_table, **kwargs)


def build_component_breakdown_production(
    price_table: dict[str, Any],
    price_rows: list | None = None,
    **kwargs: Any,
) -> list:
    if _build_component_breakdown_production is None:
        raise _port_not_registered("build_component_breakdown_production")
    return _build_component_breakdown_production(price_table, price_rows, **kwargs)


def draw_segment(ax: Any, x0: float, length: float, color: str, label: str, **kwargs: Any) -> None:
    if _draw_segment is None:
        raise _port_not_registered("draw_segment")
    _draw_segment(ax, x0, length, color, label, **kwargs)


def draw_split_plate(
    ax: Any,
    x0: float,
    length: float,
    main_w: float,
    rest_w: float,
    label_main: str,
    **kwargs: Any,
) -> None:
    if _draw_split_plate is None:
        raise _port_not_registered("draw_split_plate")
    _draw_split_plate(ax, x0, length, main_w, rest_w, label_main, **kwargs)


def draw_transverse_cut(
    ax: Any,
    x0: float,
    total_length: float,
    target_length: float,
    width: float,
    label_target: str,
    remainder_length: float,
    **kwargs: Any,
) -> None:
    if _draw_transverse_cut is None:
        raise _port_not_registered("draw_transverse_cut")
    _draw_transverse_cut(
        ax,
        x0,
        total_length,
        target_length,
        width,
        label_target,
        remainder_length,
        **kwargs,
    )
