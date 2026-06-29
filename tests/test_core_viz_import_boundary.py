"""WP3/WP4 A1: grep gate — core must not import viz_modules on migrated hot paths."""

from __future__ import annotations

from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]


def _read(rel: str) -> str:
    return (REPO_ROOT / rel).read_text(encoding="utf-8")


def test_production_planning_has_no_viz_modules_import() -> None:
    source = _read("core/production/planning.py")
    assert "viz_modules" not in source


def test_visualization_has_no_viz_modules_import() -> None:
    source = _read("core/visualization/__init__.py")
    layout_source = _read("core/visualization/layout.py")
    assert "viz_modules" not in source
    assert "viz_modules" not in layout_source


def test_visualization_migrated_ports_not_direct_viz_imports() -> None:
    source = _read("core/visualization/__init__.py")
    assert "from viz_modules.layout_sequence import" not in source
    assert "from viz_modules.price_utils import" not in source
    assert "from viz_modules.procurement import" not in source
    assert "from viz_modules.visualization_drawing import" not in source


def test_app_layer_only_adapter_imports_viz_modules() -> None:
    app_root = REPO_ROOT / "app"
    offenders: list[str] = []
    for path in app_root.rglob("*.py"):
        rel = path.relative_to(REPO_ROOT).as_posix()
        if rel == "app/adapters/visualization.py":
            continue
        source = path.read_text(encoding="utf-8")
        if "viz_modules" in source:
            offenders.append(rel)
    assert offenders == []


def test_app_services_do_not_import_visualize_plan_directly() -> None:
    app_root = REPO_ROOT / "app"
    needle = "from core.visualization import visualize_plan"
    offenders: list[str] = []
    for path in app_root.rglob("*.py"):
        source = path.read_text(encoding="utf-8")
        if needle in source:
            offenders.append(path.relative_to(REPO_ROOT).as_posix())
    assert offenders == []


def test_visualization_ports_registered_in_pytest() -> None:
    from core.ports.visualization import (
        build_component_breakdown,
        build_component_breakdown_production,
        build_layout_sequence,
        build_price_rows,
        build_price_rows_production,
        build_procurement_items,
        draw_segment,
        draw_split_plate,
        draw_transverse_cut,
        get_orders_from_opt_plan,
        get_visualize_plan,
        load_price_table_from_xlsx,
    )

    # conftest pytest_configure wires defaults; facades must be callable.
    assert callable(build_layout_sequence)
    assert callable(load_price_table_from_xlsx)
    assert callable(build_price_rows)
    assert callable(build_component_breakdown)
    assert callable(build_procurement_items)
    assert callable(get_orders_from_opt_plan)
    assert callable(build_price_rows_production)
    assert callable(build_component_breakdown_production)
    assert callable(draw_segment)
    assert callable(draw_split_plate)
    assert callable(draw_transverse_cut)
    assert callable(get_visualize_plan())
