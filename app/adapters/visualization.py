"""Wire visualization ports at application startup."""

from __future__ import annotations

from viz_modules.adapters.visualization_ports import register_default_visualization_ports


def wire_visualization_ports() -> None:
    """Register viz_modules as the default implementation of core visualization ports."""
    register_default_visualization_ports()
