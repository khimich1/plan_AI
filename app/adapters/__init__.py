"""App-layer wiring of core ports to concrete adapters."""

from app.adapters.visualization import wire_visualization_ports

__all__ = ["wire_visualization_ports"]
