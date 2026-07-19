# -*- coding: utf-8 -*-
"""Narrow dependency-inversion ports for the optimizer (read-only cfg/order views)."""

from __future__ import annotations

from core.optimization.ports.order_data import (
    ConfigModulePlateOrderAdapter,
    PlateOrderDataPort,
    default_plate_order_port,
    resolve_plate_order_port,
)

__all__ = (
    "ConfigModulePlateOrderAdapter",
    "PlateOrderDataPort",
    "default_plate_order_port",
    "resolve_plate_order_port",
)
