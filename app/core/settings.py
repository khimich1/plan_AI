# -*- coding: utf-8 -*-
"""Backward-compatible entry: settings are defined in ``core.config.settings``."""

from core.config.settings import Settings, get_settings

__all__ = ["Settings", "get_settings"]
