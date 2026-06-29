#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Entry point for the deprecated Telegram bot (stub only).

Run from project root: python run_bot.py
See bot/README.md for decommission details (P5 WP1, 2026-06-21).
"""

from __future__ import annotations

import sys

_DEPRECATION_MESSAGE = (
    "DEPRECATED: Telegram bot is soft-decommissioned (2026-06-21).\n"
    "Production path is the web app. See bot/README.md and bot_archived/.\n"
    "Hard delete is planned for P6."
)


def main() -> int:
    print(_DEPRECATION_MESSAGE, file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
