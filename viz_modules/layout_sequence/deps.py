# -*- coding: utf-8 -*-
"""DI для layout_sequence: путь БД прайса, логгер, пути трассировки."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path

from core.project_paths import PRICE_DB_PATH

from viz_modules.layout_sequence.debug_trace import LayoutSequenceTracePaths


@dataclass
class LayoutSequenceDeps:
    pb_db_path: Path
    log: logging.Logger
    traces: LayoutSequenceTracePaths

    @classmethod
    def create(
        cls,
        pb_db_path: Path | str | None = None,
        log: logging.Logger | None = None,
        traces: LayoutSequenceTracePaths | None = None,
    ) -> LayoutSequenceDeps:
        resolved_db = Path(pb_db_path) if pb_db_path is not None else Path(PRICE_DB_PATH)
        resolved_log = log or logging.getLogger("viz_modules.layout_sequence")
        resolved_traces = traces or LayoutSequenceTracePaths.default()
        return cls(pb_db_path=resolved_db, log=resolved_log, traces=resolved_traces)
