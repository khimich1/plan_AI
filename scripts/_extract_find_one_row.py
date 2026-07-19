from pathlib import Path

lines = Path("core/kp_db_plates.py").read_text(encoding="utf-8").splitlines()
body = lines[138:344]
dedented = [ln[8:] if ln.startswith("        ") else ln for ln in body]
body_text = "\n".join(dedented)
body_text = body_text.replace("_debug_log4", "_DEBUG_LOG")

header = '''"""Domain matching strategies for kp_plates row lookup (A2)."""

from __future__ import annotations

import sqlite3
from typing import Sequence

from core.debug_paths import append_agent_debug_log, get_debug_log_path

_DEBUG_LOG = get_debug_log_path("debug.log")
_DEBUG_LOG_8E9428 = get_debug_log_path("debug-8e9428.log")


def _normalize_plate_name(name: str) -> str:
    from core import plate_name as _pn
    return _pn.canonical(name)


def find_kp_plate_row(
    cur: sqlite3.Cursor,
    plate_name: str,
    length_m: float,
    width_m: float,
    load_class: int,
    prefer_kp_id: int,
    *,
    length_dm_raw: str | None = None,
    allow_cross_kp: bool = False,
    plan_ids: Sequence[str] | None = None,
) -> tuple | None:
    """Find one kp_plates row for write-off (steps 0-7)."""
'''

Path("core/domain/plate_completion_matching.py").write_text(
    header + body_text + "\n",
    encoding="utf-8",
)
print("ok", len(dedented))
