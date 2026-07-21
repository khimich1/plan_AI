"""Extract plates/rests slices from kp_db.py into separate modules."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
KP_DB = ROOT / "core" / "kp_db.py"

HEADER_PLATES = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Plate lifecycle persistence — slice of kp_db (A1 decomposition)."""

from __future__ import annotations

import sqlite3
from collections import Counter
from datetime import datetime
from typing import Dict, List, Optional, Tuple, TypedDict

from core.debug_paths import (
    append_agent_debug_log,
    append_agent_debug_session,
    get_debug_log_path,
)
from core.destructive_db_guard import require_destructive_db_reset
from core.kp_db_common import DEFAULT_DB, _audit_append, _connect

_DEBUG_SESSION_LOG = get_debug_log_path("debug-d7e22e.log")
_DEBUG_NOMENCLATURE_LOG = get_debug_log_path("debug-00f316.log")
_DEBUG_LOG = get_debug_log_path("debug.log")
_DEBUG_AGENT_LOG = get_debug_log_path("debug-ebb546.log")
_DEBUG_LOG_A9176E = get_debug_log_path("debug-a9176e.log")
_DEBUG_LOG_B59370 = get_debug_log_path("debug-b59370.log")
_DEBUG_LOG_8E9428 = get_debug_log_path("debug-8e9428.log")


def _init_schema(db_path: str) -> None:
    from core import kp_db

    kp_db.init_schema(db_path)


def _debug_session_write(run_id: str, hypothesis_id: str, location: str, message: str, data: Dict) -> None:
    append_agent_debug_session(
        _DEBUG_SESSION_LOG,
        run_id=run_id,
        hypothesis_id=hypothesis_id,
        location=location,
        message=message,
        data=data,
    )


'''

HEADER_RESTS = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Plate rests persistence — slice of kp_db (A1 decomposition)."""

from __future__ import annotations

import sqlite3
from typing import Dict, List, Optional

from core.kp_db_common import DEFAULT_DB, _connect


def _init_schema(db_path: str) -> None:
    from core import kp_db

    kp_db.init_schema(db_path)


'''


def _patch_body(body: str) -> str:
    body = body.replace("init_schema(db_path)", "_init_schema(db_path)")
    body = body.replace("init_schema(self.db_path)", "_init_schema(self.db_path)")
    return body


def main() -> None:
    lines = KP_DB.read_text(encoding="utf-8").splitlines(keepends=True)

    plates_a = "".join(lines[1261:1887])  # through end of move_plates
    plates_b = "".join(lines[2326:3484])  # check_and_update .. get_all_completed
    rests = "".join(lines[1888:2325])

    plates_body = _patch_body(plates_a + plates_b)
    rests_body = _patch_body(rests)

    (ROOT / "core" / "kp_db_plates.py").write_text(
        HEADER_PLATES + plates_body,
        encoding="utf-8",
    )
    (ROOT / "core" / "kp_db_rests.py").write_text(
        HEADER_RESTS + rests_body,
        encoding="utf-8",
    )
    print("wrote kp_db_plates.py", len(plates_body), "chars")
    print("wrote kp_db_rests.py", len(rests_body), "chars")


if __name__ == "__main__":
    main()
