#!/usr/bin/env python3
"""Extract kp_db_schema, kp_db_managers, kp_db_offers for Stage 5."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
KP_DB = ROOT / "core" / "kp_db.py"
lines = KP_DB.read_text(encoding="utf-8").splitlines(keepends=True)


def slice_lines(start: int, end: int) -> str:
    """1-based inclusive start/end line numbers."""
    return "".join(lines[start - 1 : end])


# --- kp_db_schema.py ---
SCHEMA_HEADER = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SQLite schema initialization for plita.db (A1 / A4)."""

from __future__ import annotations

import os
import sqlite3
import threading

from core.kp_db_common import DEFAULT_DB, _connect

'''
schema_content = (
    SCHEMA_HEADER
    + slice_lines(52, 54)
    + slice_lines(56, 345)
    + slice_lines(348, 362)
)
(ROOT / "core" / "kp_db_schema.py").write_text(schema_content, encoding="utf-8")
print("wrote kp_db_schema.py", len(schema_content))

# --- kp_db_managers.py ---
MANAGERS_HEADER = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Managers persistence — slice of kp_db (A1 decomposition)."""

from __future__ import annotations

import json
import logging
import os
import sqlite3
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from core.kp_db_common import DEFAULT_DB, _connect
from core.kp_db_schema import ensure_schema

'''
managers_body = slice_lines(1286, 1608)
(ROOT / "core" / "kp_db_managers.py").write_text(
    MANAGERS_HEADER + managers_body,
    encoding="utf-8",
)
print("wrote kp_db_managers.py")

# --- kp_db_offers.py ---
OFFERS_HEADER = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Commercial offers (KP) persistence — slice of kp_db (A1 decomposition)."""

from __future__ import annotations

import sqlite3
import traceback
from typing import Dict, List, Optional

from core.destructive_db_guard import require_destructive_db_reset
from core.kp_db_common import DEFAULT_DB, _connect
from core.kp_db_schema import ensure_schema

'''
offers_body = (
    slice_lines(372, 1221)
    + slice_lines(1254, 1281)
    + slice_lines(1611, 1907)
)
(ROOT / "core" / "kp_db_offers.py").write_text(
    OFFERS_HEADER + offers_body,
    encoding="utf-8",
)
print("wrote kp_db_offers.py", len(offers_body))
