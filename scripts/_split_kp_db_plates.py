#!/usr/bin/env python3
"""One-off splitter for kp_db_plates (phase 3 stage 16)."""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
src_path = ROOT / "core" / "kp_db_plates.py"
lines = src_path.read_text(encoding="utf-8").splitlines(keepends=True)


def slice_lines(start: int, end: int) -> str:
    return "".join(lines[start - 1 : end])


HEADER = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""{doc}"""

from __future__ import annotations

'''

common_hdr = HEADER.format(doc="Shared plate persistence helpers (A1 slice).") + """import sqlite3
from typing import List, Optional

from core.kp_db_common import DEFAULT_DB, _connect

"""

completion_hdr = HEADER.format(doc="Plate completion persistence (A1 slice).") + """import sqlite3
from typing import Dict, List, Optional

from core.kp_db_audit import audit_append
from core.kp_db_common import DEFAULT_DB, _connect
from core.domain.plate_completion_types import UnmovedPlateInfo
from core.kp_db_plates_common import (
    _deduct_kp_plate_qty,
    _fetch_kp_plate_row_by_id,
    _insert_completed_plate,
    _normalize_plate_name,
    _purge_zero_qty_plates,
    _record_plate_completion,
)

"""

planning_hdr = HEADER.format(doc="Plate planning and status transitions (A1 slice).") + """import sqlite3
from collections import Counter
from typing import Dict, List, Optional, Tuple

from core.kp_db_audit import audit_append
from core.kp_db_common import DEFAULT_DB, _connect

"""

queries_hdr = HEADER.format(doc="Read-only plate queries (A1 slice).") + """import sqlite3
from typing import Dict, List

from core.kp_db_common import DEFAULT_DB, _connect

"""

(ROOT / "core" / "kp_db_plates_common.py").write_text(
    common_hdr + slice_lines(19, 157), encoding="utf-8"
)
(ROOT / "core" / "kp_db_plates_completion.py").write_text(
    completion_hdr + slice_lines(160, 244), encoding="utf-8"
)
(ROOT / "core" / "kp_db_plates_planning.py").write_text(
    planning_hdr + slice_lines(284, 1101), encoding="utf-8"
)
(ROOT / "core" / "kp_db_plates_queries.py").write_text(
    queries_hdr + slice_lines(247, 277) + slice_lines(1104, 1302),
    encoding="utf-8",
)

shim = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Plate lifecycle persistence — re-export shim (A1 decomposition)."""

from core.kp_db_plates_common import *  # noqa: F403
from core.kp_db_plates_completion import *  # noqa: F403
from core.kp_db_plates_planning import *  # noqa: F403
from core.kp_db_plates_queries import *  # noqa: F403
'''
(ROOT / "core" / "kp_db_plates.py").write_text(shim, encoding="utf-8")
print("split ok")
