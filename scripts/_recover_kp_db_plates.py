#!/usr/bin/env python3
"""Recover kp_db_plates from monolithic kp_db in git HEAD, then split by function (phase 3)."""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MONOLITH = ROOT / "core" / "_kp_db_monolith_recovery.py"

HEADER_PLATES = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Plate lifecycle persistence — slice of kp_db (A1 decomposition)."""

from __future__ import annotations

import sqlite3
from collections import Counter
from datetime import datetime
from typing import Dict, List, Optional, Tuple, TypedDict

from core.kp_db_audit import audit_append
from core.kp_db_common import DEFAULT_DB, _connect


'''

COMMON_FUNCS = {
    "_normalize_plate_name",
    "_fetch_kp_plate_row_by_id",
    "_deduct_kp_plate_qty",
    "_insert_completed_plate",
    "_record_plate_completion",
    "_purge_zero_qty_plates",
}

COMPLETION_FUNCS = {
    "move_plates_to_completed",
    "check_and_update_kp_completion",
}

PLANNING_FUNCS = {
    "mark_plates_as_planned",
    "return_plates_to_production",
    "return_plate_rows_for_plan",
    "return_plan_plates_to_production",
    "recover_stuck_plates",
    "return_lost_plates_to_production",
}

QUERY_FUNCS = {
    "get_remaining_plates_for_kp",
    "get_completed_plates_for_kp",
    "get_completed_plates_stats",
    "get_completed_plates_by_day",
    "get_all_plates_in_production",
    "get_all_completed_plates",
}


def _patch_body(body: str) -> str:
    body = body.replace("_audit_append(", "audit_append(")
    return body


def extract_full_plates() -> str:
    lines = MONOLITH.read_text(encoding="utf-8-sig").splitlines(keepends=True)
    # HEAD layout: _normalize..move_plates, then rests block, then check_and_update..get_all_completed
    plates_a = "".join(lines[1462:2008])
    plates_b = "".join(lines[2444:3575])
    return HEADER_PLATES + _patch_body(plates_a + plates_b)


def _patch_move_plates(full: str) -> str:
    if "PlateCompletionService" in full:
        return full
    delegation = '''def move_plates_to_completed(
    kp_id: int,
    plates_to_complete: List[Dict],
    production_day: int,
    db_path: str = DEFAULT_DB,
    plan_ids: Optional[List[str]] = None,
    allow_cross_kp: bool = False,
    *,
    actor: str | None = None,
    return_unmoved: bool = False,
    _external_conn: Optional[sqlite3.Connection] = None,
) -> int | tuple[int, list]:
    """Facade over PlateCompletionService (A2)."""
    from core.plate_completion_service import PlateCompletionService

    return PlateCompletionService.move_plates_to_completed(
        kp_id,
        plates_to_complete,
        production_day,
        db_path,
        plan_ids,
        allow_cross_kp,
        actor=actor,
        return_unmoved=return_unmoved,
        _external_conn=_external_conn,
    )

'''
    start = full.find("def move_plates_to_completed(")
    next_def = full.find("\ndef check_and_update_kp_completion(", start)
    if start == -1 or next_def == -1:
        raise SystemExit("move_plates / check_and_update markers not found")
    return full[:start] + delegation + full[next_def + 1 :]


def _split_functions(full: str) -> dict[str, str]:
    pattern = re.compile(r"^def ([a-zA-Z_][a-zA-Z0-9_]*)\(", re.MULTILINE)
    matches = list(pattern.finditer(full))
    chunks: dict[str, str] = {}
    for i, match in enumerate(matches):
        name = match.group(1)
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(full)
        chunks[name] = full[start:end]
    return chunks


def _remove_debug_blocks(text: str) -> str:
    text = re.sub(
        r"\n\s*# #region agent log.*?# #endregion",
        "",
        text,
        flags=re.DOTALL,
    )
    return text


def split_plates(full: str) -> None:
    full = _remove_debug_blocks(full)
    chunks = _split_functions(full)

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

from core.kp_db_common import DEFAULT_DB, _connect
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

    def join_funcs(names: set[str]) -> str:
        return "".join(chunks[n] for n in chunks if n in names)

    # kp_db_plates_common.py maintained separately (A2 helpers + split DRY)
    (ROOT / "core" / "kp_db_plates_completion.py").write_text(
        completion_hdr + join_funcs(COMPLETION_FUNCS), encoding="utf-8"
    )
    (ROOT / "core" / "kp_db_plates_planning.py").write_text(
        planning_hdr + join_funcs(PLANNING_FUNCS), encoding="utf-8"
    )
    (ROOT / "core" / "kp_db_plates_queries.py").write_text(
        queries_hdr + join_funcs(QUERY_FUNCS), encoding="utf-8"
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

    missing = (
        COMPLETION_FUNCS | PLANNING_FUNCS | QUERY_FUNCS
    ) - set(chunks)
    if missing:
        raise SystemExit(f"missing functions in extract: {missing}")


def main() -> None:
    full = _patch_move_plates(extract_full_plates())
    split_plates(full)
    print("recovered and split ok", len(full), "chars")


if __name__ == "__main__":
    main()
