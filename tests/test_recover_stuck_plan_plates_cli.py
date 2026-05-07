"""CLI safety tests for scripts/recover_stuck_plan_plates.py (subprocess isolation)."""

from __future__ import annotations

import os
import sqlite3
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from core import kp_db


def _subprocess_env() -> dict[str, str]:
    env = os.environ.copy()
    env.setdefault("PYTHONUTF8", "1")
    env.setdefault("PYTHONIOENCODING", "utf-8")
    return env


def test_recover_apply_without_plan_id_exits_code_2_no_apply_path() -> None:
    """--apply without --plan-id must fail via argparse before any recover work."""
    result = subprocess.run(
        [sys.executable, "scripts/recover_stuck_plan_plates.py", "--apply"],
        cwd=REPO_ROOT,
        env=_subprocess_env(),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    assert result.returncode == 2
    combined = (result.stderr or "") + (result.stdout or "")
    assert "--plan-id" in combined or "plan-id" in combined.lower()


def test_recover_help_exits_zero() -> None:
    result = subprocess.run(
        [sys.executable, "scripts/recover_stuck_plan_plates.py", "--help"],
        cwd=REPO_ROOT,
        env=_subprocess_env(),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    assert result.returncode == 0
    assert "usage:" in (result.stdout or "").lower()


def test_recover_dry_run_explicit_db_no_stuck_plans_exits_zero(
    tmp_path: Path,
) -> None:
    """Dry-run without --plan-id uses DB listing; empty schema → no stuck plans."""
    db_path = tmp_path / "plita.db"
    kp_db.init_schema(str(db_path))
    result = subprocess.run(
        [
            sys.executable,
            "scripts/recover_stuck_plan_plates.py",
            "--db-path",
            str(db_path),
        ],
        cwd=REPO_ROOT,
        env=_subprocess_env(),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    assert result.returncode == 0
    out = result.stdout or ""
    assert "Зависших плит" in out and "не найдено" in out


def test_recover_dry_run_explicit_plan_id_zero_stuck_exits_zero(
    tmp_path: Path,
) -> None:
    """Explicit --plan-id always enters the plan loop (plan JSON may be missing)."""
    db_path = tmp_path / "plita.db"
    kp_db.init_schema(str(db_path))
    plan_id = "__cli_test_no_stuck_rows__"
    result = subprocess.run(
        [
            sys.executable,
            "scripts/recover_stuck_plan_plates.py",
            "--db-path",
            str(db_path),
            "--plan-id",
            plan_id,
        ],
        cwd=REPO_ROOT,
        env=_subprocess_env(),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    assert result.returncode == 0
    out = result.stdout or ""
    assert plan_id in out
    assert "зависло 0 плит" in out
    assert "Dry-run" in out or "изменения не применялись" in out


def test_recover_apply_with_plan_id_returns_plates_temp_db(
    tmp_path: Path,
) -> None:
    """--apply + --plan-id invokes kp_db return helper on isolated DB (no prod DB)."""
    db_path = tmp_path / "plita.db"
    kp_db.init_schema(str(db_path))
    plan_id = "__cli_apply_isolated_test__"

    conn = sqlite3.connect(str(db_path))
    try:
        conn.execute(
            "INSERT INTO KP_offers "
            "(kp_id, creation_date, execution_terms, customer_name) "
            "VALUES (1, '2026-01-01', '-', 'ТестКлиент')"
        )
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name,
                length_m, width_m, load_class,
                qty, status, plan_id
            )
            VALUES (1, 1, 'ПБ CLI-ТЕСТ', 6.0, 1.2, 800, 2, 'в плане', ?)
            """,
            (plan_id,),
        )
        conn.commit()
    finally:
        conn.close()

    result = subprocess.run(
        [
            sys.executable,
            "scripts/recover_stuck_plan_plates.py",
            "--db-path",
            str(db_path),
            "--plan-id",
            plan_id,
            "--apply",
        ],
        cwd=REPO_ROOT,
        env=_subprocess_env(),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    assert result.returncode == 0
    combined = (result.stdout or "") + (result.stderr or "")
    assert "Итого возвращено в производство" in combined

    with sqlite3.connect(str(db_path)) as conn2:
        row = conn2.execute(
            "SELECT status, plan_id FROM kp_plates WHERE kp_id = 1"
        ).fetchone()
    assert row == ("в производстве", None)
