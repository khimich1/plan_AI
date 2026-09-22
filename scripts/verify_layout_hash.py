#!/usr/bin/env python3
"""Сверка hash раскладки по layout-полям (D-hash).

Печатает sha256 и статус match/mismatch относительно эталона.
Атрибутивные поля (concrete_grade, kp_id, kp_plate_id) в hash не входят.

Запуск из корня репозитория:
    python scripts/verify_layout_hash.py
    python scripts/verify_layout_hash.py --write-baseline
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.layout_hash import (  # noqa: E402
    build_golden_sequence,
    golden_plan,
    hash_layout_sequence,
)

def _grade_report(sequence: list) -> tuple[int, int, list[str]]:
    """Возвращает (всего root, с маркой, перечень без марки / с подозрением на дефолт)."""
    total = 0
    with_grade = 0
    problems: list[str] = []
    expected = {
        cut["plate_name"]: cut["concrete_grade"]
        for cut in golden_plan()["primary_cuts"]
    }
    for item in sequence:
        if not isinstance(item, dict):
            continue
        total += 1
        name = str(item.get("plate_name") or "")
        grade = str(item.get("concrete_grade") or "").strip()
        if grade:
            with_grade += 1
        else:
            problems.append(f"нет марки: {name or item.get('mode')}")
            continue
        want = expected.get(name)
        if want and grade != want:
            problems.append(f"{name}: {grade} != заказ {want}")
    return total, with_grade, problems

BASELINE_PATH = Path(__file__).resolve().parent / "layout_hash_baseline.sha256"


def _load_baseline() -> str | None:
    if not BASELINE_PATH.is_file():
        return None
    text = BASELINE_PATH.read_text(encoding="utf-8").strip()
    return text or None


def _write_baseline(digest: str) -> None:
    BASELINE_PATH.write_text(digest + "\n", encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Сверка hash раскладки (layout-поля)")
    parser.add_argument(
        "--write-baseline",
        action="store_true",
        help="Записать текущий hash как эталон (только для первичной фиксации)",
    )
    args = parser.parse_args(argv)

    sequence = build_golden_sequence()
    first = hash_layout_sequence(sequence)
    second = hash_layout_sequence(build_golden_sequence())
    if first != second:
        print(f"hash: {first}")
        print("status: mismatch")
        print("детерминизм: два прогона дали разный hash")
        return 1

    if args.write_baseline:
        _write_baseline(first)
        print(f"hash: {first}")
        print("status: match")
        print(f"эталон записан: {BASELINE_PATH}")
        return 0

    baseline = _load_baseline()
    if baseline is None:
        print(f"hash: {first}")
        print("status: mismatch")
        print(f"эталон не найден: {BASELINE_PATH}")
        return 1

    status = "match" if first == baseline else "mismatch"
    print(f"hash: {first}")
    print(f"baseline: {baseline}")
    print(f"status: {status}")

    total, with_grade, grade_problems = _grade_report(sequence)
    print(f"grades: {with_grade}/{total} root items с маркой")
    for problem in grade_problems:
        print(f"grade-mismatch: {problem}")
    if grade_problems or total == 0 or with_grade != total:
        print("grades-status: mismatch")
        return 1
    print("grades-status: match")
    return 0 if status == "match" else 1


if __name__ == "__main__":
    raise SystemExit(main())
