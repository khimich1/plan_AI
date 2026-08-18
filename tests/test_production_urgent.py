"""Unit tests for core.production.urgent (pure domain, no app I/O)."""

from __future__ import annotations

import ast
from datetime import date, datetime, timedelta
from pathlib import Path

from core.execution_terms import DEFAULT_EXECUTION_TERMS_DAYS
from core.production.urgent import UrgentPosition, collect_urgent_positions

FIXED_NOW = datetime(2026, 8, 1, 12, 0, 0)


def test_batch_only_urgent() -> None:
    plates = [
        {
            "plate_id": 10,
            "kp_id": 1,
            "plate_name": "ПБ 60",
            "qty_remaining": 5,
        }
    ]
    batches = {
        10: [
            {"produce_by": "2026-08-10", "qty": 2, "batch_name": "П1"},
            {"produce_by": date(2026, 8, 15), "qty": 3, "batch_name": "П2"},
        ]
    }
    result = collect_urgent_positions(
        plates, batches, {1: {}}, date(2026, 8, 20), now=FIXED_NOW
    )
    assert len(result) == 1
    pos = result[0]
    assert isinstance(pos, UrgentPosition)
    assert pos.deadline == date(2026, 8, 10)
    assert pos.deadline_source == "delivery_batch"
    assert pos.qty_remaining == 5
    assert pos.conflict is None
    assert len(pos.deadline_details) == 2
    assert pos.deadline_details[0]["type"] == "delivery_batch"
    assert pos.deadline_details[0]["deadline"] == "2026-08-10"


def test_terms_only_urgent() -> None:
    plates = [
        {"id": 20, "kp_id": 2, "plate_name": "ПБ 50", "qty_remaining": 3}
    ]
    result = collect_urgent_positions(
        plates,
        {},
        {2: {"execution_terms": "10.08.2026"}},
        date(2026, 8, 20),
        now=FIXED_NOW,
    )
    assert len(result) == 1
    pos = result[0]
    assert pos.plate_id == 20
    assert pos.deadline == date(2026, 8, 10)
    assert pos.deadline_source == "execution_terms"
    assert pos.deadline_details == [
        {"type": "execution_terms", "deadline": "2026-08-10", "qty": 3}
    ]
    assert pos.conflict is None


def test_batches_win_as_primary_details_include_both() -> None:
    plates = [
        {
            "plate_id": 30,
            "kp_id": 3,
            "plate_name": "ПБ 40",
            "qty_remaining": 7,
        }
    ]
    batches = {
        30: [{"produce_by": "2026-08-05", "qty": 4, "batch_name": "A"}]
    }
    result = collect_urgent_positions(
        plates,
        batches,
        {3: {"execution_terms": "12.08.2026"}},
        date(2026, 8, 20),
        now=FIXED_NOW,
    )
    assert len(result) == 1
    pos = result[0]
    assert pos.deadline == date(2026, 8, 5)
    assert pos.deadline_source == "delivery_batch"
    types = [d["type"] for d in pos.deadline_details]
    assert types == ["delivery_batch", "execution_terms"]
    assert pos.deadline_details[1]["deadline"] == "2026-08-12"
    assert pos.deadline_details[1]["qty"] == 7
    assert pos.conflict is None  # 7 days exactly → no conflict


def test_filter_by_deadline_until() -> None:
    plates = [
        {"plate_id": 1, "kp_id": 1, "plate_name": "A", "qty_remaining": 1},
        {"plate_id": 2, "kp_id": 1, "plate_name": "B", "qty_remaining": 1},
    ]
    batches = {
        1: [{"produce_by": "2026-08-10", "qty": 1}],
        2: [{"produce_by": "2026-08-25", "qty": 1}],
    }
    result = collect_urgent_positions(
        plates, batches, {}, date(2026, 8, 15), now=FIXED_NOW
    )
    assert [p.plate_id for p in result] == [1]


def test_conflict_schedule_earlier() -> None:
    plates = [
        {"plate_id": 40, "kp_id": 4, "plate_name": "X", "qty_remaining": 2}
    ]
    batches = {40: [{"produce_by": "2026-08-01", "qty": 2}]}
    result = collect_urgent_positions(
        plates,
        batches,
        {4: {"execution_terms": "15.08.2026"}},
        date(2026, 8, 30),
        now=FIXED_NOW,
    )
    assert result[0].conflict == "schedule_earlier"
    assert result[0].deadline_source == "delivery_batch"


def test_conflict_kp_earlier() -> None:
    plates = [
        {"plate_id": 50, "kp_id": 5, "plate_name": "Y", "qty_remaining": 2}
    ]
    batches = {50: [{"produce_by": "2026-08-20", "qty": 2}]}
    result = collect_urgent_positions(
        plates,
        batches,
        {5: {"execution_terms": "05.08.2026"}},
        date(2026, 8, 30),
        now=FIXED_NOW,
    )
    assert result[0].conflict == "kp_earlier"
    assert result[0].deadline == date(2026, 8, 20)


def test_no_conflict_when_within_7_days() -> None:
    plates = [
        {"plate_id": 60, "kp_id": 6, "plate_name": "Z", "qty_remaining": 1}
    ]
    batches = {60: [{"produce_by": "2026-08-10", "qty": 1}]}
    result = collect_urgent_positions(
        plates,
        batches,
        {6: {"execution_terms": "17.08.2026"}},
        date(2026, 8, 30),
        now=FIXED_NOW,
    )
    assert result[0].conflict is None


def test_conflict_at_8_days_boundary() -> None:
    """AC: conflict only when |batch − KP| > 7 days (8 days must flag)."""
    plates = [
        {"plate_id": 61, "kp_id": 6, "plate_name": "B8", "qty_remaining": 1}
    ]
    batches = {61: [{"produce_by": "2026-08-10", "qty": 1}]}
    result = collect_urgent_positions(
        plates,
        batches,
        {6: {"execution_terms": "18.08.2026"}},
        date(2026, 8, 30),
        now=FIXED_NOW,
    )
    assert result[0].conflict == "schedule_earlier"


def test_empty_produce_by_skipped_falls_back_to_terms() -> None:
    """Batches with null/empty produce_by must not count as delivery deadlines."""
    plates = [
        {
            "plate_id": 62,
            "kp_id": 6,
            "plate_name": "SkipEmpty",
            "qty_remaining": 2,
        }
    ]
    batches = {
        62: [
            {"produce_by": None, "qty": 1, "batch_name": "empty"},
            {"produce_by": "", "qty": 1, "batch_name": "blank"},
        ]
    }
    result = collect_urgent_positions(
        plates,
        batches,
        {6: {"execution_terms": "10.08.2026"}},
        date(2026, 8, 20),
        now=FIXED_NOW,
    )
    assert len(result) == 1
    pos = result[0]
    assert pos.deadline == date(2026, 8, 10)
    assert pos.deadline_source == "execution_terms"
    assert pos.conflict is None
    assert all(d["type"] == "execution_terms" for d in pos.deadline_details)


def test_deadline_until_is_inclusive() -> None:
    plates = [
        {"plate_id": 63, "kp_id": 6, "plate_name": "Eq", "qty_remaining": 1}
    ]
    batches = {63: [{"produce_by": "2026-08-15", "qty": 1}]}
    included = collect_urgent_positions(
        plates, batches, {}, date(2026, 8, 15), now=FIXED_NOW
    )
    assert [p.plate_id for p in included] == [63]
    excluded = collect_urgent_positions(
        plates, batches, {}, date(2026, 8, 14), now=FIXED_NOW
    )
    assert excluded == []


def test_unparseable_terms_no_batches_fallback_plus_14() -> None:
    plates = [
        {
            "plate_id": 70,
            "kp_id": 7,
            "plate_name": "Fallback",
            "qty_remaining": 4,
        }
    ]
    until = date(2026, 8, 1) + timedelta(days=DEFAULT_EXECUTION_TERMS_DAYS)
    result = collect_urgent_positions(
        plates,
        {},
        {7: {"execution_terms": "как можно скорее"}},
        until,
        now=FIXED_NOW,
    )
    assert len(result) == 1
    pos = result[0]
    assert pos.deadline == date(2026, 8, 15)  # 2026-08-01 + 14
    assert pos.deadline_source == "execution_terms"
    assert pos.conflict is None
    # Outside until → excluded
    excluded = collect_urgent_positions(
        plates,
        {},
        {7: {"execution_terms": "как можно скорее"}},
        date(2026, 8, 14),
        now=FIXED_NOW,
    )
    assert excluded == []


def test_fallback_not_used_for_conflict() -> None:
    """Unparseable KP terms must not create a conflict vs batch dates."""
    plates = [
        {"plate_id": 80, "kp_id": 8, "plate_name": "C", "qty_remaining": 1}
    ]
    batches = {80: [{"produce_by": "2026-08-01", "qty": 1}]}
    result = collect_urgent_positions(
        plates,
        batches,
        {8: {"execution_terms": "непонятно"}},
        date(2026, 8, 30),
        now=FIXED_NOW,
    )
    assert result[0].deadline_source == "delivery_batch"
    assert result[0].conflict is None
    assert all(d["type"] == "delivery_batch" for d in result[0].deadline_details)


def test_sort_order() -> None:
    plates = [
        {"plate_id": 3, "kp_id": 2, "plate_name": "c", "qty_remaining": 1},
        {"plate_id": 1, "kp_id": 1, "plate_name": "a", "qty_remaining": 1},
        {"plate_id": 2, "kp_id": 1, "plate_name": "b", "qty_remaining": 1},
        {"plate_id": 4, "kp_id": 1, "plate_name": "d", "qty_remaining": 1},
    ]
    batches = {
        1: [{"produce_by": "2026-08-12", "qty": 1}],
        2: [{"produce_by": "2026-08-10", "qty": 1}],
        3: [{"produce_by": "2026-08-10", "qty": 1}],
        4: [{"produce_by": "2026-08-12", "qty": 1}],
    }
    result = collect_urgent_positions(
        plates, batches, {}, date(2026, 8, 30), now=FIXED_NOW
    )
    assert [(p.deadline.isoformat(), p.kp_id, p.plate_id) for p in result] == [
        ("2026-08-10", 1, 2),
        ("2026-08-10", 2, 3),
        ("2026-08-12", 1, 1),
        ("2026-08-12", 1, 4),
    ]


def test_relative_execution_terms_uses_now() -> None:
    plates = [
        {"plate_id": 90, "kp_id": 9, "plate_name": "Rel", "qty_remaining": 1}
    ]
    result = collect_urgent_positions(
        plates,
        {},
        {9: {"execution_terms": "5 дней"}},
        date(2026, 8, 30),
        now=FIXED_NOW,
    )
    assert result[0].deadline == date(2026, 8, 6)


def test_urgent_module_has_no_app_imports() -> None:
    path = Path(__file__).resolve().parents[1] / "core" / "production" / "urgent.py"
    tree = ast.parse(path.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                assert not alias.name.startswith("app"), alias.name
        elif isinstance(node, ast.ImportFrom):
            module = node.module or ""
            assert not module.startswith("app"), module
