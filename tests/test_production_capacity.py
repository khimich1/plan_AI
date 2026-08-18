"""Unit tests for core.production.capacity (pure domain, no app I/O)."""

from __future__ import annotations

import ast
from datetime import date
from pathlib import Path

import pytest

from core.production.capacity import (
    TRACKS_PER_DAY_HARD_CAP,
    CapacityDeficit,
    CapacityOption,
    calculate_capacity_deficit,
    clamp_day_max,
    day_free_tracks,
    get_day_capacity,
    validate_fill_targets,
)
from core.production.errors import PlanBuildError
from core.production_capacity import MAX_TRACK_LENGTH_M, TRACKS_PER_DAY_DEFAULT


def test_get_day_capacity_default() -> None:
    assert get_day_capacity("2026-08-12", {}) == int(TRACKS_PER_DAY_DEFAULT)
    assert get_day_capacity(date(2026, 8, 12), {}) == int(TRACKS_PER_DAY_DEFAULT)


def test_get_day_capacity_override_by_str_and_date_keys() -> None:
    overrides = {
        "2026-08-12": 4,
        date(2026, 8, 13): 3,
    }
    assert get_day_capacity("2026-08-12", overrides) == 4
    assert get_day_capacity(date(2026, 8, 13), overrides) == 3
    assert get_day_capacity("2026-08-14", overrides) == int(TRACKS_PER_DAY_DEFAULT)


def test_get_day_capacity_clamps_above_hard_cap() -> None:
    assert get_day_capacity("2026-08-12", {"2026-08-12": 8}) == TRACKS_PER_DAY_HARD_CAP
    assert get_day_capacity("2026-08-12", {}, default=9) == TRACKS_PER_DAY_HARD_CAP
    assert clamp_day_max(0) == 0
    assert clamp_day_max(5) == 5
    assert clamp_day_max(99) == 5


def test_get_day_capacity_normalizes_mixed_date_key_types() -> None:
    """Day arg and override key may differ (str vs date); both normalize to ISO."""
    assert get_day_capacity(date(2026, 8, 12), {"2026-08-12": 4}) == 4
    assert get_day_capacity("2026-08-13", {date(2026, 8, 13): 3}) == 3


def test_get_day_capacity_explicit_default() -> None:
    assert get_day_capacity("2026-08-12", {}, default=3) == 3
    assert get_day_capacity("2026-08-12", {"2026-08-12": 2}, default=3) == 2


def test_day_free_tracks_respects_occupancy_and_cap() -> None:
    assert day_free_tracks("2026-08-12", {"2026-08-12": 5}, {"2026-08-12": 2}) == 3
    assert day_free_tracks("2026-08-12", {"2026-08-12": 8}, {"2026-08-12": 1}) == 4
    assert day_free_tracks("2026-08-12", {"2026-08-12": 0}, {}) == 0


def test_validate_fill_targets_ok() -> None:
    validate_fill_targets(
        [{"date": "2026-08-12", "tracks": 3}, {"date": "2026-08-13", "tracks": 5}],
        {"2026-08-12": 5, "2026-08-13": 5},
    )


def test_validate_fill_targets_empty_raises() -> None:
    with pytest.raises(PlanBuildError, match="fill_targets пуст"):
        validate_fill_targets([], {"2026-08-12": 5})


def test_validate_fill_targets_over_capacity_raises() -> None:
    with pytest.raises(PlanBuildError, match=r"свободно 5.*запрошено 6"):
        validate_fill_targets(
            [{"date": "2026-08-12", "tracks": 6}],
            {"2026-08-12": 5},
        )


def test_validate_fill_targets_clamps_stale_override_above_cap() -> None:
    with pytest.raises(PlanBuildError, match=r"свободно 5.*запрошено 6"):
        validate_fill_targets(
            [{"date": "2026-08-12", "tracks": 6}],
            {"2026-08-12": 9},
        )


def test_validate_fill_targets_uses_default_when_date_missing() -> None:
    default = int(TRACKS_PER_DAY_DEFAULT)
    validate_fill_targets(
        [{"date": "2026-08-12", "tracks": default}],
        {},
    )
    with pytest.raises(PlanBuildError, match="свободно"):
        validate_fill_targets(
            [{"date": "2026-08-12", "tracks": default + 1}],
            {},
        )


def test_validate_fill_targets_invalid_date_and_tracks() -> None:
    with pytest.raises(PlanBuildError, match="Неверный формат даты"):
        validate_fill_targets(
            [{"date": "12-08-2026", "tracks": 1}],
            {"2026-08-12": 5},
        )
    with pytest.raises(PlanBuildError, match="должно быть >= 1"):
        validate_fill_targets(
            [{"date": "2026-08-12", "tracks": 0}],
            {"2026-08-12": 5},
        )


def test_validate_fill_targets_respects_occupancy() -> None:
    """occupancy=3, max=5 → free=2; tracks=2 OK; tracks=3/4 → PlanBuildError.

    Plan text said «tracks 3 OK»; free slots = max − occupied = 2, so tracks≤2.
    """
    day_capacity = {"2026-09-10": 5}
    occupancy = {"2026-09-10": 3}
    validate_fill_targets(
        [{"date": "2026-09-10", "tracks": 2}],
        day_capacity,
        occupancy=occupancy,
    )
    with pytest.raises(PlanBuildError, match=r"свободно 2.*запрошено 3"):
        validate_fill_targets(
            [{"date": "2026-09-10", "tracks": 3}],
            day_capacity,
            occupancy=occupancy,
        )
    with pytest.raises(PlanBuildError, match=r"свободно 2.*запрошено 4"):
        validate_fill_targets(
            [{"date": "2026-09-10", "tracks": 4}],
            day_capacity,
            occupancy=occupancy,
        )


def test_validate_fill_targets_occupancy_default_empty() -> None:
    """Without occupancy arg, free == day_max (same as empty map)."""
    validate_fill_targets(
        [{"date": "2026-09-10", "tracks": 5}],
        {"2026-09-10": 5},
    )
    validate_fill_targets(
        [{"date": "2026-09-10", "tracks": 5}],
        {"2026-09-10": 5},
        occupancy={},
    )


def test_calculate_capacity_deficit_none_when_enough() -> None:
    result = calculate_capacity_deficit(
        urgent_length_m=2 * MAX_TRACK_LENGTH_M,
        fill_targets=[{"date": "2026-08-12", "tracks": 2}],
        day_capacity={"2026-08-12": 5},
    )
    assert result is None


def test_calculate_capacity_deficit_tracks_missing_uses_ceil() -> None:
    fill_targets = [{"date": "2026-08-12", "tracks": 1}]
    day_capacity = {"2026-08-12": 5}
    urgent = 2 * MAX_TRACK_LENGTH_M + 0.01  # just over 2 tracks → need 3
    result = calculate_capacity_deficit(urgent, fill_targets, day_capacity)
    assert result is not None
    assert result.tracks_needed == 3
    assert result.tracks_available == 1
    assert result.tracks_missing == 2
    assert result.options[0] == CapacityOption(
        action="bump_fill",
        date="2026-08-12",
        add_tracks=4,
        free=5,
    )


def test_calculate_capacity_deficit_zero_and_negative_length() -> None:
    assert (
        calculate_capacity_deficit(
            0.0,
            [{"date": "2026-08-12", "tracks": 1}],
            {"2026-08-12": 5},
        )
        is None
    )
    assert (
        calculate_capacity_deficit(
            -10.0,
            [{"date": "2026-08-12", "tracks": 1}],
            {"2026-08-12": 5},
        )
        is None
    )


def test_calculate_capacity_deficit_options_a_bump_fill_by_date() -> None:
    fill_targets = [
        {"date": "2026-08-12", "tracks": 4},
        {"date": "2026-08-13", "tracks": 1},
    ]
    day_capacity = {"2026-08-12": 5, "2026-08-13": 5}
    urgent = 6 * MAX_TRACK_LENGTH_M  # need 6, available 5, missing 1
    result = calculate_capacity_deficit(urgent, fill_targets, day_capacity)
    assert result is not None
    assert isinstance(result, CapacityDeficit)
    assert result.tracks_missing == 1
    assert result.deficit_until == "2026-08-13"
    assert result.options[0] == CapacityOption(
        action="bump_fill", date="2026-08-12", add_tracks=1, free=5
    )
    assert result.options[1] == CapacityOption(
        action="bump_fill", date="2026-08-13", add_tracks=4, free=5
    )


def test_calculate_capacity_deficit_options_b_previous_then_c_future() -> None:
    fill_targets = [{"date": "2026-08-14", "tracks": 5}]
    day_capacity = {
        "2026-08-12": 5,
        "2026-08-13": 5,
        "2026-08-14": 5,
        "2026-08-17": 5,
    }
    urgent = 8 * MAX_TRACK_LENGTH_M
    result = calculate_capacity_deficit(
        urgent,
        fill_targets,
        day_capacity,
        occupancy={"2026-08-12": 2, "2026-08-13": 5, "2026-08-17": 0},
        completed_dates={"2026-08-11"},
        today="2026-08-12",
        is_workday=lambda d: d.weekday() < 5,
    )
    assert result is not None
    assert result.tracks_missing == 3
    # A: no headroom on selected (free=5, fill=5)
    assert all(o.action != "bump_fill" for o in result.options)
    # B: 12 has free=3; 13 free=0 skipped; weekend 15-16 skipped later in C
    assert result.options[0] == CapacityOption(
        action="propose_day", date="2026-08-12", add_tracks=3, free=3
    )
    # C: next workday after 14 is 17
    assert any(o.date == "2026-08-17" and o.action == "propose_day" for o in result.options)


def test_calculate_capacity_deficit_skips_completed_and_zero_max() -> None:
    fill_targets = [{"date": "2026-08-14", "tracks": 5}]
    day_capacity = {"2026-08-12": 0, "2026-08-13": 5, "2026-08-14": 5}
    result = calculate_capacity_deficit(
        8 * MAX_TRACK_LENGTH_M,
        fill_targets,
        day_capacity,
        occupancy={},
        completed_dates={"2026-08-13"},
        today="2026-08-12",
        is_workday=lambda _d: True,
    )
    assert result is not None
    dates = [o.date for o in result.options]
    assert "2026-08-12" not in dates  # day_max 0
    assert "2026-08-13" not in dates  # completed / SGP


def test_calculate_capacity_deficit_max_ten_options() -> None:
    fill_targets = [{"date": "2026-09-01", "tracks": 5}]
    day_capacity: dict[str, int] = {}
    result = calculate_capacity_deficit(
        20 * MAX_TRACK_LENGTH_M,
        fill_targets,
        day_capacity,
        today="2026-08-01",
        is_workday=lambda _d: True,
        max_options=10,
    )
    assert result is not None
    assert len(result.options) == 10


def test_capacity_module_has_no_app_imports() -> None:
    path = Path(__file__).resolve().parents[1] / "core" / "production" / "capacity.py"
    tree = ast.parse(path.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                assert not alias.name.startswith("app"), alias.name
        elif isinstance(node, ast.ImportFrom):
            module = node.module or ""
            assert not module.startswith("app"), module
