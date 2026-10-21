"""Driver-day pool: groups, rotation by substitute count, weekend code."""

from __future__ import annotations

from core.gsm.driver_pool import DriverCandidate, order_candidates, pick_driver

PRIMARY = 143
POLYAKOVA = 1
SKRYABIN = 29
LONSHAKOVA = 470


def _candidate(
    driver_id: int,
    code: str,
    personnel: str,
    *,
    has_own_vehicle: bool = False,
) -> DriverCandidate:
    return DriverCandidate(
        driver_id=driver_id,
        code=code,
        personnel_number=personnel,
        has_own_vehicle=has_own_vehicle,
    )


def test_free_primary_is_chosen_without_warning() -> None:
    ordered = order_candidates(
        PRIMARY,
        (
            _candidate(POLYAKOVA, "Я", "1"),
            _candidate(PRIMARY, "Я", "143", has_own_vehicle=True),
            _candidate(SKRYABIN, "Я", "29"),
        ),
        {},
    )
    chosen, warnings = pick_driver(ordered)
    assert chosen == PRIMARY
    assert warnings == ()


def test_absent_primary_yields_head_of_the_no_vehicle_group() -> None:
    ordered = order_candidates(
        PRIMARY,
        (
            _candidate(PRIMARY, "ОТ", "143", has_own_vehicle=True),
            _candidate(SKRYABIN, "Я", "29"),
            _candidate(POLYAKOVA, "Я", "1"),
            _candidate(LONSHAKOVA, "Я", "470", has_own_vehicle=True),
        ),
        {},
    )
    ids = [item.driver_id for item in ordered]
    assert PRIMARY not in ids
    assert ids[0] == POLYAKOVA
    assert LONSHAKOVA in ids
    assert ids.index(POLYAKOVA) < ids.index(LONSHAKOVA)


def test_rotation_prefers_lower_personnel_then_the_other_after_increment() -> None:
    pool = (
        _candidate(POLYAKOVA, "Я", "1"),
        _candidate(SKRYABIN, "Я", "29"),
    )
    first = order_candidates(PRIMARY, pool, {POLYAKOVA: 0, SKRYABIN: 0})
    assert first[0].driver_id == POLYAKOVA

    second = order_candidates(PRIMARY, pool, {POLYAKOVA: 1, SKRYABIN: 0})
    assert second[0].driver_id == SKRYABIN

    tied = order_candidates(PRIMARY, pool, {POLYAKOVA: 1, SKRYABIN: 1})
    assert tied[0].driver_id == POLYAKOVA


def test_vacation_of_one_substitute_is_a_series_for_the_other() -> None:
    only_polyakova = (_candidate(POLYAKOVA, "Я", "1"), _candidate(SKRYABIN, "ОТ", "29"))
    ordered = order_candidates(PRIMARY, only_polyakova, {POLYAKOVA: 4, SKRYABIN: 0})
    assert [item.driver_id for item in ordered] == [POLYAKOVA]


def test_weekend_code_on_the_chosen_driver_is_a_soft_warning() -> None:
    ordered = order_candidates(PRIMARY, (_candidate(PRIMARY, "В", "143"),), {})
    chosen, warnings = pick_driver(ordered)
    assert chosen == PRIMARY
    assert warnings == ("weekend_driver",)


def test_empty_pool_needs_a_driver() -> None:
    chosen, warnings = pick_driver(())
    assert chosen is None
    assert warnings == ("no_available_driver",)
