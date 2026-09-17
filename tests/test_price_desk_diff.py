"""PD-004: diff current prices vs file (changed/new/missing/unchanged)."""

from __future__ import annotations

from core.price_desk_diff import diff_price_maps, prices_to_map


def test_diff_four_buckets() -> None:
    current = {
        (17, 6): 100.0,
        (17, 8): 200.0,
        (18, 10): 300.0,
        (19, 6): 400.0,
    }
    incoming = {
        (17, 6): 100.0,  # unchanged
        (17, 8): 250.0,  # changed
        (20, 6): 500.0,  # new
        # (18, 10) missing from file
        # (19, 6) missing from file
    }
    result = diff_price_maps(current, incoming)
    assert result.unchanged == 1
    assert result.changed == 1
    assert result.new == 1
    assert result.missing == 2
    assert result.examples.changed[0].key == "17-8"
    assert result.examples.changed[0].old_price == 200.0
    assert result.examples.changed[0].new_price == 250.0
    missing_keys = {item.key for item in result.examples.missing}
    assert missing_keys == {"18-10", "19-6"}


def test_diff_threshold_005() -> None:
    current = {("ФБС 9.3.6-Т", "B25"): 10.0}
    same = diff_price_maps(current, {("ФБС 9.3.6-Т", "B25"): 10.004})
    assert same.unchanged == 1
    assert same.changed == 0
    changed = diff_price_maps(current, {("ФБС 9.3.6-Т", "B25"): 10.006})
    assert changed.changed == 1
    assert changed.unchanged == 0


def test_prices_to_map_step_keys() -> None:
    mapping = prices_to_map([("ЛС11", 1409.9, "Лестничные ступени ЛС11")], key_len=1)
    assert mapping[("ЛС11",)] == 1409.9


def test_empty_maps() -> None:
    result = diff_price_maps({}, {(17, 6): 1.0})
    assert result.new == 1
    assert result.changed == 0
    assert result.missing == 0
    assert result.unchanged == 0
