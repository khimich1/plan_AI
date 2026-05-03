"""Тесты Phase 2-4 для core.rescue_tracks.build_rescue_tracks.

Инварианты:
1. Кол-во ``items`` в rescue_tracks == кол-во записей в rescue_assignments.
2. У каждой записи ``rescue_assignments[i]`` ``source='rescue'`` и identity
   совпадает с парным ``rescue_tracks[*].items[*]``.
3. Phase 4: если оптимизатор уже произвёл секондари (``plate_assignments``
   полный) — ``missing_counts == {}``, фантомные RESCUE не создаются.
   Это ключевая регрессия из ошибки пользователя (узкие 530/665/700 мм).
"""
from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.rescue_tracks import (  # noqa: E402
    build_rescue_tracks,
    build_track_gap_rescue_tracks,
)


def _order(kp_id: int, plate_name: str, length: float, width_mm: int,
           load_code: int = 8, qty: int = 1) -> dict:
    return {
        "kp_id": kp_id,
        "plate_name": plate_name,
        "length": length,
        "width": width_mm,
        "load_code": load_code,
        "qty": qty,
    }


def _assignment(length: float, width_mm: int, load_code: int = 8,
                source: str = "primary", kp_id: int | None = None,
                plate_name: str | None = None) -> dict:
    return {
        "length": length,
        "width": width_mm,
        "load_code": load_code,
        "source": source,
        "kp_id": kp_id,
        "plate_name": plate_name,
    }


class TestRescueAssignmentsInvariant:
    def test_returns_three_tuple(self):
        result = build_rescue_tracks([], [])
        assert isinstance(result, tuple)
        assert len(result) == 3

    def test_empty_when_demand_fully_covered(self):
        orders = [_order(1, "X", 6.0, 1200, 8, qty=1)]
        plate_assignments = [
            _assignment(6.0, 1200, 8, kp_id=1, plate_name="X"),
        ]

        rescue_tracks, missing, rescue_pa = build_rescue_tracks(
            orders, plate_assignments
        )

        assert missing == {}
        assert rescue_tracks == []
        assert rescue_pa == []

    def test_one_to_one_items_and_assignments(self):
        orders = [_order(1, "X", 6.0, 1200, 8, qty=3)]
        plate_assignments: list = []

        rescue_tracks, missing, rescue_pa = build_rescue_tracks(
            orders, plate_assignments
        )

        items_total = sum(len(t.get("items") or []) for t in rescue_tracks)
        assert items_total == 3
        assert len(rescue_pa) == 3
        assert all(a["source"] == "rescue" for a in rescue_pa)
        assert all(a["kp_id"] == 1 for a in rescue_pa)
        assert all(a["plate_name"] == "X" for a in rescue_pa)

    def test_assignment_keeps_load_code_normalized(self):
        orders = [_order(2, "Y", 4.5, 700, 6, qty=2)]
        plate_assignments: list = []

        _, _, rescue_pa = build_rescue_tracks(orders, plate_assignments)

        assert len(rescue_pa) == 2
        for a in rescue_pa:
            assert a["load_code"] == 6
            assert a["width"] == 700
            assert a["length"] == 4.5


class TestPhase4NoPhantoms:
    """Phase 4: главный инвариант — если оптимизатор УЖЕ произвёл плиты с
    identity, RESCUE не должен создавать фантомных копий."""

    def test_secondary_already_in_plate_assignments_no_phantom(self):
        """Сценарий из ошибки пользователя: 4 плиты 45-7,0-6п уже в
        plate_assignments как source='secondary', RESCUE НЕ должен их
        дублировать."""
        orders = [_order(700, "Плиты ПБ 45-7,0-6п", 4.5, 700, 6, qty=4)]
        plate_assignments = [
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=700, plate_name="Плиты ПБ 45-7,0-6п"),
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=700, plate_name="Плиты ПБ 45-7,0-6п"),
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=700, plate_name="Плиты ПБ 45-7,0-6п"),
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=700, plate_name="Плиты ПБ 45-7,0-6п"),
        ]

        rescue_tracks, missing, rescue_pa = build_rescue_tracks(
            orders, plate_assignments
        )

        assert missing == {}, "Фантомные RESCUE для покрытых секондари недопустимы"
        assert rescue_tracks == []
        assert rescue_pa == []

    def test_partial_coverage_creates_only_deficit(self):
        """Если 2 из 4 секондари в plate_assignments — RESCUE создаёт ровно 2."""
        orders = [_order(700, "Y", 4.5, 700, 6, qty=4)]
        plate_assignments = [
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=700, plate_name="Y"),
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=700, plate_name="Y"),
        ]

        rescue_tracks, missing, rescue_pa = build_rescue_tracks(
            orders, plate_assignments
        )

        assert missing == {(700, "Y"): 2}
        assert len(rescue_pa) == 2
        assert all(a["plate_name"] == "Y" for a in rescue_pa)

    def test_screenshot_scenario_no_phantoms(self):
        """13 узких плит из ошибки пользователя — все УЖЕ в plate_assignments
        как secondary с identity. После Phase 4 missing должен быть пуст."""
        screenshot = [
            (700, "Плиты ПБ 26,7-6,65-8п", 2.67, 665, 8, 1),
            (700, "Плиты ПБ 25,4-3,0-8п", 2.54, 300, 8, 2),
            (700, "Плиты ПБ 42-5,3-6п", 4.2, 530, 6, 4),
            (700, "Плиты ПБ 45-7,0-6п", 4.5, 700, 6, 4),
            (700, "Плиты ПБ 42-3,0-8п", 4.2, 300, 8, 1),
            (700, "Плиты ПБ 42,6-5,3-10п", 4.26, 530, 10, 1),
            (700, "Плиты ПБ 49,9-6,65-8п", 4.99, 665, 8, 3),
            (700, "Плиты ПБ 51-3,2-8п", 5.1, 320, 8, 3),
            (700, "Плиты ПБ 54,3-5,3-8п", 5.43, 530, 8, 12),
            (700, "Плиты ПБ 60-5,3-8п", 6.0, 530, 8, 4),
            (700, "Плиты ПБ 60-6,65-8п", 6.0, 665, 8, 4),
            (700, "Плиты ПБ 61,8-5,0-8п", 6.18, 500, 8, 1),
            (700, "Плиты ПБ 63,9-5,3-10п", 6.39, 530, 10, 1),
        ]
        orders = [_order(*row) for row in screenshot]
        plate_assignments = []
        for kp_id, name, length, width_mm, lc, qty in screenshot:
            for _ in range(qty):
                plate_assignments.append(_assignment(
                    length, width_mm, lc, source="secondary",
                    kp_id=kp_id, plate_name=name,
                ))

        rescue_tracks, missing, rescue_pa = build_rescue_tracks(
            orders, plate_assignments
        )

        assert missing == {}
        assert rescue_tracks == []
        assert rescue_pa == []

    def test_load_code_normalization_800_in_assignment_vs_8_in_order(self):
        """Регрессия: assignment с load_code=800 (raw из БД) и order с
        load_code=8 — должны матчиться через identity (kp_id, plate_name)."""
        orders = [_order(1, "Z", 6.0, 1200, 8, qty=1)]
        plate_assignments = [
            _assignment(6.0, 1200, 800, source="primary",
                        kp_id=1, plate_name="Z"),
        ]

        _, missing, _ = build_rescue_tracks(orders, plate_assignments)

        assert missing == {}, "Identity match не должен зависеть от raw load_code"


class TestRescueWithMissingDemand:
    """Когда орган не покрыт оптимизатором — RESCUE покрывает дефицит."""

    def test_creates_rescue_for_uncovered_orders(self):
        orders = [
            _order(1, "A", 6.0, 1200, 8, qty=2),
            _order(2, "B", 4.5, 700, 6, qty=3),
        ]
        plate_assignments = [
            _assignment(6.0, 1200, 8, source="primary", kp_id=1, plate_name="A"),
        ]

        rescue_tracks, missing, rescue_pa = build_rescue_tracks(
            orders, plate_assignments
        )

        assert missing == {(1, "A"): 1, (2, "B"): 3}
        assert len(rescue_pa) == 4
        a_count = sum(1 for a in rescue_pa if a["plate_name"] == "A")
        b_count = sum(1 for a in rescue_pa if a["plate_name"] == "B")
        assert a_count == 1
        assert b_count == 3


class TestTrackGapRescue:
    def test_creates_tracks_without_duplicate_assignments_for_track_gap(self):
        orders = [
            _order(1, "Плиты ПБ 27-7,2-8п", 2.7, 720, 8, qty=3),
            _order(1, "Плиты ПБ 73-3,2-10п", 7.3, 320, 10, qty=1),
        ]
        tracks = [
            {
                "items": [
                    {
                        "length": 2.7,
                        "mode": "solid",
                        "width": 0.72,
                        "load_code": 8,
                        "kp_id": 1,
                        "plate_name": "Плиты ПБ 27-7,2-8п",
                    }
                ]
            }
        ]

        rescue_tracks, missing = build_track_gap_rescue_tracks(orders, tracks)

        assert missing == {
            (1, "Плиты ПБ 27-7,2-8п"): 2,
            (1, "Плиты ПБ 73-3,2-10п"): 1,
        }
        items_total = sum(len(t.get("items") or []) for t in rescue_tracks)
        assert items_total == 3
        assert all(track.get("label") == "РЕСКЬЮ" for track in rescue_tracks)
