"""Тесты backfill identity у plate_assignments.

Phase 1 P8: гарантия, что после ``backfill_assignment_identity`` каждая
запись ``plate_assignments`` имеет ``kp_id``+``plate_name`` для случаев,
когда оптимизатор пометил их как ``slot_exhausted`` /
``secondary_unmapped`` (proportional slots исчерпаны).
"""
from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.plate_attribution import (  # noqa: E402
    backfill_assignment_identity,
    backfill_track_items_identity,
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


class TestBackfillSlotExhausted:
    def test_assigns_identity_when_orders_match_by_canonical_key(self):
        orders = [_order(101, "Плиты ПБ 60-12-8п", 6.0, 1200, 8, qty=2)]
        assignments = [
            _assignment(6.0, 1200, 8, kp_id=101, plate_name="Плиты ПБ 60-12-8п"),
            _assignment(6.0, 1200, 8, kp_id=None, plate_name=None),
        ]

        backfilled = backfill_assignment_identity(assignments, orders)

        assert backfilled == 1
        assert assignments[1]["kp_id"] == 101
        assert assignments[1]["plate_name"] == "Плиты ПБ 60-12-8п"
        assert assignments[1]["identity_match_type"] == "backfilled"

    def test_picks_order_with_largest_remaining_qty(self):
        orders = [
            _order(201, "Плиты ПБ 50-12-8п", 5.0, 1200, 8, qty=3),
            _order(202, "Плиты ПБ 50-12-8п", 5.0, 1200, 8, qty=10),
        ]
        assignments = [
            _assignment(5.0, 1200, 8, kp_id=201, plate_name="Плиты ПБ 50-12-8п"),
            _assignment(5.0, 1200, 8, kp_id=201, plate_name="Плиты ПБ 50-12-8п"),
            _assignment(5.0, 1200, 8, kp_id=201, plate_name="Плиты ПБ 50-12-8п"),
            _assignment(5.0, 1200, 8, kp_id=None, plate_name=None),
        ]

        backfilled = backfill_assignment_identity(assignments, orders)

        assert backfilled == 1
        assert assignments[3]["kp_id"] == 202

    def test_no_op_when_all_have_identity(self):
        orders = [_order(301, "X", 6.0, 1200, 8, qty=1)]
        assignments = [_assignment(6.0, 1200, 8, kp_id=301, plate_name="X")]

        backfilled = backfill_assignment_identity(assignments, orders)

        assert backfilled == 0
        assert assignments[0].get("identity_match_type") != "backfilled"

    def test_idempotent_on_second_call(self):
        orders = [_order(401, "Y", 4.5, 700, 6, qty=2)]
        assignments = [
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=None, plate_name=None),
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=None, plate_name=None),
        ]

        first = backfill_assignment_identity(assignments, orders)
        second = backfill_assignment_identity(assignments, orders)

        assert first == 2
        assert second == 0
        assert all(a["kp_id"] == 401 for a in assignments)
        assert all(a["plate_name"] == "Y" for a in assignments)

    def test_keeps_unchanged_when_no_matching_order(self):
        orders = [_order(501, "A", 6.0, 1200, 8, qty=1)]
        assignments = [_assignment(5.0, 320, 8, kp_id=None, plate_name=None)]

        backfilled = backfill_assignment_identity(assignments, orders)

        assert backfilled == 0
        assert assignments[0]["kp_id"] is None
        assert assignments[0]["plate_name"] is None

    def test_load_code_normalization_800_vs_8(self):
        orders = [_order(601, "Z", 6.0, 1200, 8, qty=1)]
        assignments = [_assignment(6.0, 1200, load_code=800,
                                   kp_id=None, plate_name=None)]

        backfilled = backfill_assignment_identity(assignments, orders)

        assert backfilled == 1
        assert assignments[0]["kp_id"] == 601


class TestBackfillRealScenario:
    """Сценарий из ошибки пользователя: узкие secondary плиты с
    ``identity_match_type='secondary_unmapped'`` после исчерпания слотов."""

    def test_screenshot_secondaries_get_identity(self):
        orders = [
            _order(700, "Плиты ПБ 45-7,0-6п", 4.5, 700, 6, qty=4),
            _order(700, "Плиты ПБ 42-5,3-6п", 4.2, 530, 6, qty=4),
            _order(700, "Плиты ПБ 54,3-5,3-8п", 5.43, 530, 8, qty=12),
        ]
        assignments = [
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=None, plate_name=None),
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=None, plate_name=None),
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=None, plate_name=None),
            _assignment(4.5, 700, 6, source="secondary",
                        kp_id=None, plate_name=None),
            _assignment(4.2, 530, 6, source="secondary",
                        kp_id=None, plate_name=None),
            _assignment(5.43, 530, 8, source="secondary",
                        kp_id=None, plate_name=None),
        ]

        backfilled = backfill_assignment_identity(assignments, orders)

        assert backfilled == 6
        assert all(a.get("kp_id") == 700 for a in assignments)
        assert assignments[0]["plate_name"] == "Плиты ПБ 45-7,0-6п"
        assert assignments[4]["plate_name"] == "Плиты ПБ 42-5,3-6п"
        assert assignments[5]["plate_name"] == "Плиты ПБ 54,3-5,3-8п"


# =============================================================================
# Tests for backfill_track_items_identity
# =============================================================================


def _track(items: list[dict], label: str | None = None) -> dict:
    track: dict = {"items": items}
    if label is not None:
        track["label"] = label
    return track


def _solid_item(length: float, width_m: float, load_code: int = 8,
                kp_id: int | None = None, plate_name: str | None = None) -> dict:
    return {
        "length": length,
        "mode": "solid",
        "width": width_m,
        "load_code": load_code,
        "kp_id": kp_id,
        "plate_name": plate_name,
    }


def _split_item(length: float, main_w: float, rest_w: float, load_code: int = 8,
                kp_id: int | None = None, plate_name: str | None = None,
                secondary_cuts: list | None = None) -> dict:
    return {
        "length": length,
        "mode": "split",
        "main_w": main_w,
        "rest_w": rest_w,
        "load_code": load_code,
        "kp_id": kp_id,
        "plate_name": plate_name,
        "secondary_cuts": secondary_cuts or [],
    }


class TestBackfillTrackItemsRoot:
    def test_backfills_root_solid_item(self):
        orders = [_order(101, "Плиты ПБ 60-12-8п", 6.0, 1200, 8, qty=1)]
        tracks = [_track([_solid_item(6.0, 1.2, 8)])]

        backfilled = backfill_track_items_identity(tracks, orders)

        assert backfilled == 1
        item = tracks[0]["items"][0]
        assert item["kp_id"] == 101
        assert item["plate_name"] == "Плиты ПБ 60-12-8п"
        assert item["identity_match_type"] == "backfilled"

    def test_backfills_split_root_using_main_w(self):
        orders = [_order(202, "Плиты ПБ 60-12-8п", 6.0, 1200, 8, qty=1)]
        item = _split_item(6.0, main_w=1.2, rest_w=0.32)

        backfilled = backfill_track_items_identity([_track([item])], orders)

        assert backfilled == 1
        assert item["kp_id"] == 202

    def test_skip_already_attributed_root(self):
        orders = [_order(303, "Плиты ПБ 60-12-8п", 6.0, 1200, 8, qty=2)]
        tracks = [
            _track([
                _solid_item(6.0, 1.2, 8, kp_id=303, plate_name="Плиты ПБ 60-12-8п"),
                _solid_item(6.0, 1.2, 8),
            ])
        ]

        backfilled = backfill_track_items_identity(tracks, orders)

        assert backfilled == 1
        assert tracks[0]["items"][0].get("identity_match_type") != "backfilled"
        assert tracks[0]["items"][1]["kp_id"] == 303

    def test_picks_order_with_max_remaining_when_multiple_match(self):
        orders = [
            _order(401, "Плиты ПБ 50-12-8п", 5.0, 1200, 8, qty=2),
            _order(402, "Плиты ПБ 50-12-8п", 5.0, 1200, 8, qty=10),
        ]
        tracks = [
            _track([
                _solid_item(5.0, 1.2, 8, kp_id=401, plate_name="Плиты ПБ 50-12-8п"),
                _solid_item(5.0, 1.2, 8),
            ])
        ]

        backfilled = backfill_track_items_identity(tracks, orders)

        assert backfilled == 1
        assert tracks[0]["items"][1]["kp_id"] == 402


class TestBackfillTrackItemsSecondary:
    def test_backfills_secondary_via_target_order_key(self):
        orders = [
            _order(500, "Плиты ПБ 32-3,2-8п", 3.2, 320, 8, qty=2),
        ]
        sec = {
            "width": 0.32,
            "label": "[2] ...",
            "transverse_cut": True,
            "target_length": 3.2,
            "load_code": 8,
            "target_order_key": (3.2, 320, 8),
        }
        item = _split_item(
            6.0, main_w=1.2, rest_w=0.32, load_code=8,
            kp_id=500, plate_name="Плиты ПБ 60-12-8п",
            secondary_cuts=[sec],
        )

        backfilled = backfill_track_items_identity(
            [_track([item])], orders
        )

        assert backfilled == 1
        assert sec["kp_id"] == 500
        assert sec["plate_name"] == "Плиты ПБ 32-3,2-8п"

    def test_backfills_secondary_via_target_length_when_no_token(self):
        orders = [
            _order(600, "Плиты ПБ 30-3,2-8п", 3.0, 320, 8, qty=1),
        ]
        sec = {
            "width": 0.32,
            "label": "О ...",
            "has_transverse": True,
            "target_length": 3.0,
            "load_code": 8,
        }
        item = _split_item(
            6.0, main_w=1.2, rest_w=0.32, load_code=8,
            kp_id=600, plate_name="Плиты ПБ 60-12-8п",
            secondary_cuts=[sec],
        )

        backfilled = backfill_track_items_identity(
            [_track([item])], orders
        )

        assert backfilled == 1
        assert sec["kp_id"] == 600
        assert sec["plate_name"] == "Плиты ПБ 30-3,2-8п"

    def test_backfills_narrowing_secondary_via_parent_length(self):
        # narrowing: target_length не выставлен, target_order_key тоже None
        # (например, restored из старого плана). Длина = parent.length.
        orders = [
            _order(700, "Плиты ПБ 60-3,2-8п", 6.0, 320, 8, qty=1),
        ]
        sec = {
            "width": 0.32,
            "label": "[2] ...",
            "load_code": 8,
        }
        item = _split_item(
            6.0, main_w=1.2, rest_w=0.32, load_code=8,
            kp_id=700, plate_name="Плиты ПБ 60-12-8п",
            secondary_cuts=[sec],
        )

        backfilled = backfill_track_items_identity(
            [_track([item])], orders
        )

        assert backfilled == 1
        assert sec["kp_id"] == 700

    def test_secondary_inherits_load_code_from_parent_when_missing(self):
        orders = [_order(800, "Плиты ПБ 30-3,2-8п", 3.0, 320, 8, qty=1)]
        sec = {
            "width": 0.32,
            "target_length": 3.0,
        }
        item = _split_item(
            6.0, main_w=1.2, rest_w=0.32, load_code=8,
            kp_id=800, plate_name="Плиты ПБ 60-12-8п",
            secondary_cuts=[sec],
        )

        backfilled = backfill_track_items_identity(
            [_track([item])], orders
        )

        assert backfilled == 1
        assert sec["kp_id"] == 800
        assert sec["load_code"] == 8

    def test_load_class_1250_via_canonical_key(self):
        # 12.5п (load_class 1250) корректно нормализуется через canonical_plate_key.
        orders = [_order(900, "Плиты ПБ 30-3,2-12,5п", 3.0, 320, 12.5, qty=1)]
        sec = {
            "width": 0.32,
            "target_length": 3.0,
            "load_code": 12.5,
        }
        item = _split_item(
            6.0, main_w=1.2, rest_w=0.32, load_code=12.5,
            kp_id=900, plate_name="Плиты ПБ 60-12-12,5п",
            secondary_cuts=[sec],
        )

        backfilled = backfill_track_items_identity(
            [_track([item])], orders
        )

        assert backfilled == 1
        assert sec["kp_id"] == 900


class TestBackfillTrackItemsSharedScenarios:
    def test_idempotent_on_second_call(self):
        orders = [_order(1000, "Плиты ПБ 30-3,2-8п", 3.0, 320, 8, qty=2)]
        sec1 = {"width": 0.32, "target_length": 3.0, "load_code": 8}
        sec2 = {"width": 0.32, "target_length": 3.0, "load_code": 8}
        tracks = [
            _track([
                _split_item(6.0, 1.2, 0.32, 8, kp_id=1000,
                            plate_name="Плиты ПБ 60-12-8п",
                            secondary_cuts=[sec1, sec2]),
            ])
        ]

        first = backfill_track_items_identity(tracks, orders)
        second = backfill_track_items_identity(tracks, orders)

        assert first == 2
        assert second == 0
        assert sec1["kp_id"] == 1000
        assert sec2["kp_id"] == 1000

    def test_rescue_track_items_skipped_but_pre_counted(self):
        # РЕСКЬЮ items уже имеют identity и НЕ должны пере-аттрибутироваться.
        orders = [_order(1100, "X", 6.0, 1200, 8, qty=2)]
        rescue_track = _track(
            [_solid_item(6.0, 1.2, 8, kp_id=1100, plate_name="X")],
            label="РЕСКЬЮ",
        )

        backfilled = backfill_track_items_identity([rescue_track], orders)

        assert backfilled == 0
        assert rescue_track["items"][0]["kp_id"] == 1100

    def test_consumed_counter_prevents_over_attribution_within_items(self):
        # Если у заказа qty=1, а в треках 2 НЕаттрибутированных item с тем же
        # ключом — оба матчатся в один заказ (нет других кандидатов), но
        # consumed правильно отражает оба.
        orders = [_order(1200, "Плиты ПБ 60-12-8п", 6.0, 1200, 8, qty=1)]
        tracks = [
            _track([
                _solid_item(6.0, 1.2, 8),
                _solid_item(6.0, 1.2, 8),
            ])
        ]

        backfilled = backfill_track_items_identity(tracks, orders)

        assert backfilled == 2
        assert tracks[0]["items"][0]["kp_id"] == 1200
        assert tracks[0]["items"][1]["kp_id"] == 1200

    def test_no_op_when_orders_empty(self):
        tracks = [_track([_solid_item(6.0, 1.2, 8)])]
        assert backfill_track_items_identity(tracks, []) == 0
        assert tracks[0]["items"][0]["kp_id"] is None

    def test_secondary_without_any_length_signal_skipped(self):
        # Если target_order_key=None, target_length=None, parent.length=None —
        # ключ не построить, item остаётся без identity.
        orders = [_order(1300, "Плиты ПБ 60-3,2-8п", 6.0, 320, 8, qty=1)]
        sec = {"width": 0.32, "load_code": 8}
        item = {
            "length": None,  # parent тоже без length
            "mode": "split",
            "main_w": 1.2,
            "rest_w": 0.32,
            "load_code": 8,
            "kp_id": None,
            "plate_name": None,
            "secondary_cuts": [sec],
        }

        backfilled = backfill_track_items_identity([_track([item])], orders)

        assert backfilled == 0
        assert sec.get("kp_id") is None
