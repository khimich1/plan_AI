"""NG-203 / NG-204: строка замены марки и окно «было / станет»."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.config_and_data import approximate_weight_kg
from core.guid_replace import (
    GuidReplaceError,
    ReplaceTarget,
    build_replacement,
    preview_replacement,
)
from core.kp_db_schema import init_schema
from core.kp_persistence_service import KpPersistenceService
from core.pile_price_db import init_pile_prices_schema
from core.price_db import init_schema as init_price_schema
from core.step_price_db import init_step_prices_schema


def _pb(tmp_path: Path) -> str:
    return str(tmp_path / "pb.db")


def _seed_plate_price(pb: str, length_dm: int, load_code: int, price: float) -> None:
    init_price_schema(pb)
    with sqlite3.connect(pb) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO prices (length_dm, load_code, price) VALUES (?, ?, ?)",
            (length_dm, load_code, price),
        )


def _seed_pile_price(pb: str, mark: str, grade: str, price: float) -> None:
    init_pile_prices_schema(pb)
    with sqlite3.connect(pb) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO pile_prices (mark, concrete_grade, price) VALUES (?, ?, ?)",
            (mark, grade, price),
        )


def _seed_step_price(pb: str, mark: str, price: float) -> None:
    init_step_prices_schema(pb)
    with sqlite3.connect(pb) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO step_prices (mark, price) VALUES (?, ?)",
            (mark, price),
        )


def _plate(**overrides: object) -> dict:
    base = {
        "line_id": "ln",
        "product_type": "plates",
        "name": "Плиты ПБ 51-12-12п",
        "length_m": 5.1,
        "width_m": 1.2,
        "load_class": 1200,
        "qty": 3,
        "unit_price": 1000.0,
        "discounted_price": 1000.0,
        "weight": 100.0,
        "length_dm_raw": "51",
        "status": None,
        "plan_id": "plan-1",
        "concrete_grade": "М500",
    }
    base.update(overrides)
    return base


def _pile(**overrides: object) -> dict:
    base = {
        "line_id": "ln_p",
        "product_type": "piles",
        "product_kind": "pile",
        "name": "C14-40T4",
        "mark": "C14-40T4",
        "concrete_grade": "B25",
        "qty": 2,
        "unit_price": 100.0,
        "status": "в архиве",
    }
    base.update(overrides)
    return base


def test_plate_replacement_writes_chosen_card_and_load_1250(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_plate_price(pb, 51, 12, 4000.0)
    new = build_replacement(
        _plate(),
        ReplaceTarget(
            "Плиты ПБ 51-12-12,5п",
            "0de0f4a7-dc8e-11ea-8f69-04d9f5387845",
        ),
        10,
        "plate",
        actor="manager",
        pb_db_path=pb,
    )
    assert new["name"] == "Плиты ПБ 51-12-12,5п"
    assert new["load_class"] == 1250
    assert new["length_m"] == pytest.approx(5.1)
    assert new["width_m"] == pytest.approx(1.2)
    assert new["length_dm_raw"] == "51"
    assert new["nomenclature_id"] == "0de0f4a7-dc8e-11ea-8f69-04d9f5387845"
    assert new["unit_price"] == pytest.approx(4000.0)
    assert new["discounted_price"] == pytest.approx(3600.0)
    assert new["qty"] == 3
    assert new["line_id"] == "ln"
    assert new["plan_id"] == "plan-1"
    assert new["concrete_grade"] == "М500"
    assert new["weight"] == pytest.approx(approximate_weight_kg(5.1, 1.2) * 3)


def test_plate_width_change_rewrites_size_weight_and_price(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_plate_price(pb, 29, 8, 2500.0)
    new = build_replacement(
        _plate(
            name="Плиты ПБ 29-5,25-8п",
            length_m=2.9,
            width_m=0.525,
            load_class=800,
            length_dm_raw="29",
            qty=4,
        ),
        ReplaceTarget("Плиты ПБ 29-5,3-8п", "guid-53"),
        0,
        "plate",
        actor="manager",
        pb_db_path=pb,
    )
    assert new["name"] == "Плиты ПБ 29-5,3-8п"
    assert new["width_m"] == pytest.approx(0.53)
    assert new["length_m"] == pytest.approx(2.9)
    assert new["load_class"] == 800
    assert new["nomenclature_id"] == "guid-53"
    assert new["unit_price"] == pytest.approx(2500.0)
    assert new["qty"] == 4
    assert new["weight"] == pytest.approx(approximate_weight_kg(2.9, 0.53) * 4)


def test_replacement_does_not_parse_load_or_like(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pb = _pb(tmp_path)
    _seed_plate_price(pb, 51, 12, 4000.0)

    def boom(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("разбор не должен вызывать это")

    monkeypatch.setattr("core.config_and_data.parse_load_code_from_name", boom)
    monkeypatch.setattr("core.kp_db_nomenclature.find_plate_card", boom)
    new = build_replacement(
        _plate(),
        ReplaceTarget("Плиты ПБ 51-12-12,5п", "guid"),
        0,
        "plate",
        actor="economist",
        pb_db_path=pb,
    )
    assert new["load_class"] == 1250
    assert new["name"] == "Плиты ПБ 51-12-12,5п"


def test_discount_is_applied_once_from_list_price(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_plate_price(pb, 51, 12, 1000.0)
    new = build_replacement(
        _plate(unit_price=500.0, discounted_price=250.0),
        ReplaceTarget("Плиты ПБ 51-12-12,5п", "guid"),
        10,
        "plate",
        actor="manager",
        pb_db_path=pb,
    )
    assert new["unit_price"] == pytest.approx(1000.0)
    assert new["discounted_price"] == pytest.approx(900.0)


def test_pile_without_price_in_concrete_grade_is_refused(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_pile_price(pb, "С14-40T4", "B25", 5000.0)
    with pytest.raises(GuidReplaceError, match="классе бетона"):
        build_replacement(
            _pile(concrete_grade="B30", mark="C14-40T4", name="C14-40T4"),
            ReplaceTarget("С14-40T4"),
            0,
            "pile",
            actor="economist",
            pb_db_path=pb,
        )


def test_pile_u_spelling_is_kept_and_grade_unchanged(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_pile_price(pb, "С110.30-9у", "B25", 8000.0)
    new = build_replacement(
        _pile(mark="С110.30-9", name="С110.30-9", concrete_grade="B25", qty=5),
        ReplaceTarget("С110.30-9у"),
        20,
        "pile",
        actor="manager",
        pb_db_path=pb,
    )
    assert new["mark"] == "С110.30-9у"
    assert new["name"] == "С110.30-9у"
    assert new["concrete_grade"] == "B25"
    assert new["qty"] == 5
    assert new["line_id"] == "ln_p"
    assert new["unit_price"] == pytest.approx(8000.0)
    assert new["discounted_price"] == pytest.approx(6400.0)


def test_step_price_ignores_concrete_grade(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_step_price(pb, "ЛС11", 1500.0)
    line = {
        "line_id": "ln_s",
        "product_type": "steps",
        "product_kind": "step",
        "name": "ЛС11-1",
        "mark": "ЛС11-1",
        "concrete_grade": "B30",
        "qty": 1,
        "unit_price": 10.0,
    }
    new = build_replacement(
        line,
        ReplaceTarget("ЛС11"),
        0,
        "stair_step",
        actor="manager",
        pb_db_path=pb,
    )
    assert new["mark"] == "ЛС11"
    assert new["concrete_grade"] == "B30"
    assert new["unit_price"] == pytest.approx(1500.0)


def test_oneoff_manager_refused_economist_needs_flag(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_pile_price(pb, "С14-40T4", "B25", 5000.0)
    line = _pile(price_source="oneoff")
    with pytest.raises(GuidReplaceError, match="менеджер"):
        build_replacement(
            line,
            ReplaceTarget("С14-40T4"),
            0,
            "pile",
            actor="manager",
            pb_db_path=pb,
        )
    with pytest.raises(GuidReplaceError, match="(?i)договорн"):
        build_replacement(
            line,
            ReplaceTarget("С14-40T4"),
            0,
            "pile",
            actor="economist",
            pb_db_path=pb,
        )
    new = build_replacement(
        line,
        ReplaceTarget("С14-40T4"),
        0,
        "pile",
        actor="economist",
        replace_oneoff=True,
        pb_db_path=pb,
    )
    assert new["mark"] == "С14-40T4"
    assert new.get("price_source") != "oneoff"
    assert new["unit_price"] == pytest.approx(5000.0)


def test_preview_two_discounts_and_same_qty(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_pile_price(pb, "С14-40T4", "B25", 10000.0)
    totals_db = str(tmp_path / "totals.db")
    sqlite3.connect(totals_db).close()
    order = [_pile(qty=2, unit_price=100.0, discounted_price=100.0)]
    low = preview_replacement(
        order,
        source_mark="C14-40T4",
        target=ReplaceTarget("С14-40T4"),
        discount_percent=10,
        kind="pile",
        actor="economist",
        pb_db_path=pb,
        totals_db_path=totals_db,
    )
    high = preview_replacement(
        order,
        source_mark="C14-40T4",
        target=ReplaceTarget("С14-40T4"),
        discount_percent=25,
        kind="pile",
        actor="economist",
        pb_db_path=pb,
        totals_db_path=totals_db,
    )
    assert low.qty == 2
    assert high.qty == 2
    assert low.will.unit_discounted == pytest.approx(9000.0)
    assert high.will.unit_discounted == pytest.approx(7500.0)
    assert low.will.line_sum == pytest.approx(18000.0)
    assert high.will.line_sum == pytest.approx(15000.0)
    assert low.will.total != high.will.total
    assert low.delivery_was is None
    assert low.delivery_will is None


def test_preview_shows_delivery_only_when_it_changes(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_plate_price(pb, 51, 12, 4000.0)
    _seed_plate_price(pb, 29, 8, 2500.0)
    totals_db = str(tmp_path / "totals.db")
    sqlite3.connect(totals_db).close()
    same = preview_replacement(
        [_plate(name="Плиты ПБ 51-12-12п", qty=1, unit_price=4000.0)],
        source_mark="Плиты ПБ 51-12-12п",
        target=ReplaceTarget("Плиты ПБ 51-12-12,5п", "guid"),
        discount_percent=0,
        kind="plate",
        actor="manager",
        pb_db_path=pb,
        totals_db_path=totals_db,
        logistics_cost=5000,
    )
    assert same.delivery_was is None
    assert same.delivery_will is None
    changed = preview_replacement(
        [_plate(name="Плиты ПБ 51-12-12п", qty=12, unit_price=4000.0)],
        source_mark="Плиты ПБ 51-12-12п",
        target=ReplaceTarget("Плиты ПБ 29-5,3-8п", "guid-w"),
        discount_percent=0,
        kind="plate",
        actor="manager",
        pb_db_path=pb,
        totals_db_path=totals_db,
        logistics_cost=5000,
    )
    assert changed.delivery_was is not None
    assert changed.delivery_will is not None
    assert changed.delivery_was != pytest.approx(changed.delivery_will)


def test_preview_oneoff_notice_for_economist(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_pile_price(pb, "С14-40T4", "B25", 1000.0)
    totals_db = str(tmp_path / "totals.db")
    sqlite3.connect(totals_db).close()
    view = preview_replacement(
        [_pile(price_source="oneoff")],
        source_mark="C14-40T4",
        target=ReplaceTarget("С14-40T4"),
        discount_percent=0,
        kind="pile",
        actor="economist",
        pb_db_path=pb,
        totals_db_path=totals_db,
    )
    assert view.oneoff_notice
    assert "договорн" in view.oneoff_notice.casefold()
    assert "прайс" in view.oneoff_notice.casefold()


def test_preview_totals_match_saved_kp(tmp_path: Path) -> None:
    pb = _pb(tmp_path)
    _seed_plate_price(pb, 51, 12, 4000.0)
    plita = str(tmp_path / "plita.db")
    init_schema(plita)
    order = [_plate(qty=2, unit_price=1000.0, discounted_price=900.0)]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        order,
        customer_name="Окно",
        discount_percent=10,
        product_type="plates",
        status="в архиве",
        db_path=plita,
    )
    view = preview_replacement(
        order,
        source_mark="Плиты ПБ 51-12-12п",
        target=ReplaceTarget("Плиты ПБ 51-12-12,5п", "guid-save"),
        discount_percent=10,
        kind="plate",
        actor="economist",
        pb_db_path=pb,
        totals_db_path=plita,
        kp_id=kp_id,
    )
    new_order = [
        build_replacement(
            order[0],
            ReplaceTarget("Плиты ПБ 51-12-12,5п", "guid-save"),
            10,
            "plate",
            actor="economist",
            pb_db_path=pb,
        )
    ]
    KpPersistenceService.update_kp_from_order_data(
        kp_id,
        new_order,
        discount_percent=10,
        db_path=plita,
    )
    with sqlite3.connect(plita) as conn:
        conn.row_factory = sqlite3.Row
        offer = conn.execute(
            "SELECT subtotal, vat_amount, total_amount FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()
        plate = conn.execute(
            "SELECT plate_name, load_class, unit_price, discounted_price, qty, nomenclature_id "
            "FROM kp_plates WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()
    assert offer is not None and plate is not None
    assert float(offer["total_amount"]) == pytest.approx(view.will.total)
    assert float(offer["subtotal"]) == pytest.approx(view.will.subtotal)
    assert float(offer["vat_amount"]) == pytest.approx(view.will.vat_amount)
    assert plate["plate_name"] == "Плиты ПБ 51-12-12,5п"
    assert int(plate["load_class"]) == 1250
    assert float(plate["unit_price"]) == pytest.approx(view.will.unit_price)
    assert float(plate["discounted_price"]) == pytest.approx(view.will.unit_discounted)
    assert int(plate["qty"]) == view.qty
    assert plate["nomenclature_id"] == "guid-save"


_CYR_C14 = "С14-40T4"
_LAT_C14 = "C14-40T4"
_CYR_C15 = "С15-35T7"
_LAT_C15 = "C15-35T7"
_CHOSEN_GUID = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"


def _plita(tmp_path: Path) -> str:
    path = str(tmp_path / "plita.db")
    init_schema(path)
    return path


def _demand(pb: str, kind: str, mark: str, kp_ids: list[int]) -> None:
    from core.guid_demand import ensure_schema as ensure_demand_schema

    conn = sqlite3.connect(pb)
    try:
        ensure_demand_schema(conn)
        conn.execute(
            "INSERT INTO guid_demand (product_kind, mark, field, reason, hint, state, "
            "kp_ids_json, opened_at) VALUES (?, ?, 'guid_1c', 'нет GUID 1С', "
            "'заведите карточку в 1С и загрузите отчёт', 'open', ?, '2026-11-01T00:00:00Z')",
            (kind, mark, __import__("json").dumps(kp_ids)),
        )
        conn.commit()
    finally:
        conn.close()


def _states(pb: str) -> dict[str, str]:
    conn = sqlite3.connect(pb)
    try:
        return dict(conn.execute("SELECT mark, state FROM guid_demand"))
    finally:
        conn.close()


def _offer_bits(plita: str, kp_id: int) -> tuple[float, float]:
    with sqlite3.connect(plita) as conn:
        row = conn.execute(
            "SELECT discount_percent, total_amount FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()
    assert row is not None
    return float(row[0]), float(row[1])


def test_economist_replaces_mark_in_two_kps_with_one_commit(tmp_path: Path) -> None:
    """NG-206: C14 → С14 в двух архивных КП, скидки свои, старая строка закрыта, новая открыта."""
    from app.services.guid_replace_service import GuidReplaceService

    pb = _pb(tmp_path)
    plita = _plita(tmp_path)
    _seed_pile_price(pb, _CYR_C14, "B25", 10000.0)
    first = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [_pile(mark=_LAT_C14, name=_LAT_C14, qty=2, unit_price=100.0)],
        customer_name="Первый",
        discount_percent=10,
        product_type="piles",
        status="в архиве",
        db_path=plita,
    )
    second = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [_pile(line_id="ln_p2", mark=_LAT_C14, name=_LAT_C14, qty=2, unit_price=100.0)],
        customer_name="Второй",
        discount_percent=25,
        product_type="piles",
        status="в архиве",
        db_path=plita,
    )
    _demand(pb, "pile", _LAT_C14, [first, second])

    result = GuidReplaceService(plita, pb).replace_in_queue(
        product_kind="pile",
        mark=_LAT_C14,
        target_mark=_CYR_C14,
        changed_by="иванова",
    )

    assert result.queue_notice is None
    assert {item.kp_id for item in result.written} == {first, second}
    assert result.skipped == ()
    with sqlite3.connect(plita) as conn:
        marks = dict(conn.execute("SELECT kp_id, mark FROM kp_piles"))
        history = conn.execute(
            "SELECT source, changed_by FROM kp_mark_history ORDER BY kp_id"
        ).fetchall()
    assert marks == {first: _CYR_C14, second: _CYR_C14}
    discount_a, total_a = _offer_bits(plita, first)
    discount_b, total_b = _offer_bits(plita, second)
    assert discount_a == pytest.approx(10)
    assert discount_b == pytest.approx(25)
    assert total_a != pytest.approx(total_b)
    assert history == [("replace_economist", "иванова"), ("replace_economist", "иванова")]
    assert _states(pb)[_LAT_C14] == "resolved"
    assert _states(pb)[_CYR_C14] == "open"


def test_approval_kp_is_skipped_and_queue_keeps_the_mark(tmp_path: Path) -> None:
    from app.services.guid_replace_service import GuidReplaceService

    pb = _pb(tmp_path)
    plita = _plita(tmp_path)
    _seed_pile_price(pb, _CYR_C14, "B25", 10000.0)
    archived = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [_pile(mark=_LAT_C14, name=_LAT_C14)],
        customer_name="Архив",
        product_type="piles",
        status="в архиве",
        db_path=plita,
    )
    waiting = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [_pile(line_id="ln_w", mark=_LAT_C14, name=_LAT_C14)],
        customer_name="Согласование",
        product_type="piles",
        status="на согласовании",
        db_path=plita,
    )
    _demand(pb, "pile", _LAT_C14, [archived, waiting])

    result = GuidReplaceService(plita, pb).replace_in_queue(
        product_kind="pile",
        mark=_LAT_C14,
        target_mark=_CYR_C14,
    )

    assert [item.kp_id for item in result.written] == [archived]
    assert len(result.skipped) == 1
    assert "на согласовании" in result.skipped[0].reason
    with sqlite3.connect(plita) as conn:
        marks = dict(conn.execute("SELECT kp_id, mark FROM kp_piles"))
    assert marks[archived] == _CYR_C14
    assert marks[waiting] == _LAT_C14
    assert _states(pb)[_LAT_C14] == "open"


def test_plate_in_plan_is_not_rewritten(tmp_path: Path) -> None:
    from app.services.guid_replace_service import GuidReplaceService

    pb = _pb(tmp_path)
    plita = _plita(tmp_path)
    _seed_plate_price(pb, 51, 12, 4000.0)
    with sqlite3.connect(pb) as conn:
        conn.execute(
            'CREATE TABLE IF NOT EXISTS prays_plity ('
            '"Уникальный идентификатор (Номенклатура)" TEXT, "Товар" TEXT)'
        )
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [_plate(name="Плиты ПБ 60-12-8п", qty=1)],
        customer_name="План",
        product_type="plates",
        status="в архиве",
        db_path=plita,
    )
    with sqlite3.connect(plita) as conn:
        conn.execute("UPDATE kp_plates SET status = 'в плане' WHERE kp_id = ?", (kp_id,))
    _demand(pb, "plate", "Плиты ПБ 60-12-8п", [kp_id])

    result = GuidReplaceService(plita, pb).replace_in_queue(
        product_kind="plate",
        mark="Плиты ПБ 60-12-8п",
        target_mark="Плиты ПБ 51-12-12,5п",
        nomenclature_id="guid-plan",
    )

    assert result.written == ()
    assert "плане" in result.skipped[0].reason
    with sqlite3.connect(plita) as conn:
        name = conn.execute(
            "SELECT plate_name FROM kp_plates WHERE kp_id = ?", (kp_id,)
        ).fetchone()
    assert name is not None
    assert name[0] == "Плиты ПБ 60-12-8п"


def test_plate_single_guid_is_filled_when_client_omits_it(tmp_path: Path) -> None:
    """Один номер у выбранной плиты пишется в строку, даже если клиент GUID не прислал."""
    from app.services.guid_replace_service import GuidReplaceService

    pb = _pb(tmp_path)
    plita = _plita(tmp_path)
    _seed_plate_price(pb, 51, 12, 4000.0)
    with sqlite3.connect(pb) as conn:
        conn.execute(
            'CREATE TABLE IF NOT EXISTS prays_plity ('
            '"Уникальный идентификатор (Номенклатура)" TEXT, "Товар" TEXT)'
        )
        conn.execute(
            'INSERT INTO prays_plity ("Уникальный идентификатор (Номенклатура)", "Товар") '
            "VALUES ('guid-one', 'Плиты ПБ 51-12-12,5п')"
        )
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [_plate(name="Плиты ПБ 51-12-12п", qty=1, unit_price=100.0, plan_id=None)],
        customer_name="Плита",
        product_type="plates",
        status="в архиве",
        db_path=plita,
    )
    _demand(pb, "plate", "Плиты ПБ 51-12-12п", [kp_id])

    result = GuidReplaceService(plita, pb).replace_in_queue(
        product_kind="plate",
        mark="Плиты ПБ 51-12-12п",
        target_mark="Плиты ПБ 51-12-12,5п",
    )

    assert [item.kp_id for item in result.written] == [kp_id]
    with sqlite3.connect(plita) as conn:
        row = conn.execute(
            "SELECT plate_name, load_class, nomenclature_id FROM kp_plates WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()
    assert row == ("Плиты ПБ 51-12-12,5п", 1250, "guid-one")


def test_second_kp_error_rolls_back_the_first(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from app.services import guid_replace_service as mod
    from app.services.guid_replace_service import GuidReplaceService

    pb = _pb(tmp_path)
    plita = _plita(tmp_path)
    _seed_pile_price(pb, _CYR_C14, "B25", 10000.0)
    first = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [_pile(mark=_LAT_C14, name=_LAT_C14)],
        customer_name="Первый",
        product_type="piles",
        status="в архиве",
        db_path=plita,
    )
    second = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [_pile(line_id="ln_p2", mark=_LAT_C14, name=_LAT_C14)],
        customer_name="Второй",
        product_type="piles",
        status="в архиве",
        db_path=plita,
    )
    _demand(pb, "pile", _LAT_C14, [first, second])
    real = mod.update_kp_from_order_data
    calls = {"n": 0}

    def boom(*args: object, **kwargs: object) -> int:
        calls["n"] += 1
        if calls["n"] == 2:
            raise RuntimeError("второй КП")
        return real(*args, **kwargs)

    monkeypatch.setattr(mod, "update_kp_from_order_data", boom)
    with pytest.raises(RuntimeError, match="второй КП"):
        GuidReplaceService(plita, pb).replace_in_queue(
            product_kind="pile",
            mark=_LAT_C14,
            target_mark=_CYR_C14,
        )
    with sqlite3.connect(plita) as conn:
        marks = dict(conn.execute("SELECT kp_id, mark FROM kp_piles"))
    assert marks[first] == _LAT_C14
    assert marks[second] == _LAT_C14


def test_queue_failure_keeps_successful_replace(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture) -> None:
    from app.services.guid_replace_service import GuidReplaceService

    pb = _pb(tmp_path)
    plita = _plita(tmp_path)
    _seed_pile_price(pb, _CYR_C14, "B25", 10000.0)
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [_pile(mark=_LAT_C14, name=_LAT_C14)],
        customer_name="Очередь",
        product_type="piles",
        status="в архиве",
        db_path=plita,
    )
    _demand(pb, "pile", _LAT_C14, [kp_id])

    def broken(*_args: object, **_kwargs: object) -> None:
        raise RuntimeError("pb.db недоступна")

    monkeypatch.setattr("app.services.guid_queue_sync.sync_guid_queue_with_kp", broken)
    with caplog.at_level("ERROR"):
        result = GuidReplaceService(plita, pb).replace_in_queue(
            product_kind="pile",
            mark=_LAT_C14,
            target_mark=_CYR_C14,
        )
    assert result.queue_notice == "Очередь обновится при открытии страницы."
    assert result.written[0].kp_id == kp_id
    assert any("сверить" in record.message.casefold() for record in caplog.records)
    with sqlite3.connect(plita) as conn:
        mark = conn.execute("SELECT mark FROM kp_piles WHERE kp_id = ?", (kp_id,)).fetchone()
    assert mark is not None and mark[0] == _CYR_C14


def test_choosing_one_of_two_cards_closes_the_task(tmp_path: Path) -> None:
    from app.services.guid_replace_service import GuidReplaceService
    from core.bridge_pile_price_db import init_bridge_pile_prices_schema
    from core.nomenclature_guid import ensure_schema as ensure_guid_schema
    from core.nomenclature_guid import get_by_mark, upsert

    pb = _pb(tmp_path)
    plita = _plita(tmp_path)
    init_bridge_pile_prices_schema(pb)
    with sqlite3.connect(pb) as conn:
        conn.execute(
            "INSERT INTO bridge_pile_prices (mark, concrete_grade, price) VALUES (?, 'B25', 7000)",
            (_CYR_C15,),
        )
    guid_conn = sqlite3.connect(pb)
    try:
        ensure_guid_schema(guid_conn)
        upsert(guid_conn, "bridge_pile", _CYR_C15, match_status="ambiguous")
        guid_conn.commit()
    finally:
        guid_conn.close()
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [
            {
                "line_id": "ln_b",
                "product_type": "bridge_piles",
                "product_kind": "bridge_pile",
                "mark": _LAT_C15,
                "name": _LAT_C15,
                "concrete_grade": "B25",
                "qty": 1,
                "unit_price": 100.0,
            }
        ],
        customer_name="Дубль",
        product_type="bridge_piles",
        status="в архиве",
        db_path=plita,
    )
    _demand(pb, "bridge_pile", _LAT_C15, [kp_id])

    GuidReplaceService(plita, pb).replace_in_queue(
        product_kind="bridge_pile",
        mark=_LAT_C15,
        target_mark=_CYR_C15,
        nomenclature_id=_CHOSEN_GUID,
        changed_by="экономист",
    )

    states = _states(pb)
    assert states[_LAT_C15] == "resolved"
    assert states.get(_CYR_C15) != "open"
    check = sqlite3.connect(pb)
    try:
        row = get_by_mark(check, "bridge_pile", _CYR_C15)
    finally:
        check.close()
    assert row is not None
    assert row.match_status == "manual"
    assert row.guid_1c == _CHOSEN_GUID


def test_archive_card_returns_mark_history(tmp_path: Path) -> None:
    from app.repositories.kp_archive_repository import KpArchiveRepository
    from app.services.archive_service import ArchiveService

    plita = _plita(tmp_path)
    with sqlite3.connect(plita) as conn:
        conn.execute(
            "INSERT INTO kp_mark_history (kp_id, line_id, product_kind, mark_old, mark_new, "
            "changed_by, changed_at, source) VALUES (7, 'ln', 'plates', '12п', '12,5п', "
            "'иванова', '2026-11-12T10:00:00', 'edit')"
        )
    service = ArchiveService(
        repository=KpArchiveRepository(db_path=plita),
        outputs_dir=tmp_path / "out",
    )
    notes = service._mark_changes(7)
    assert len(notes) == 1
    assert notes[0].mark_old == "12п"
    assert notes[0].mark_new == "12,5п"
    assert notes[0].changed_by == "иванова"
    assert service._mark_changes(8) == []
    with sqlite3.connect(plita) as conn:
        conn.execute("DROP TABLE kp_mark_history")
    assert service._mark_changes(7) == []

