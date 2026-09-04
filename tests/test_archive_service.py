from __future__ import annotations

import asyncio
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock

import pytest

from app.services.archive_service import (
    ArchiveNotFoundError,
    ArchiveService,
    ArchiveValidationError,
)
from core.cargo_delivery_pricing import (
    delivery_service_charge_rub,
    total_order_cargo_weight_kg,
)
from core.kp_plate_weight import resolve_kp_line_weight_kg

ADMIN = {"id": 1, "role": "admin"}


def _make_raw(**overrides: Any) -> dict:
    base = {
        "kp_id": 42,
        "creation_date": "01.03.2026",
        "customer_name": "ООО Тест",
        "manager_name": "Иван Иванов",
        "discount_percent": 5.0,
        "subtotal": 1000.0,
        "vat_amount": 220.0,
        "total_amount": 1220.0,
        "logistics_cost": 0.0,
        "delivery_conditions": "Самовывоз",
        "payment_conditions": "100% предоплата",
        "execution_terms": "",
        "status": "в архиве",
        "owner_user_id": 1,
        "plates": [
            {
                "position_number": 1,
                "plate_name": "ПБ 78-12-8п",
                "length_m": 7.8,
                "width_m": 1.2,
                "qty": 2,
                "load_class": 800,
                "unit_price": 500.0,
                "discounted_price": 475.0,
                "unit_weight": 1000,
                "total_weight": 2000,
            }
        ],
    }
    base.update(overrides)
    return base


def _passthrough_promise_service(repository: MagicMock) -> MagicMock:
    stub = MagicMock()

    def _commit(kp_id, execution_terms, *, user, raw=None):
        from core.kp.offers_write import commit_move_to_production

        return commit_move_to_production(kp_id, execution_terms, repository.db_path)

    stub.commit_move_with_gate.side_effect = _commit
    return stub


def _make_service(
    repository: MagicMock,
    tmp_path: Path,
    promise_service: MagicMock | None = None,
) -> ArchiveService:
    repository.db_path = str(tmp_path / "plita.db")
    return ArchiveService(
        repository=repository,
        outputs_dir=tmp_path,
        promise_service=promise_service or _passthrough_promise_service(repository),
    )


def test_list_offers_for_archived_skips_completion(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.list_by_section.return_value = [_make_raw()]
    service = _make_service(repository, tmp_path)

    items = service.list_offers("archived", user=ADMIN)

    assert len(items) == 1
    assert items[0].kp_id == 42
    assert items[0].product_type == "plates"
    assert items[0].completion_percentage is None
    repository.get_completion_percentage.assert_not_called()
    repository.list_by_section.assert_called_once_with("archived", product_type="all")


def test_list_offers_for_production_includes_completion(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.list_by_section.return_value = [_make_raw(status="в работе")]
    repository.get_completion_percentage.return_value = {"percentage": 42.5}
    service = _make_service(repository, tmp_path)

    items = service.list_offers("in_production", user=ADMIN)

    assert items[0].completion_percentage == pytest.approx(42.5)
    repository.get_completion_percentage.assert_called_once_with(42)


def test_get_details_raises_when_missing(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = None
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.get_details(999, user=ADMIN)


def test_get_details_mixed_cargo_weight_uses_plates_only_helper(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """MNA-303: archive details cargo_kg ignores non-plates (plates-only helper)."""
    plate_line = {
        "name": "ПБ 60-12-8п",
        "product_type": "plates",
        "qty": 2,
        "length_m": 1.0,
        "width_m": 1.0,
        "unit_price": 100.0,
    }
    pile_line = {
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "product_type": "piles",
        "qty": 10,
        "length_m": 3.0,
        "width_m": 0.3,
        "unit_price": 1000.0,
    }
    mixed_order = [plate_line, pile_line]
    plates_only_kg = total_order_cargo_weight_kg(
        mixed_order, product_types={"plates"}
    )
    all_lines_kg = total_order_cargo_weight_kg(mixed_order)
    assert plates_only_kg > 0.0
    assert all_lines_kg > plates_only_kg
    assert resolve_kp_line_weight_kg(pile_line)[1] > 0.0

    monkeypatch.setattr(
        "app.services.archive_service.order_data_from_kp_info",
        lambda _raw: mixed_order,
    )

    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(
        # Header type stays schema-valid; mixed composition via order_data spy.
        product_type="plates",
        logistics_cost=500.0,
        plates=[
            {
                "position_number": 1,
                "plate_name": "ПБ 60-12-8п",
                "length_m": 1.0,
                "width_m": 1.0,
                "qty": 2,
                "load_class": 800,
                "unit_price": 100.0,
                "discounted_price": 100.0,
                "unit_weight": 0.0,
                "total_weight": 0.0,
                "product_type": "plates",
                "line_id": "ln_plate",
            }
        ],
        piles=[
            {
                "position_number": 2,
                "mark": "С120.35-12",
                "concrete_grade": "B25",
                "qty": 10,
                "unit_price": 1000.0,
                "discounted_price": 1000.0,
                "product_type": "piles",
                "line_id": "ln_pile",
            }
        ],
    )
    service = _make_service(repository, tmp_path)

    details = service.get_details(42, user=ADMIN)

    assert details.total_cargo_weight_kg == pytest.approx(plates_only_kg)
    assert details.total_cargo_weight_kg != pytest.approx(all_lines_kg)
    assert details.delivery_service_total_rub == pytest.approx(
        delivery_service_charge_rub(500.0, plates_only_kg)
    )


def test_get_details_mixed_cargo_calls_helper_with_plates_filter(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """MNA-303: _to_details must pass product_types={'plates'} into cargo helper."""
    captured: dict[str, object] = {}

    def _spy_cargo(order_data, product_types=None):  # type: ignore[no-untyped-def]
        captured["product_types"] = product_types
        return 1234.5

    monkeypatch.setattr(
        "app.services.archive_service.total_order_cargo_weight_kg",
        _spy_cargo,
    )
    monkeypatch.setattr(
        "app.services.archive_service.order_data_from_kp_info",
        lambda _raw: [
            {"name": "ПБ", "product_type": "plates", "qty": 1, "length_m": 1, "width_m": 1},
            {"name": "С30", "product_type": "piles", "qty": 1, "length_m": 3, "width_m": 0.3},
        ],
    )

    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(
        product_type="plates",
        piles=[{"position_number": 2, "mark": "С30", "qty": 1, "concrete_grade": "B25"}],
    )
    service = _make_service(repository, tmp_path)

    details = service.get_details(42, user=ADMIN)

    assert captured.get("product_types") == {"plates"}
    assert details.total_cargo_weight_kg == pytest.approx(1234.5)


def test_update_discount_out_of_range(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw()
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveValidationError):
        service.update_discount(1, 150, user=ADMIN)


def test_update_discount_not_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw()
    repository.update_discount.return_value = False
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.update_discount(1, 10, user=ADMIN)


def test_update_logistics_cost_not_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw()
    repository.update_logistics_cost.return_value = False
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.update_logistics_cost(1, 100.0, user=ADMIN)


def test_update_logistics_cost_calls_repository_and_returns_details(tmp_path: Path) -> None:
    """После успешного апдейта возвращаются детали с logistics_cost из свежего снимка БД (get_by_id)."""
    repository = MagicMock()
    repository.update_logistics_cost.return_value = True
    repository.get_by_id.return_value = _make_raw(
        discount_percent=0.0,
        subtotal=607.0,
        vat_amount=143.0,
        total_amount=750.0,
        logistics_cost=100.0,
        plates=[
            {
                "position_number": 1,
                "plate_name": "ПБ",
                "length_m": 1.0,
                "width_m": 1.0,
                "qty": 65,
                "load_class": 800,
                "unit_price": 10.0,
                "discounted_price": 10.0,
                "unit_weight": 0.0,
                "total_weight": 0.0,
            }
        ],
    )
    service = _make_service(repository, tmp_path)

    details = service.update_logistics_cost(42, 100.0, user=ADMIN)

    repository.update_logistics_cost.assert_called_once_with(
        42, 100.0, pile_logistics_cost=None, pile_trip_overrides=None
    )
    assert details.logistics_cost == 100.0
    assert details.finance.total_amount == 750.0
    assert details.delivery_service_total_rub == pytest.approx(100.0)


def test_move_to_production_requires_archived_status(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(status="в работе")
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveValidationError):
        service.move_to_production(42, "5 дней", user=ADMIN)


def test_move_to_production_happy_path(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    repository = MagicMock()
    repository.db_path = str(tmp_path / "plita.db")
    repository.get_by_id.side_effect = [
        _make_raw(status="в архиве"),
        _make_raw(status="в работе", execution_terms="01.04.2026"),
    ]
    commit = MagicMock(return_value=2)
    monkeypatch.setattr(
        "core.kp.offers_write.commit_move_to_production", commit
    )
    monkeypatch.setattr(
        "app.services.archive_service.ArchiveService._load_occupancy",
        staticmethod(lambda: {}),
    )
    service = _make_service(repository, tmp_path)
    service._today_override = "2026-03-01"

    details = service.move_to_production(42, "01.04.2026", user=ADMIN)

    assert details.status == "в работе"
    assert details.execution_terms == "01.04.2026"
    commit.assert_called_once_with(42, "01.04.2026", repository.db_path)


def test_move_to_production_blocks_before_promised_date(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from datetime import date

    from app.services.promise_service import PromiseGateError

    repository = MagicMock()
    repository.db_path = str(tmp_path / "plita.db")
    repository.get_by_id.return_value = _make_raw(status="в архиве")
    commit = MagicMock(return_value=2)
    monkeypatch.setattr(
        "core.kp.offers_write.commit_move_to_production", commit
    )
    promise_service = MagicMock()
    promise_service.commit_move_with_gate.side_effect = PromiseGateError(
        "Срок раньше ближайшей возможной даты 06.03.2026.",
        earliest=date(2026, 3, 6),
    )
    service = _make_service(repository, tmp_path, promise_service=promise_service)

    with pytest.raises(ArchiveValidationError, match="06.03.2026"):
        service.move_to_production(42, "2026-03-03", user=ADMIN)

    commit.assert_not_called()


def test_move_to_production_normalizes_iso_date(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    repository = MagicMock()
    repository.db_path = str(tmp_path / "plita.db")
    repository.get_by_id.side_effect = [
        _make_raw(status="в архиве"),
        _make_raw(status="в работе", execution_terms="05.06.2026"),
    ]
    commit = MagicMock(return_value=1)
    monkeypatch.setattr(
        "core.kp.offers_write.commit_move_to_production", commit
    )
    monkeypatch.setattr(
        "app.services.archive_service.ArchiveService._load_occupancy",
        staticmethod(lambda: {}),
    )
    service = _make_service(repository, tmp_path)
    service._today_override = "2026-03-01"

    service.move_to_production(42, "2026-06-05", user=ADMIN)

    commit.assert_called_once_with(42, "05.06.2026", repository.db_path)


def test_move_to_production_normalizes_ddmmyyyy(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    repository = MagicMock()
    repository.db_path = str(tmp_path / "plita.db")
    repository.get_by_id.side_effect = [
        _make_raw(status="в архиве"),
        _make_raw(status="в работе", execution_terms="05.06.2026"),
    ]
    commit = MagicMock(return_value=1)
    monkeypatch.setattr(
        "core.kp.offers_write.commit_move_to_production", commit
    )
    monkeypatch.setattr(
        "app.services.archive_service.ArchiveService._load_occupancy",
        staticmethod(lambda: {}),
    )
    service = _make_service(repository, tmp_path)
    service._today_override = "2026-03-01"

    service.move_to_production(42, "05.06.2026", user=ADMIN)

    commit.assert_called_once_with(42, "05.06.2026", repository.db_path)


def test_move_to_production_normalizes_five_days(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from datetime import datetime as dt_cls

    base = dt_cls(2026, 5, 31, 12, 0, 0)

    class FixedNowDatetime(dt_cls):
        @classmethod
        def now(cls, tz=None):
            return cls(base.year, base.month, base.day, base.hour, base.minute, base.second)

    monkeypatch.setattr("core.execution_terms.datetime", FixedNowDatetime)
    repository = MagicMock()
    repository.db_path = str(tmp_path / "plita.db")
    repository.get_by_id.side_effect = [
        _make_raw(status="в архиве"),
        _make_raw(status="в работе", execution_terms="05.06.2026"),
    ]
    commit = MagicMock(return_value=1)
    monkeypatch.setattr(
        "core.kp.offers_write.commit_move_to_production", commit
    )
    monkeypatch.setattr(
        "app.services.archive_service.ArchiveService._load_occupancy",
        staticmethod(lambda: {}),
    )
    service = _make_service(repository, tmp_path)
    service._today_override = "2026-05-31"

    service.move_to_production(42, "5 дней", user=ADMIN)

    commit.assert_called_once_with(42, "05.06.2026", repository.db_path)


def test_move_to_production_raises_archive_error_on_commit_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveError

    repository = MagicMock()
    repository.db_path = str(tmp_path / "plita.db")
    repository.get_by_id.return_value = _make_raw(status="в архиве")
    monkeypatch.setattr(
        "core.kp.offers_write.commit_move_to_production",
        MagicMock(side_effect=RuntimeError("boom")),
    )
    monkeypatch.setattr(
        "app.services.archive_service.ArchiveService._load_occupancy",
        staticmethod(lambda: {}),
    )
    service = _make_service(repository, tmp_path)
    service._today_override = "2026-03-01"

    with pytest.raises(ArchiveError, match="Не удалось перевести"):
        service.move_to_production(42, "01.04.2026", user=ADMIN)


def test_parse_execution_terms_formats() -> None:
    assert ArchiveService._parse_execution_terms("01.04.2026") == "01.04.2026"
    assert ArchiveService._parse_execution_terms("2026-04-01") == "01.04.2026"
    result_days = ArchiveService._parse_execution_terms("7 дней")
    assert len(result_days) == 10


def test_parse_execution_terms_invalid() -> None:
    with pytest.raises(ArchiveValidationError):
        ArchiveService._parse_execution_terms("")
    with pytest.raises(ArchiveValidationError):
        ArchiveService._parse_execution_terms("скоро")


def test_estimate_production(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(
        plates=[
            {"length_m": 50.0, "qty": 2},
            {"length_m": 100.0, "qty": 1},
        ]
    )
    service = _make_service(repository, tmp_path)

    estimate = service.estimate_production(42, user=ADMIN)

    assert estimate["total_length_m"] == 200.0
    assert estimate["estimated_tracks"] >= 1
    assert estimate["estimated_days"] >= 1


def test_delete_offer_not_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = None
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveNotFoundError):
        service.delete_offer(42, user=ADMIN)


def test_delete_offer_releases_active_promises_before_delete(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(kp_id=42)
    repository.delete.return_value = True
    promise_service = MagicMock()
    service = _make_service(repository, tmp_path, promise_service=promise_service)

    service.delete_offer(42, user=ADMIN)

    promise_service.release_on_delete.assert_called_once_with(42)
    repository.delete.assert_called_once_with(42)


def test_generate_document_pdf(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw()
    service = _make_service(repository, tmp_path)

    class FakeBuffer:
        def getvalue(self) -> bytes:
            return b"%PDF-FAKE"

    fake_pdf = MagicMock(return_value=FakeBuffer())
    monkeypatch.setattr(
        "app.services.archive_service.generate_commercial_offer_pdf",
        fake_pdf,
    )

    path = asyncio.run(service.generate_document(42, "pdf", user=ADMIN))

    assert path.exists()
    assert path.name == "КП_42.pdf"
    assert path.read_bytes() == b"%PDF-FAKE"
    fake_pdf.assert_called_once()
    call_kwargs = fake_pdf.call_args.kwargs
    assert call_kwargs["delivery_conditions"] == "Самовывоз"
    assert call_kwargs["payment_conditions"] == "100% предоплата"


def test_generate_document_rejects_empty_plates(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(plates=[])
    service = _make_service(repository, tmp_path)

    with pytest.raises(ArchiveValidationError):
        asyncio.run(service.generate_document(42, "xlsx", user=ADMIN))


def test_search_by_number_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(kp_id=42)
    service = _make_service(repository, tmp_path)

    result = service.search(user=ADMIN, kp_id=42)

    assert result.mode == "number"
    assert result.total == 1
    assert result.items[0].kp_id == 42
    assert result.truncated is False


def test_search_by_number_not_found(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = None
    service = _make_service(repository, tmp_path)

    result = service.search(user=ADMIN, kp_id=999)

    assert result.mode == "number"
    assert result.total == 0
    assert result.items == []


def test_search_by_customer_delegates_to_repository(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.search_by_customer_name.return_value = (
        [_make_raw(kp_id=10, customer_name="ООО Ромашка")],
        1,
    )
    service = _make_service(repository, tmp_path)

    result = service.search(user=ADMIN, customer="Ромашка")

    assert result.mode == "customer"
    assert result.total == 1
    assert result.items[0].customer_name == "ООО Ромашка"
    assert result.truncated is False
    repository.search_by_customer_name.assert_called_once_with("Ромашка", limit=50)


def test_search_by_customer_truncated_flag(tmp_path: Path) -> None:
    repository = MagicMock()
    repository.search_by_customer_name.return_value = (
        [_make_raw(kp_id=i) for i in range(50)],
        60,
    )
    service = _make_service(repository, tmp_path)

    result = service.search(user=ADMIN, customer="Тест")

    assert result.truncated is True
    assert result.total == 60
    assert len(result.items) == 50


# --- MNA-602: multi badges (product_types) + filter «contains type» -------------


def test_archive_offer_list_item_schema_has_product_types() -> None:
    """MNA-602 / Q3: list DTO exposes product_types for N type badges."""
    from app.schemas.archive import ArchiveOfferListItem

    assert "product_types" in ArchiveOfferListItem.model_fields, (
        "ArchiveOfferListItem.product_types missing (MNA-602)"
    )
    item = ArchiveOfferListItem(
        kp_id=7,
        product_types=["plates", "piles"],
    )
    assert item.product_types == ["plates", "piles"]


def test_list_offers_mono_serializes_single_product_types(tmp_path: Path) -> None:
    """MNA-602: mono KP → product_types is a one-element list matching product_type."""
    from app.schemas.archive import ArchiveOfferListItem

    assert "product_types" in ArchiveOfferListItem.model_fields

    repository = MagicMock()
    repository.list_by_section.return_value = [_make_raw(product_type="piles")]
    service = _make_service(repository, tmp_path)

    items = service.list_offers("archived", user=ADMIN)

    assert len(items) == 1
    assert items[0].product_type == "piles"
    assert items[0].product_types == ["piles"]


def test_list_offers_mixed_serializes_product_types_badges(tmp_path: Path) -> None:
    """MNA-602 / Q3: mixed KP list row carries all present types for badges."""
    from app.schemas.archive import ArchiveOfferListItem

    assert "product_types" in ArchiveOfferListItem.model_fields

    repository = MagicMock()
    repository.list_by_section.return_value = [
        _make_raw(
            kp_id=42,
            product_type="mixed",
            product_types=["plates", "piles"],
            plates=[
                {
                    "position_number": 1,
                    "plate_name": "ПБ 60-12-8п",
                    "qty": 1,
                    "product_type": "plates",
                }
            ],
            piles=[
                {
                    "position_number": 2,
                    "mark": "С120.35-12",
                    "qty": 2,
                    "concrete_grade": "B25",
                    "product_type": "piles",
                }
            ],
        )
    ]
    service = _make_service(repository, tmp_path)

    items = service.list_offers("archived", user=ADMIN)

    assert len(items) == 1
    assert set(items[0].product_types) == {"plates", "piles"}
    assert len(items[0].product_types) == 2


def test_list_offers_filter_plates_includes_mixed_with_plates(
    tmp_path: Path,
) -> None:
    """MNA-602 / Q3: product_type=plates filter is «contains type» (mixed with plates OK)."""
    from app.repositories.kp_archive_repository import KpArchiveRepository
    from core.kp_db_schema import init_schema
    from core.kp_persistence_service import KpPersistenceService

    db_path = str(tmp_path / "plita.db")
    init_schema(db_path)

    plate_line = {
        "line_id": "ln_plate_1",
        "product_type": "plates",
        "name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 1,
        "unit_price": 1000.0,
        "weight": 500.0,
        "concrete_grade": "М500",
    }
    pile_line = {
        "line_id": "ln_pile_1",
        "product_type": "piles",
        "product_kind": "pile",
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "concrete_grade": "B25",
        "qty": 2,
        "unit_price": 100.0,
    }

    mixed_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [plate_line, {**pile_line, "line_id": "ln_pile_mixed"}],
        customer_name="Mixed with plates",
        status="в архиве",
        db_path=db_path,
    )
    piles_only_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [{**pile_line, "line_id": "ln_pile_only"}],
        customer_name="Piles only",
        status="в архиве",
        product_type="piles",
        db_path=db_path,
    )
    plates_only_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [{**plate_line, "line_id": "ln_plate_only"}],
        customer_name="Plates only",
        status="в архиве",
        db_path=db_path,
    )

    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path / "out",
    )
    items = service.list_offers("archived", product_type="plates", user=ADMIN)
    ids = {item.kp_id for item in items}

    assert mixed_id in ids, "mixed-with-plates must pass filter product_type=plates"
    assert plates_only_id in ids
    assert piles_only_id not in ids


def test_list_offers_filter_piles_includes_mixed_with_piles(
    tmp_path: Path,
) -> None:
    """MNA-602: product_type=piles also uses contains-type (mixed with piles included)."""
    from app.repositories.kp_archive_repository import KpArchiveRepository
    from core.kp_db_schema import init_schema
    from core.kp_persistence_service import KpPersistenceService

    db_path = str(tmp_path / "plita.db")
    init_schema(db_path)

    plate_line = {
        "line_id": "ln_plate_1",
        "product_type": "plates",
        "name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 1,
        "unit_price": 1000.0,
        "weight": 500.0,
        "concrete_grade": "М500",
    }
    pile_line = {
        "line_id": "ln_pile_1",
        "product_type": "piles",
        "product_kind": "pile",
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "concrete_grade": "B25",
        "qty": 2,
        "unit_price": 100.0,
    }

    mixed_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [plate_line, pile_line],
        customer_name="Mixed with piles",
        status="в архиве",
        db_path=db_path,
    )
    plates_only_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [{**plate_line, "line_id": "ln_plate_only"}],
        customer_name="Plates only",
        status="в архиве",
        db_path=db_path,
    )

    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path / "out",
    )
    items = service.list_offers("archived", product_type="piles", user=ADMIN)
    ids = {item.kp_id for item in items}

    assert mixed_id in ids, "mixed-with-piles must pass filter product_type=piles"
    assert plates_only_id not in ids


def test_list_offers_mixed_from_db_exposes_product_types(
    tmp_path: Path,
) -> None:
    """MNA-602: real mixed KP list row exposes product_types for UI badges."""
    from app.repositories.kp_archive_repository import KpArchiveRepository
    from app.schemas.archive import ArchiveOfferListItem
    from core.kp_db_schema import init_schema
    from core.kp_persistence_service import KpPersistenceService

    assert "product_types" in ArchiveOfferListItem.model_fields

    db_path = str(tmp_path / "plita.db")
    init_schema(db_path)

    mixed_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [
            {
                "line_id": "ln_plate_1",
                "product_type": "plates",
                "name": "ПБ 60-12-8п",
                "length_m": 6.0,
                "width_m": 1.2,
                "load_class": 800,
                "qty": 1,
                "unit_price": 1000.0,
                "weight": 500.0,
                "concrete_grade": "М500",
            },
            {
                "line_id": "ln_pile_1",
                "product_type": "piles",
                "product_kind": "pile",
                "name": "С120.35-12",
                "mark": "С120.35-12",
                "concrete_grade": "B25",
                "qty": 2,
                "unit_price": 100.0,
            },
        ],
        customer_name="Mixed badges",
        status="в архиве",
        db_path=db_path,
    )

    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path / "out",
    )
    items = service.list_offers("archived", user=ADMIN)
    item = next(row for row in items if row.kp_id == mixed_id)

    assert set(item.product_types) == {"plates", "piles"}
    assert len(item.product_types) == 2
