from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest
from fastapi.testclient import TestClient

from tests.helpers.csrf import CsrfAwareTestClient

from app.dependencies.services import get_archive_service
from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from tests.helpers.auth_fixtures import patch_auth_users
from app.schemas.archive import (
    ArchiveOfferDetails,
    ArchiveOfferFinance,
    ArchiveOfferListItem,
    ArchiveSearchResponse,
    KpReadinessPositionItem,
    KpReadinessPositionsResponse,
    KpReadinessStep,
    KpReadinessStepState,
    KpReadinessSummary,
)
from app.schemas.sgp import SgpProgress
from app.security.session import create_session_token

TESTER_USER = {
    "id": 1,
    "username": "tester",
    "role": "admin",
    "manager_id": None,
    "is_active": 1,
    "created_at": "2026-01-01 00:00:00",
    "session_version": 0,
}


@pytest.fixture()
def fake_service() -> MagicMock:
    return MagicMock()


@pytest.fixture()
def client(
    monkeypatch: pytest.MonkeyPatch,
    fake_service: MagicMock,
) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    get_settings.cache_clear()
    patch_auth_users(
        monkeypatch,
        [
            {
                "id": 1,
                "username": "tester",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
                "session_version": 0,
            }
        ],
    )
    app = create_app()
    app.dependency_overrides[get_archive_service] = lambda: fake_service
    return CsrfAwareTestClient(app)


@pytest.fixture()
def auth_cookie() -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": 1, "username": "tester", "role": "admin"},
            ttl_seconds=300,
        ),
    }


def _fake_details(
    kp_id: int = 42,
    status: str = "в архиве",
    *,
    finance: ArchiveOfferFinance | None = None,
    logistics_cost: float = 0.0,
    total_cargo_weight_kg: float = 0.0,
    delivery_service_total_rub: float = 0.0,
) -> ArchiveOfferDetails:
    return ArchiveOfferDetails(
        kp_id=kp_id,
        creation_date="01.03.2026",
        customer_name="ООО Тест",
        manager_name="Иван Иванов",
        status=status,
        execution_terms=None,
        delivery_conditions=None,
        payment_conditions=None,
        finance=finance
        or ArchiveOfferFinance(
            subtotal=1000, vat_amount=220, total_amount=1220, discount_percent=5
        ),
        logistics_cost=logistics_cost,
        total_cargo_weight_kg=total_cargo_weight_kg,
        delivery_service_total_rub=delivery_service_total_rub,
        plates=[],
        completion_percentage=None,
    )


def test_list_requires_auth(client: TestClient) -> None:
    response = client.get("/api/v1/commercial/archive?section=archived")
    assert response.status_code == 401


def test_list_returns_items(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.list_offers.return_value = [
        ArchiveOfferListItem(
            kp_id=42,
            creation_date="01.03.2026",
            customer_name="ООО Тест",
            manager_name="Иван",
            discount_percent=5.0,
            subtotal=1000,
            vat_amount=220,
            total_amount=1220,
            status="в архиве",
        )
    ]

    response = client.get(
        "/api/v1/commercial/archive?section=archived",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert len(payload) == 1
    assert payload[0]["kp_id"] == 42
    fake_service.list_offers.assert_called_once_with(
        "archived",
        product_type="all",
        user=TESTER_USER,
    )


def test_list_returns_product_types_for_mixed_badges(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    """MNA-602 / Q3: archive list JSON includes product_types for multi badges."""
    assert "product_types" in ArchiveOfferListItem.model_fields, (
        "ArchiveOfferListItem.product_types missing (MNA-602)"
    )
    fake_service.list_offers.return_value = [
        ArchiveOfferListItem(
            kp_id=42,
            creation_date="01.03.2026",
            customer_name="ООО Тест",
            manager_name="Иван",
            discount_percent=5.0,
            subtotal=1000,
            vat_amount=220,
            total_amount=1220,
            status="в архиве",
            product_type="plates",
            product_types=["plates", "piles"],
        )
    ]

    response = client.get(
        "/api/v1/commercial/archive?section=archived",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload[0]["product_types"] == ["plates", "piles"]


def test_get_details_404(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    from app.services.archive_service import ArchiveNotFoundError

    fake_service.get_details.side_effect = ArchiveNotFoundError("нет такого")

    response = client.get(
        "/api/v1/commercial/archive/999",
        cookies=auth_cookie,
    )

    assert response.status_code == 404


def test_update_discount_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.update_discount.return_value = _fake_details()

    response = client.patch(
        "/api/v1/commercial/archive/42/discount",
        json={"discount": 10},
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    fake_service.update_discount.assert_called_once_with(42, 10.0, user=TESTER_USER)


def test_update_discount_validation(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    response = client.patch(
        "/api/v1/commercial/archive/42/discount",
        json={"discount": 150},
        cookies=auth_cookie,
    )
    assert response.status_code == 422


def test_update_logistics_cost_requires_auth(client: TestClient) -> None:
    response = client.patch(
        "/api/v1/commercial/archive/42/logistics-cost",
        json={"logistics_cost": 100},
    )
    assert response.status_code == 401


def test_update_logistics_cost_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.update_logistics_cost.return_value = _fake_details(
        finance=ArchiveOfferFinance(
            subtotal=607.0,
            vat_amount=143.0,
            total_amount=750.0,
            discount_percent=0.0,
        ),
        logistics_cost=100.0,
        total_cargo_weight_kg=18414.5,
        delivery_service_total_rub=100.0,
    )

    response = client.patch(
        "/api/v1/commercial/archive/42/logistics-cost",
        json={"logistics_cost": 100},
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    fake_service.update_logistics_cost.assert_called_once_with(42, 100.0, user=TESTER_USER)
    payload = response.json()
    assert payload["logistics_cost"] == 100.0
    assert payload["finance"]["total_amount"] == 750.0


def test_update_logistics_cost_validation_negative(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    response = client.patch(
        "/api/v1/commercial/archive/42/logistics-cost",
        json={"logistics_cost": -1},
        cookies=auth_cookie,
    )
    assert response.status_code == 422


def test_update_logistics_cost_404(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    from app.services.archive_service import ArchiveNotFoundError

    fake_service.update_logistics_cost.side_effect = ArchiveNotFoundError("нет такого")

    response = client.patch(
        "/api/v1/commercial/archive/999/logistics-cost",
        json={"logistics_cost": 50},
        cookies=auth_cookie,
    )

    assert response.status_code == 404


def test_delete_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    response = client.delete(
        "/api/v1/commercial/archive/42",
        cookies=auth_cookie,
    )

    assert response.status_code == 204
    fake_service.delete_offer.assert_called_once_with(42, user=TESTER_USER)


def test_move_to_production_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.move_to_production.return_value = _fake_details(status="в работе")

    response = client.post(
        "/api/v1/commercial/archive/42/move-to-production",
        json={"execution_terms": "5 дней"},
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    assert response.json()["status"] == "в работе"
    fake_service.move_to_production.assert_called_once_with(42, "5 дней", user=TESTER_USER)


def test_move_to_production_validation_error(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    from app.services.archive_service import ArchiveValidationError

    fake_service.move_to_production.side_effect = ArchiveValidationError("bad")

    response = client.post(
        "/api/v1/commercial/archive/42/move-to-production",
        json={"execution_terms": "скоро"},
        cookies=auth_cookie,
    )

    assert response.status_code == 400


def test_move_to_production_archive_error_returns_500(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    from app.services.archive_service import ArchiveError

    fake_service.move_to_production.side_effect = ArchiveError("tx failed")

    response = client.post(
        "/api/v1/commercial/archive/42/move-to-production",
        json={"execution_terms": "01.04.2026"},
        cookies=auth_cookie,
    )

    assert response.status_code == 500


def test_capacity_snapshot_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    from app.schemas.archive import CapacitySnapshotResponse

    fake_service.get_capacity_snapshot.return_value = CapacitySnapshotResponse(
        start_date="2026-03-03",
        target_date="2026-03-20",
        tracks_needed=2,
        tracks_free_in_window=40,
        delta=38,
        status="green",
        hint=None,
        days_info={},
        holidays=[],
        extra_workdays=[],
        calendar_from_month="2026-03",
        calendar_to_month="2026-03",
    )

    response = client.get(
        "/api/v1/commercial/archive/42/capacity-snapshot?target=2026-03-20",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    body = response.json()
    assert body["status"] == "green"
    assert body["tracks_needed"] == 2
    assert body["calendar_from_month"] == "2026-03"
    fake_service.get_capacity_snapshot.assert_called_once()


def test_capacity_snapshot_not_found(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    from app.services.archive_service import ArchiveNotFoundError

    fake_service.get_capacity_snapshot.side_effect = ArchiveNotFoundError("нет")

    response = client.get(
        "/api/v1/commercial/archive/999/capacity-snapshot?target=2026-03-20",
        cookies=auth_cookie,
    )

    assert response.status_code == 404


def test_capacity_snapshot_forbidden_for_accountant(
    monkeypatch: pytest.MonkeyPatch,
    fake_service: MagicMock,
) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    get_settings.cache_clear()
    patch_auth_users(
        monkeypatch,
        [
            {
                "id": 2,
                "username": "acc",
                "role": "accountant",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
                "session_version": 0,
            }
        ],
    )
    app = create_app()
    app.dependency_overrides[get_archive_service] = lambda: fake_service
    client = CsrfAwareTestClient(app)
    cookie = {
        "app_session": create_session_token(
            {"id": 2, "username": "acc", "role": "accountant"},
            ttl_seconds=300,
        ),
    }

    response = client.get(
        "/api/v1/commercial/archive/42/capacity-snapshot?target=2026-03-20",
        cookies=cookie,
    )
    assert response.status_code == 403


def _fake_list_item(kp_id: int = 42, customer_name: str = "ООО Тест") -> ArchiveOfferListItem:
    return ArchiveOfferListItem(
        kp_id=kp_id,
        creation_date="01.03.2026",
        customer_name=customer_name,
        manager_name="Иван",
        discount_percent=5.0,
        subtotal=1000,
        vat_amount=220,
        total_amount=1220,
        status="в архиве",
    )


def _fake_search_response(
    *,
    mode: str = "number",
    items: list[ArchiveOfferListItem] | None = None,
    total: int | None = None,
    truncated: bool = False,
) -> ArchiveSearchResponse:
    resolved_items = items if items is not None else [_fake_list_item()]
    resolved_total = total if total is not None else len(resolved_items)
    return ArchiveSearchResponse(
        mode=mode,  # type: ignore[arg-type]
        items=resolved_items,
        total=resolved_total,
        truncated=truncated,
    )


def test_search_by_number_found(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(mode="number")

    response = client.get(
        "/api/v1/commercial/archive/search?kp_id=42",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["mode"] == "number"
    assert payload["total"] == 1
    assert len(payload["items"]) == 1
    assert payload["items"][0]["kp_id"] == 42
    fake_service.search.assert_called_once_with(user=TESTER_USER, kp_id=42)


def test_search_by_number_not_found(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(mode="number", items=[], total=0)

    response = client.get(
        "/api/v1/commercial/archive/search?kp_id=999",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["mode"] == "number"
    assert payload["total"] == 0
    assert payload["items"] == []


def test_search_by_customer_returns_items(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(
        mode="customer",
        items=[_fake_list_item(10, "ООО Ромашка"), _fake_list_item(5, "ООО Ромашка-2")],
        total=2,
    )

    response = client.get(
        "/api/v1/commercial/archive/search?customer=Ромашка",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["mode"] == "customer"
    assert payload["total"] == 2
    assert len(payload["items"]) == 2
    fake_service.search.assert_called_once_with(user=TESTER_USER, customer="Ромашка")


def test_search_by_customer_not_found(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(mode="customer", items=[], total=0)

    response = client.get(
        "/api/v1/commercial/archive/search?customer=Несуществующий",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["total"] == 0
    assert payload["items"] == []


def test_search_customer_too_short_returns_400(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    response = client.get(
        "/api/v1/commercial/archive/search?customer=А",
        cookies=auth_cookie,
    )

    assert response.status_code == 400
    assert "2" in response.json()["detail"]
    fake_service.search.assert_not_called()


def test_search_without_params_returns_422(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    response = client.get(
        "/api/v1/commercial/archive/search",
        cookies=auth_cookie,
    )

    assert response.status_code == 422
    fake_service.search.assert_not_called()


def test_search_both_params_prefers_kp_id(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(mode="number")

    response = client.get(
        "/api/v1/commercial/archive/search?kp_id=42&customer=Ромашка",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    fake_service.search.assert_called_once_with(user=TESTER_USER, kp_id=42)


def test_search_customer_truncated_flag(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.search.return_value = _fake_search_response(
        mode="customer",
        items=[_fake_list_item(i, "ООО Тест") for i in range(50)],
        total=73,
        truncated=True,
    )

    response = client.get(
        "/api/v1/commercial/archive/search?customer=Тест",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["truncated"] is True
    assert payload["total"] == 73
    assert len(payload["items"]) == 50


def test_download_file_returns_file(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
    tmp_path: Path,
) -> None:
    target = tmp_path / "КП_42.pdf"
    target.write_bytes(b"%PDF-TEST")

    async def fake_generate(kp_id: int, kind: str, **kwargs) -> Path:
        return target

    fake_service.generate_document = fake_generate  # type: ignore[assignment]

    response = client.get(
        "/api/v1/commercial/archive/42/files/pdf",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    assert response.content == b"%PDF-TEST"


def _fake_readiness_summary(*, kp_id: int = 42) -> KpReadinessSummary:
    return KpReadinessSummary(
        completion_percentage=72.0,
        sgp_progress=SgpProgress(n=14, m=20),
        issuable_qty=14,
        in_production_qty=6,
        summary_text="14 из 20 шт на складе, 6 в производстве. Можно выдать 14 шт.",
        client_copy_text="Здравствуйте! По вашему заказу №42: 14 из 20 шт уже на складе.",
        steps=[
            KpReadinessStep(id="kp", label="КП", state=KpReadinessStepState.DONE),
            KpReadinessStep(
                id="production",
                label="Производство",
                state=KpReadinessStepState.ACTIVE,
                hint="72%",
            ),
            KpReadinessStep(
                id="sgp",
                label="СГП",
                state=KpReadinessStepState.ACTIVE,
                hint="14/20",
            ),
            KpReadinessStep(id="release", label="Выдача", state=KpReadinessStepState.DISABLED),
            KpReadinessStep(id="closed", label="Закрыто", state=KpReadinessStepState.DISABLED),
        ],
        release_note="Выдача с СГП — в следующем обновлении",
    )


def test_get_details_includes_readiness(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    details = _fake_details(status="в работе")
    details = details.model_copy(update={"readiness": _fake_readiness_summary()})
    fake_service.get_details.return_value = details

    response = client.get("/api/v1/commercial/archive/42", cookies=auth_cookie)

    assert response.status_code == 200
    payload = response.json()
    assert payload["readiness"] is not None
    assert payload["readiness"]["issuable_qty"] == 14
    assert payload["readiness"]["sgp_progress"]["n"] == 14


def test_get_details_readiness_null_for_archived(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.get_details.return_value = _fake_details(status="в архиве")

    response = client.get("/api/v1/commercial/archive/42", cookies=auth_cookie)

    assert response.status_code == 200
    assert response.json()["readiness"] is None


def test_get_details_readiness_null_for_completed(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.get_details.return_value = _fake_details(status="выполнено")

    response = client.get("/api/v1/commercial/archive/42", cookies=auth_cookie)

    assert response.status_code == 200
    assert response.json()["readiness"] is None


def test_readiness_positions_requires_auth(client: TestClient) -> None:
    response = client.get("/api/v1/commercial/archive/42/readiness/positions")
    assert response.status_code == 401


def test_readiness_positions_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.get_readiness_positions.return_value = KpReadinessPositionsResponse(
        items=[
            KpReadinessPositionItem(
                position_number=1,
                plate_name="ПБ 59-12-8",
                length_m=5.9,
                width_m=1.2,
                load_class=800,
                label="ПБ 59-12-8",
                ordered=10,
                in_plan=4,
                on_sgp=6,
                remaining=0,
            )
        ],
        count=1,
    )

    response = client.get(
        "/api/v1/commercial/archive/42/readiness/positions",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["count"] == 1
    assert payload["items"][0]["ordered"] == 10
    fake_service.get_readiness_positions.assert_called_once_with(42, user=TESTER_USER)


def test_readiness_positions_empty_for_archived(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    fake_service.get_readiness_positions.return_value = KpReadinessPositionsResponse(
        items=[], count=0
    )

    response = client.get(
        "/api/v1/commercial/archive/42/readiness/positions",
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    assert response.json()["items"] == []


def test_readiness_positions_404(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    from app.services.archive_service import ArchiveNotFoundError

    fake_service.get_readiness_positions.side_effect = ArchiveNotFoundError("нет")

    response = client.get(
        "/api/v1/commercial/archive/999/readiness/positions",
        cookies=auth_cookie,
    )

    assert response.status_code == 404


# --- MNA-402: HTTP archive download/regen for mixed KP (real service) ---------


_MIXED_ORDER_HTTP = [
    {
        "line_id": "ln_plate",
        "product_type": "plates",
        "name": "ПБ 60-12-8п",
        "mark": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 10,
        "unit_price": 100.0,
        "weight": 2000.0,
        # Explicit grade so resolve_concrete_grade short-circuits (no pb.db).
        "concrete_grade": "М500",
    },
    {
        "line_id": "ln_pile",
        "product_type": "piles",
        "product_kind": "pile",
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "concrete_grade": "B25",
        "qty": 3,
        "unit_price": 44634.03,
    },
]


@pytest.fixture()
def mixed_archive_client(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> tuple[TestClient, str, Path]:
    """Real app + iso DB for mixed archive file download/regen."""
    from tests.helpers import kp_db_fixtures as fx
    from tests.helpers.auth_fixtures import patch_auth_users

    db_path = fx.make_iso_db(tmp_path)
    outputs_dir = tmp_path / "archive_outputs"
    outputs_dir.mkdir()
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    monkeypatch.setenv("OUTPUTS_DIR", str(outputs_dir))
    get_settings.cache_clear()
    patch_auth_users(
        monkeypatch,
        [
            {
                "id": 1,
                "username": "tester",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
                "session_version": 0,
            }
        ],
    )
    return CsrfAwareTestClient(create_app()), db_path, outputs_dir


def _admin_cookie_mna402() -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": 1, "username": "tester", "role": "admin"},
            ttl_seconds=300,
        )
    }


def test_download_mixed_xlsx_regenerates_unified_layout(
    mixed_archive_client: tuple[TestClient, str, Path],
) -> None:
    """MNA-402: GET archive .../files/xlsx for mixed KP returns unified workbook."""
    import io

    import pandas as pd
    from core.kp_persistence_service import KpPersistenceService

    client, db_path, _outputs = mixed_archive_client
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        _MIXED_ORDER_HTTP,
        customer_name="Mixed HTTP",
        status="в архиве",
        logistics_cost=5000.0,
        db_path=db_path,
    )

    response = client.get(
        f"/api/v1/commercial/archive/{kp_id}/files/xlsx",
        cookies=_admin_cookie_mna402(),
    )
    assert response.status_code == 200, response.text
    assert (
        response.headers.get("content-type", "")
        == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    df = pd.read_excel(io.BytesIO(response.content), sheet_name="КП", header=None)
    headers: list[str] = []
    for _, row in df.iterrows():
        vals = [str(v).strip() for v in row.tolist() if pd.notna(v)]
        if vals and vals[0] == "№":
            raw = [v if pd.notna(v) else "" for v in row.tolist()]
            while raw and raw[-1] == "":
                raw.pop()
            headers = [str(c).strip() if c != "" else "" for c in raw]
            break
    assert headers == ["№", "Тип", "Наименование", "Кол-во", "Цена", "Сумма"], headers
    # PB-only delivery present when logistics_cost > 0 and plates have weight
    name_col = headers.index("Наименование")
    names = [
        str(row.tolist()[name_col]).strip()
        for _, row in df.iterrows()
        if name_col < len(row.tolist()) and pd.notna(row.tolist()[name_col])
    ]
    assert any("доставк" in n.lower() for n in names), names


def test_download_mixed_pdf_regenerates_successfully(
    mixed_archive_client: tuple[TestClient, str, Path],
) -> None:
    """MNA-402: GET archive .../files/pdf for mixed KP regenerates a PDF."""
    from core.kp_persistence_service import KpPersistenceService

    client, db_path, _outputs = mixed_archive_client
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        _MIXED_ORDER_HTTP,
        customer_name="Mixed PDF",
        status="в архиве",
        db_path=db_path,
    )

    response = client.get(
        f"/api/v1/commercial/archive/{kp_id}/files/pdf",
        cookies=_admin_cookie_mna402(),
    )
    assert response.status_code == 200, response.text
    assert response.headers["content-type"] == "application/pdf"
    assert response.content[:4] == b"%PDF"
    assert len(response.content) > 100


# --- MNA-601: hydrate / resume draft from archive KP (status «в работе» only) ---
#
# Assumed HTTP contract:
#   POST /api/v1/commercial/archive/{kp_id}/resume
#     → 200 CommercialDraftDetailsResponse
#        order_data + header from KP; saved_offer.kp_id == kp_id; resume_kp_id set
#     → 401 without session
#     → 404 if KP missing
#     → 409 or 400 if status ≠ «в работе»
#   Service method (ArchiveService): resume_as_draft(kp_id, *, user) -> draft details dict


def _fake_resume_draft_details(kp_id: int = 42) -> dict:
    """Minimal draft body matching CommercialDraftDetailsResponse for resume HTTP tests."""
    return {
        "draft_id": f"draft-resume-{kp_id}",
        "order": {},
        "optimization": {},
        "order_data": [
            {
                "line_id": "ln_plate_resume",
                "product_type": "plates",
                "name": "ПБ 78-12-8п",
                "qty": 2,
                "unit_price": 1000.0,
            },
            {
                "line_id": "ln_pile_resume",
                "product_type": "piles",
                "mark": "С120.35-12",
                "concrete_grade": "B25",
                "qty": 3,
                "unit_price": 40000.0,
            },
        ],
        "metadata": {
            "product_type": "plates",
            "client_name": "ООО Тест",
            "manager_name": "Иван Иванов",
            "discount_percent": 5.0,
            "logistics_cost": 0.0,
            "wide_plates_resolved": True,
            "current_step": "result",
            "resume_kp_id": kp_id,
            "append_batches": [],
        },
        "wizard_state": {
            "current_step": "result",
            "can_proceed_to": [],
            "next_required_action": "none",
            "validation_errors": [],
        },
        "files": [],
        "saved_offer": {
            "kp_id": kp_id,
            "status": "в работе",
            "mode": "database",
            "execution_terms": "",
            "saved_at": "2026-08-12T12:00:00",
        },
        "totals": {"subtotal": 1000.0, "vat_amount": 200.0, "total_with_vat": 1200.0},
        "offer_identity": {
            "offer_number": str(kp_id),
            "offer_date": "12.08.2026",
            "file_stem": f"kp_{kp_id}",
        },
    }


def test_resume_archive_kp_as_draft_requires_auth(client: TestClient) -> None:
    """MNA-601: POST .../archive/{kp_id}/resume requires session."""
    response = client.post("/api/v1/commercial/archive/42/resume")
    assert response.status_code == 401


def test_resume_archive_kp_as_draft_ok(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    """MNA-601: resume hydrates draft bound to same kp_id with order_data + header."""
    fake_service.resume_as_draft.return_value = _fake_resume_draft_details(42)

    response = client.post(
        "/api/v1/commercial/archive/42/resume",
        cookies=auth_cookie,
    )

    assert response.status_code == 200, response.text
    body = response.json()
    assert body["draft_id"]
    assert body["saved_offer"]["kp_id"] == 42
    assert body["saved_offer"]["status"] == "в работе"
    assert body["metadata"]["resume_kp_id"] == 42
    assert body["metadata"]["client_name"] == "ООО Тест"
    assert len(body["order_data"]) >= 2
    product_types = {line.get("product_type") for line in body["order_data"]}
    assert "plates" in product_types
    assert "piles" in product_types
    fake_service.resume_as_draft.assert_called_once_with(42, user=TESTER_USER)


def test_resume_archive_kp_as_draft_rejects_non_in_progress(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    """MNA-601 / R2: non-«в работе» → 409 Conflict or 400 Bad Request."""
    from app.services.archive_service import ArchiveValidationError

    fake_service.resume_as_draft.side_effect = ArchiveValidationError(
        "Дополнить КП можно только в статусе «в работе»."
    )

    response = client.post(
        "/api/v1/commercial/archive/42/resume",
        cookies=auth_cookie,
    )

    assert response.status_code in (400, 409), response.text
    fake_service.resume_as_draft.assert_called_once_with(42, user=TESTER_USER)


def test_resume_archive_kp_as_draft_not_found(
    client: TestClient,
    auth_cookie: dict[str, str],
    fake_service: MagicMock,
) -> None:
    """MNA-601: missing KP → 404."""
    from app.services.archive_service import ArchiveNotFoundError

    fake_service.resume_as_draft.side_effect = ArchiveNotFoundError("нет такого")

    response = client.post(
        "/api/v1/commercial/archive/999/resume",
        cookies=auth_cookie,
    )

    assert response.status_code == 404
    fake_service.resume_as_draft.assert_called_once_with(999, user=TESTER_USER)
