"""Task T5: Registry API — CRUD справочников ГСМ + settings.

Acceptance: GET/POST/PATCH vehicles|drivers|cards|stations;
PATCH cards — bind / archive (archived_at, не DELETE);
GET/PUT settings; norms>0, tank>0, card_number unique;
REQUIRE_ACCOUNTING; архивные карты вне default list, транзакции видны.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.gsm_repository import GsmRepository
from core import kp_db_schema
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

from app.services.gsm_registry_service import GsmRegistryService  # noqa: E402

PREFIX = "/api/v1/gsm"

TEST_USERS = [
    {
        "id": 1,
        "username": "admin",
        "role": "admin",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 5,
        "username": "accountant_user",
        "role": "accountant",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 3,
        "username": "manager_a",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
]


# =============================================================================
# Fixtures
# =============================================================================


def _fresh_db(tmp_path: Path, name: str = "gsm_reg.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


@pytest.fixture()
def db_path(tmp_path: Path) -> str:
    return _fresh_db(tmp_path)


@pytest.fixture()
def repo(db_path: str) -> GsmRepository:
    return GsmRepository(db_path=db_path)


@pytest.fixture()
def service(repo: GsmRepository) -> GsmRegistryService:
    return GsmRegistryService(repo=repo)


@pytest.fixture()
def api_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> CsrfAwareTestClient:
    db = tmp_path / "plita.db"
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", str(db))
    monkeypatch.setenv("PB_DB_PATH", str(db))
    get_settings.cache_clear()
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(str(db))
    patch_auth_users(monkeypatch, TEST_USERS)
    return CsrfAwareTestClient(create_app())


def _auth(username: str = "accountant_user") -> dict[str, str]:
    by_user = {
        "admin": (1, "admin"),
        "accountant_user": (5, "accountant"),
        "manager_a": (3, "manager"),
    }
    user_id, role = by_user[username]
    return session_cookie(user_id, role, username)


# =============================================================================
# Service-level: vehicles / drivers / cards / stations / settings
# =============================================================================


def test_service_create_list_patch_vehicle(service: GsmRegistryService) -> None:
    created = service.create_vehicle(
        name="Geely Tugella",
        plate_number="О 848 ХР 44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
    )
    assert created.id >= 1
    assert created.tank_volume_liters == pytest.approx(55.0)
    assert created.is_active is True

    listed = service.list_vehicles()
    assert len(listed) == 1
    assert listed[0].id == created.id

    patched = service.patch_vehicle(created.id, norm_summer=9.8, name="Geely 848")
    assert patched.norm_summer == pytest.approx(9.8)
    assert patched.name == "Geely 848"


def test_service_rejects_non_positive_tank_and_norms(service: GsmRegistryService) -> None:
    with pytest.raises(Exception) as exc_info:
        service.create_vehicle(
            name="Bad",
            plate_number="X",
            tank_volume_liters=0,
            norm_summer=9.4,
            norm_winter=10.3,
        )
    assert "tank" in str(exc_info.value).lower() or getattr(exc_info.value, "code", "") == "gsm_validation"

    with pytest.raises(Exception):
        service.create_vehicle(
            name="Bad",
            plate_number="Y",
            tank_volume_liters=55,
            norm_summer=-1,
            norm_winter=10.3,
        )


def test_service_create_list_patch_driver(service: GsmRegistryService) -> None:
    created = service.create_driver(
        full_name="Кулигин Никита Валерьевич",
        license_number="44 21 846315",
        license_issued_at="30.07.2015",
        personnel_number="143",
        snils="123-456-789 00",
    )
    assert created.full_name.startswith("Кулигин")
    assert len(service.list_drivers()) == 1

    patched = service.patch_driver(created.id, personnel_number="144", is_active=False)
    assert patched.personnel_number == "144"
    assert patched.is_active is False
    assert service.list_drivers() == []
    assert len(service.list_drivers(active_only=False)) == 1


def test_service_card_create_bind_archive_unique(
    service: GsmRegistryService, repo: GsmRepository
) -> None:
    v1 = service.create_vehicle(
        name="Car1",
        plate_number="A111",
        tank_volume_liters=55,
        norm_summer=9.4,
        norm_winter=10.3,
    )
    v2 = service.create_vehicle(
        name="Car2",
        plate_number="B222",
        tank_volume_liters=60,
        norm_summer=10.0,
        norm_winter=11.0,
    )
    card = service.create_card(card_number="3005454268", vehicle_id=v1.id)
    assert card.vehicle_id == v1.id
    assert card.archived_at is None

    with pytest.raises(Exception) as dup:
        service.create_card(card_number="3005454268", vehicle_id=v2.id)
    assert "unique" in str(dup.value).lower() or getattr(dup.value, "code", "") == "gsm_card_duplicate"

    bound = service.patch_card(card.id, vehicle_id=v2.id)
    assert bound.vehicle_id == v2.id

    archived = service.patch_card(card.id, archive=True)
    assert archived.archived_at is not None
    assert archived.archived_at != ""

    # Row still in DB
    assert repo.get_card(card.id) is not None
    assert len(service.list_cards()) == 0
    assert len(service.list_cards(include_archived=True)) == 1


def test_archived_card_transactions_still_listed(
    service: GsmRegistryService, repo: GsmRepository
) -> None:
    """Архивная карта не в default list, но транзакции по vehicle остаются видны."""
    vehicle = service.create_vehicle(
        name="Car",
        plate_number="C333",
        tank_volume_liters=55,
        norm_summer=9.4,
        norm_winter=10.3,
    )
    card = service.create_card(card_number="1112223334", vehicle_id=vehicle.id)
    batch_id = repo.create_import_batch(
        filename="t.xls",
        uploaded_at="2026-08-14T12:00:00",
    )
    repo.insert_transaction(
        card_id=card.id,
        ts="2025-04-03T10:00:00",
        service_type="fuel",
        qty_liters=40.0,
        amount=2500.0,
        raw_address="АЗС 1",
        batch_id=batch_id,
    )
    service.patch_card(card.id, archive=True)

    assert service.list_cards() == []
    rows = repo.list_transactions(
        vehicle_id=vehicle.id,
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
    )
    assert len(rows) == 1
    assert rows[0]["card_id"] == card.id


def test_service_stations_and_settings(service: GsmRegistryService) -> None:
    st = service.create_station(
        address="г. Кострома, ул. Примерная, 1",
        brand="TATNEFT",
        lat=57.76,
        lon=40.92,
    )
    assert st.address.startswith("г. Кострома")
    patched = service.patch_station(st.id, brand="Газпромнефть")
    assert patched.brand == "Газпромнефть"
    assert len(service.list_stations()) == 1

    defaults = service.get_settings()
    assert defaults.winter_start == "11-01"
    assert defaults.hook_threshold_km == pytest.approx(13.0)
    assert defaults.max_daily_km == 700

    updated = service.put_settings(
        winter_start="10-15", hook_threshold_km=15, max_daily_km=400
    )
    assert updated.winter_start == "10-15"
    assert updated.hook_threshold_km == pytest.approx(15.0)
    assert updated.max_daily_km == 400
    assert service.get_settings().hook_threshold_km == pytest.approx(15.0)
    assert service.get_settings().max_daily_km == 400

    with pytest.raises(Exception) as exc_info:
        service.put_settings(winter_start="11-01", hook_threshold_km=13, max_daily_km=0)
    assert getattr(exc_info.value, "code", "") == "gsm_validation"


# =============================================================================
# HTTP endpoints — AuthZ + CRUD smoke
# =============================================================================


def test_registry_endpoints_forbid_manager(api_client: CsrfAwareTestClient) -> None:
    cookie = _auth("manager_a")
    checks: list[tuple[str, str, dict | None]] = [
        ("get", f"{PREFIX}/vehicles", None),
        (
            "post",
            f"{PREFIX}/vehicles",
            {
                "name": "X",
                "plate_number": "Y",
                "tank_volume_liters": 55,
                "norm_summer": 9,
                "norm_winter": 10,
            },
        ),
        ("get", f"{PREFIX}/drivers", None),
        ("get", f"{PREFIX}/cards", None),
        ("get", f"{PREFIX}/stations", None),
        ("get", f"{PREFIX}/settings", None),
        (
            "put",
            f"{PREFIX}/settings",
            {"winter_start": "11-01", "hook_threshold_km": 13},
        ),
    ]
    for method, path, body in checks:
        kwargs: dict = {"cookies": cookie}
        if body is not None:
            kwargs["json"] = body
        resp = getattr(api_client, method)(path, **kwargs)
        assert resp.status_code == 403, f"{method.upper()} {path} → {resp.status_code}"


def test_http_vehicles_crud(api_client: CsrfAwareTestClient) -> None:
    cookie = _auth()
    create = api_client.post(
        f"{PREFIX}/vehicles",
        cookies=cookie,
        json={
            "name": "Monjaro",
            "plate_number": "О 001 АА 44",
            "tank_volume_liters": 70,
            "norm_summer": 9.5,
            "norm_winter": 10.5,
        },
    )
    assert create.status_code in (200, 201)
    body = create.json()
    vid = body["id"]
    assert body["name"] == "Monjaro"

    listed = api_client.get(f"{PREFIX}/vehicles", cookies=cookie)
    assert listed.status_code == 200
    items = listed.json()
    if isinstance(items, dict):
        items = items["items"]
    assert any(v["id"] == vid for v in items)

    patched = api_client.patch(
        f"{PREFIX}/vehicles/{vid}",
        cookies=cookie,
        json={"tank_volume_liters": 72},
    )
    assert patched.status_code == 200
    assert patched.json()["tank_volume_liters"] == pytest.approx(72)

    bad = api_client.post(
        f"{PREFIX}/vehicles",
        cookies=cookie,
        json={
            "name": "Bad",
            "plate_number": "Z",
            "tank_volume_liters": 0,
            "norm_summer": 9.5,
            "norm_winter": 10.5,
        },
    )
    assert bad.status_code == 422


def test_http_drivers_crud(api_client: CsrfAwareTestClient) -> None:
    cookie = _auth()
    create = api_client.post(
        f"{PREFIX}/drivers",
        cookies=cookie,
        json={
            "full_name": "Иванов Иван Иванович",
            "license_number": "11 22 333333",
        },
    )
    assert create.status_code in (200, 201)
    did = create.json()["id"]

    listed = api_client.get(f"{PREFIX}/drivers", cookies=cookie)
    assert listed.status_code == 200

    patched = api_client.patch(
        f"{PREFIX}/drivers/{did}",
        cookies=cookie,
        json={"snils": "000-000-000 00"},
    )
    assert patched.status_code == 200
    assert patched.json()["snils"] == "000-000-000 00"


def test_http_cards_bind_archive_and_unique(api_client: CsrfAwareTestClient) -> None:
    cookie = _auth()
    v1 = api_client.post(
        f"{PREFIX}/vehicles",
        cookies=cookie,
        json={
            "name": "V1",
            "plate_number": "A1",
            "tank_volume_liters": 55,
            "norm_summer": 9,
            "norm_winter": 10,
        },
    ).json()
    v2 = api_client.post(
        f"{PREFIX}/vehicles",
        cookies=cookie,
        json={
            "name": "V2",
            "plate_number": "B2",
            "tank_volume_liters": 55,
            "norm_summer": 9,
            "norm_winter": 10,
        },
    ).json()

    create = api_client.post(
        f"{PREFIX}/cards",
        cookies=cookie,
        json={"card_number": "3005459999", "vehicle_id": v1["id"]},
    )
    assert create.status_code in (200, 201)
    cid = create.json()["id"]

    dup = api_client.post(
        f"{PREFIX}/cards",
        cookies=cookie,
        json={"card_number": "3005459999", "vehicle_id": v2["id"]},
    )
    assert dup.status_code == 422

    bind = api_client.patch(
        f"{PREFIX}/cards/{cid}",
        cookies=cookie,
        json={"vehicle_id": v2["id"]},
    )
    assert bind.status_code == 200
    assert bind.json()["vehicle_id"] == v2["id"]

    archive = api_client.patch(
        f"{PREFIX}/cards/{cid}",
        cookies=cookie,
        json={"archive": True},
    )
    assert archive.status_code == 200
    assert archive.json()["archived_at"] is not None

    # Default list hides archived
    listed = api_client.get(f"{PREFIX}/cards", cookies=cookie)
    assert listed.status_code == 200
    items = listed.json()
    if isinstance(items, dict):
        items = items["items"]
    assert items == []

    # include_archived shows it; row not deleted
    with_arch = api_client.get(
        f"{PREFIX}/cards",
        cookies=cookie,
        params={"include_archived": "true"},
    )
    assert with_arch.status_code == 200
    items_a = with_arch.json()
    if isinstance(items_a, dict):
        items_a = items_a["items"]
    assert len(items_a) == 1
    assert items_a[0]["id"] == cid


def test_http_stations_and_settings(api_client: CsrfAwareTestClient) -> None:
    cookie = _auth()
    create = api_client.post(
        f"{PREFIX}/stations",
        cookies=cookie,
        json={"address": "ул. Тестовая, 1", "brand": "ТНК"},
    )
    assert create.status_code in (200, 201)
    sid = create.json()["id"]

    patched = api_client.patch(
        f"{PREFIX}/stations/{sid}",
        cookies=cookie,
        json={"lat": 57.7, "lon": 40.9, "geocode_source": "manual"},
    )
    assert patched.status_code == 200
    assert patched.json()["lat"] == pytest.approx(57.7)

    settings = api_client.get(f"{PREFIX}/settings", cookies=cookie)
    assert settings.status_code == 200
    body = settings.json()
    assert "winter_start" in body
    assert "hook_threshold_km" in body
    assert body["max_daily_km"] == 700

    put = api_client.put(
        f"{PREFIX}/settings",
        cookies=cookie,
        json={"winter_start": "12-01", "hook_threshold_km": 20, "max_daily_km": 500},
    )
    assert put.status_code == 200
    assert put.json()["winter_start"] == "12-01"
    assert put.json()["hook_threshold_km"] == pytest.approx(20)
    assert put.json()["max_daily_km"] == 500

    again = api_client.get(f"{PREFIX}/settings", cookies=cookie)
    assert again.json()["hook_threshold_km"] == pytest.approx(20)
    assert again.json()["max_daily_km"] == 500

    rejected = api_client.put(
        f"{PREFIX}/settings",
        cookies=cookie,
        json={"winter_start": "12-01", "hook_threshold_km": 20, "max_daily_km": 0},
    )
    assert rejected.status_code == 422


def test_http_admin_allowed(api_client: CsrfAwareTestClient) -> None:
    cookie = _auth("admin")
    resp = api_client.get(f"{PREFIX}/vehicles", cookies=cookie)
    assert resp.status_code == 200


def test_http_patch_unknown_returns_404(api_client: CsrfAwareTestClient) -> None:
    cookie = _auth()
    resp = api_client.patch(
        f"{PREFIX}/vehicles/99999",
        cookies=cookie,
        json={"name": "Nope"},
    )
    assert resp.status_code == 404
