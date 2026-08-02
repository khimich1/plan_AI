"""SHIP-300: /api/v1/logistics endpoints — RBAC, CRUD-цикл, sheet.xlsx."""

from __future__ import annotations

import sqlite3

import pytest

from tests.helpers import kp_db_fixtures as fx

PLATE = "ПБ 60-12-8п"
API = "/api/v1/logistics"

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
        "id": 4,
        "username": "logist",
        "role": "logistics",
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


@pytest.fixture()
def logistics_db(tmp_path, monkeypatch: pytest.MonkeyPatch) -> str:
    from app.core.settings import get_settings
    from tests.helpers.auth_fixtures import patch_auth_users
    from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY

    db_path = fx.make_iso_db(tmp_path)
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    monkeypatch.setenv("PB_DB_PATH", db_path)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, TEST_USERS)
    return db_path


@pytest.fixture()
def client(logistics_db: str):
    from app.main import create_app
    from tests.helpers.csrf import CsrfAwareTestClient

    del logistics_db
    return CsrfAwareTestClient(create_app())


def _cookie(user_id: int, role: str, username: str) -> dict[str, str]:
    from tests.helpers.production_api_fixtures import session_cookie

    return session_cookie(user_id, role, username)


@pytest.fixture()
def admin_cookie() -> dict[str, str]:
    return _cookie(1, "admin", "admin")


@pytest.fixture()
def logistics_cookie() -> dict[str, str]:
    return _cookie(4, "logistics", "logist")


@pytest.fixture()
def manager_cookie() -> dict[str, str]:
    return _cookie(3, "manager", "manager_a")


def _seed_sgp(db_path: str, kp_id: int = 1, qty: int = 5) -> int:
    fx.seed_kp_offer(db_path, kp_id, customer_name="КлиентА")
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_meta SET ordered_qty = ? WHERE kp_id = ?", (qty, kp_id)
        )
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, plan_id
            ) VALUES (?, ?, 6.0, 1.2, 800, ?, '27.07.2026', 1, 'plan-1')
            """,
            (kp_id, PLATE, qty),
        )
        conn.commit()
        return int(cur.lastrowid)


# ---------------------------------------------------------------------------
# RBAC
# ---------------------------------------------------------------------------


def test_manager_forbidden_everywhere(client, manager_cookie) -> None:
    assert client.get(f"{API}/shipments", cookies=manager_cookie).status_code == 403
    assert (
        client.post(
            f"{API}/shipments",
            json={"shipment_date": "2026-07-31", "delivery_type": "delivery", "kp_ids": [1]},
            cookies=manager_cookie,
        ).status_code
        == 403
    )
    assert (
        client.post(
            f"{API}/shipments/1/reuse-transport",
            json={"shipment_date": "2026-07-31", "delivery_type": "delivery", "kp_ids": [1]},
            cookies=manager_cookie,
        ).status_code
        == 403
    )
    assert client.get(f"{API}/carriers", cookies=manager_cookie).status_code == 403
    assert (
        client.post(
            f"{API}/carriers/1/merge",
            json={"into_id": 2},
            cookies=manager_cookie,
        ).status_code
        == 403
    )
    assert client.get(f"{API}/pile-catalog", cookies=manager_cookie).status_code == 403


def test_unauthenticated_rejected(client) -> None:
    assert client.get(f"{API}/shipments").status_code == 401
    assert (
        client.post(
            f"{API}/shipments/1/reuse-transport",
            json={"shipment_date": "2026-07-31", "delivery_type": "delivery", "kp_ids": [1]},
        ).status_code
        == 401
    )


def test_logistics_role_allowed(client, logistics_db, logistics_cookie) -> None:
    fx.seed_kp_offer(logistics_db, 1)
    response = client.post(
        f"{API}/shipments",
        json={"shipment_date": "2026-07-31", "delivery_type": "delivery", "kp_ids": [1]},
        cookies=logistics_cookie,
    )
    assert response.status_code == 200, response.text
    assert response.json()["status"] == "in_work"


# ---------------------------------------------------------------------------
# Полный цикл
# ---------------------------------------------------------------------------


def test_shipment_full_cycle(client, logistics_db, admin_cookie) -> None:
    cp_id = _seed_sgp(logistics_db, 1, 5)

    created = client.post(
        f"{API}/shipments",
        json={"shipment_date": "2026-07-31", "delivery_type": "delivery", "kp_ids": [1]},
        cookies=admin_cookie,
    )
    assert created.status_code == 200, created.text
    shipment_id = created.json()["id"]

    listing = client.get(f"{API}/shipments?kp_id=1", cookies=admin_cookie)
    assert listing.status_code == 200
    assert [item["id"] for item in listing.json()["items"]] == [shipment_id]

    card = client.get(f"{API}/shipments/{shipment_id}", cookies=admin_cookie)
    assert card.status_code == 200
    body = card.json()
    assert body["available_by_kp"][0]["plates"][0]["available_qty"] == 5

    propose = client.post(
        f"{API}/shipments/{shipment_id}/propose?vehicle_class=t20",
        cookies=admin_cookie,
    )
    assert propose.status_code == 200, propose.text
    proposed = propose.json()
    assert proposed["vehicle_class_limits_kg"]["t20"] == 19800
    assert [(item["completed_plate_id"], item["qty"]) for item in proposed["items"]] == [
        (cp_id, 5)
    ]
    layout = proposed["layout"]
    assert layout is not None
    assert layout["body_length_m"] == pytest.approx(13.2)
    assert len(layout["stacks"]) == 1
    assert layout["stacks"][0]["tiers"][0]["units"][0]["completed_plate_id"] == cp_id
    assert layout["loading_steps"][0]["step"] == 1

    put = client.put(
        f"{API}/shipments/{shipment_id}/items",
        json={
            "items": [
                {"item_type": "plate", "completed_plate_id": cp_id, "qty": 5},
                {"item_type": "free", "mark": "С60.30", "qty": 1, "weight_kg": 1060.0},
            ]
        },
        cookies=admin_cookie,
    )
    assert put.status_code == 200, put.text
    assert len(put.json()["items"]) == 2

    # Без ЯР-заказа complete → 422 shipment_missing_ya_order.
    missing = client.post(
        f"{API}/shipments/{shipment_id}/complete", cookies=admin_cookie
    )
    assert missing.status_code == 422
    assert missing.json()["detail"]["code"] == "shipment_missing_ya_order"

    patched = client.patch(
        f"{API}/shipments/{shipment_id}",
        json={"orders": [{"kp_id": 1, "ya_order_no": "ЯР-1"}]},
        cookies=admin_cookie,
    )
    assert patched.status_code == 200
    # PATCH возвращает полную карточку той же формы, что и GET.
    refetched = client.get(f"{API}/shipments/{shipment_id}", cookies=admin_cookie)
    assert set(patched.json()) == set(refetched.json())
    assert patched.json()["orders"] == refetched.json()["orders"]
    assert "available_by_kp" in patched.json()

    completed = client.post(
        f"{API}/shipments/{shipment_id}/complete", cookies=admin_cookie
    )
    assert completed.status_code == 200, completed.text
    assert completed.json()["status"] == "done"

    # Повторные мутации по done-рейсу → 422 shipment_not_in_work.
    again = client.post(f"{API}/shipments/{shipment_id}/complete", cookies=admin_cookie)
    assert again.status_code == 422
    assert again.json()["detail"]["code"] == "shipment_not_in_work"
    cancelled = client.post(
        f"{API}/shipments/{shipment_id}/cancel", cookies=admin_cookie
    )
    assert cancelled.status_code == 422

    sheet = client.get(
        f"{API}/shipments/{shipment_id}/sheet.xlsx", cookies=admin_cookie
    )
    assert sheet.status_code == 200
    assert sheet.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    disposition = sheet.headers["content-disposition"]
    assert "attachment" in disposition
    assert f'filename="shipment_{shipment_id}_sheet.xlsx"' in disposition
    assert len(sheet.content) > 1000


def test_shipment_cancel_via_api(client, logistics_db, admin_cookie) -> None:
    fx.seed_kp_offer(logistics_db, 1)
    created = client.post(
        f"{API}/shipments",
        json={"shipment_date": "2026-07-31", "delivery_type": "pickup", "kp_ids": [1]},
        cookies=admin_cookie,
    )
    shipment_id = created.json()["id"]
    cancelled = client.post(
        f"{API}/shipments/{shipment_id}/cancel", cookies=admin_cookie
    )
    assert cancelled.status_code == 200
    assert client.get(
        f"{API}/shipments/{shipment_id}", cookies=admin_cookie
    ).status_code == 404


def test_create_without_kp_422(client, admin_cookie) -> None:
    response = client.post(
        f"{API}/shipments",
        json={"shipment_date": "2026-07-31", "delivery_type": "delivery", "kp_ids": []},
        cookies=admin_cookie,
    )
    assert response.status_code == 422


# ---------------------------------------------------------------------------
# Переиспользование транспорта (REUSE-200)
# ---------------------------------------------------------------------------


def test_reuse_transport_via_api(client, logistics_db, logistics_cookie) -> None:
    fx.seed_kp_offer(logistics_db, 1)
    fx.seed_kp_offer(logistics_db, 2)
    with sqlite3.connect(logistics_db) as conn:
        cur = conn.execute(
            "INSERT INTO carriers (name, name_normalized) VALUES ('ООО «Альфа»', 'альфа')"
        )
        carrier_id = int(cur.lastrowid)
        conn.commit()

    source = client.post(
        f"{API}/shipments",
        json={"shipment_date": "2026-07-31", "delivery_type": "delivery", "kp_ids": [1]},
        cookies=logistics_cookie,
    )
    assert source.status_code == 200, source.text
    source_id = source.json()["id"]
    patched = client.patch(
        f"{API}/shipments/{source_id}",
        json={
            "carrier_id": carrier_id,
            "driver_name": "Пётр",
            "vehicle_text": "MAN А123ВС77",
            "vehicle_class": "t20",
            "proxy_no": "дов-9",
            "upd_no": "УПД-77",
            "freight_request_no": "З-100",
            "planned_cost": 45000,
        },
        cookies=logistics_cookie,
    )
    assert patched.status_code == 200, patched.text

    reused = client.post(
        f"{API}/shipments/{source_id}/reuse-transport",
        json={"shipment_date": "2026-08-02", "delivery_type": "pickup", "kp_ids": [2]},
        cookies=logistics_cookie,
    )
    assert reused.status_code == 200, reused.text
    body = reused.json()
    assert body["id"] != source_id
    assert body["status"] == "in_work"
    assert body["carrier_id"] == carrier_id
    assert body["driver_name"] == "Пётр"
    assert body["vehicle_text"] == "MAN А123ВС77"
    assert body["vehicle_class"] == "t20"
    assert body["proxy_no"] == "дов-9"
    assert body["upd_no"] is None
    assert body["freight_request_no"] is None
    assert body["planned_cost"] is None
    assert body["items"] == []
    assert [order["kp_id"] for order in body["orders"]] == [2]


def test_reuse_transport_source_not_found_404(client, logistics_db, admin_cookie) -> None:
    fx.seed_kp_offer(logistics_db, 1)
    response = client.post(
        f"{API}/shipments/9999/reuse-transport",
        json={"shipment_date": "2026-08-02", "delivery_type": "delivery", "kp_ids": [1]},
        cookies=admin_cookie,
    )
    assert response.status_code == 404
    assert response.json()["detail"]["code"] == "shipment_not_found"


def test_reuse_transport_without_kp_422(client, logistics_db, admin_cookie) -> None:
    fx.seed_kp_offer(logistics_db, 1)
    source_id = client.post(
        f"{API}/shipments",
        json={"shipment_date": "2026-07-31", "delivery_type": "delivery", "kp_ids": [1]},
        cookies=admin_cookie,
    ).json()["id"]
    response = client.post(
        f"{API}/shipments/{source_id}/reuse-transport",
        json={"shipment_date": "2026-08-02", "delivery_type": "delivery", "kp_ids": []},
        cookies=admin_cookie,
    )
    assert response.status_code == 422


def test_list_shipments_filters_via_query_string(client, logistics_db, admin_cookie) -> None:
    """SC-1: фильтры реестра проброшены из query-string в выборку (date_to, carrier, «без УПД»)."""
    fx.seed_kp_offer(logistics_db, 1)
    with sqlite3.connect(logistics_db) as conn:
        cur = conn.execute(
            "INSERT INTO carriers (name, name_normalized) VALUES ('ООО «Альфа»', 'альфа')"
        )
        carrier_id = int(cur.lastrowid)
        conn.commit()

    july_id = client.post(
        f"{API}/shipments",
        json={"shipment_date": "2026-07-31", "delivery_type": "delivery", "kp_ids": [1]},
        cookies=admin_cookie,
    ).json()["id"]
    august_id = client.post(
        f"{API}/shipments",
        json={"shipment_date": "2026-08-15", "delivery_type": "pickup", "kp_ids": [1]},
        cookies=admin_cookie,
    ).json()["id"]
    patched = client.patch(
        f"{API}/shipments/{july_id}",
        json={"carrier_id": carrier_id, "upd_no": "101"},
        cookies=admin_cookie,
    )
    assert patched.status_code == 200

    by_date_to = client.get(f"{API}/shipments?date_to=2026-08-01", cookies=admin_cookie)
    assert [item["id"] for item in by_date_to.json()["items"]] == [july_id]

    by_carrier = client.get(f"{API}/shipments?carrier_id={carrier_id}", cookies=admin_cookie)
    assert [item["id"] for item in by_carrier.json()["items"]] == [july_id]
    assert by_carrier.json()["items"][0]["carrier_name"] == 'ООО «Альфа»'

    no_upd = client.get(f"{API}/shipments?no_upd=1", cookies=admin_cookie)
    assert [item["id"] for item in no_upd.json()["items"]] == [august_id]


# ---------------------------------------------------------------------------
# Справочники
# ---------------------------------------------------------------------------


def test_carriers_endpoints(client, logistics_db, admin_cookie) -> None:
    with sqlite3.connect(logistics_db) as conn:
        conn.execute(
            "INSERT INTO carriers (name, name_normalized, source_sheet) "
            "VALUES ('ООО «Альфа»', 'альфа', 'Перевозчики')"
        )
        conn.execute(
            "INSERT INTO carriers (name, name_normalized, source_sheet) "
            "VALUES ('ООО «Альфа Транс»', 'альфа транс', 'Перевозчики')"
        )
        conn.commit()

    listed = client.get(f"{API}/carriers?q=альфа", cookies=admin_cookie)
    assert listed.status_code == 200
    assert listed.json()["count"] == 2
    assert {item["source_sheet"] for item in listed.json()["items"]} == {"Перевозчики"}

    merged = client.post(
        f"{API}/carriers/1/merge", json={"into_id": 2}, cookies=admin_cookie
    )
    assert merged.status_code == 200, merged.text
    assert merged.json()["moved_shipments"] == 0

    conflict = client.post(
        f"{API}/carriers/2/merge", json={"into_id": 1}, cookies=admin_cookie
    )
    # 1 уже слит (неактивен) → carrier_merge_conflict.
    assert conflict.status_code == 422
    assert conflict.json()["detail"]["code"] == "carrier_merge_conflict"


def test_archive_search_forbids_logistics_admin_still_works(
    client, logistics_db, logistics_cookie, manager_cookie, admin_cookie
) -> None:
    fx.seed_kp_offer(logistics_db, 1, customer_name="КлиентА")
    with sqlite3.connect(logistics_db) as conn:
        conn.execute(
            """
            UPDATE KP_offers
            SET discount_percent = 10, subtotal = 1000, vat_amount = 200, total_amount = 1200
            WHERE kp_id = 1
            """
        )
        conn.commit()

    # Logistics больше не имеет доступа к commercial archive search.
    assert (
        client.get(
            "/api/v1/commercial/archive/search?kp_id=1", cookies=logistics_cookie
        ).status_code
        == 403
    )

    # Admin видит полный commercial aggregate (включая финансы).
    admin_search = client.get(
        "/api/v1/commercial/archive/search?kp_id=1", cookies=admin_cookie
    )
    assert admin_search.status_code == 200, admin_search.text
    admin_item = admin_search.json()["items"][0]
    assert admin_item["subtotal"] == 1000
    assert admin_item["total_amount"] == 1200
    assert admin_item["discount_percent"] == 10

    listing = client.get("/api/v1/commercial/archive", cookies=logistics_cookie)
    assert listing.status_code == 403
    details = client.get("/api/v1/commercial/archive/1", cookies=logistics_cookie)
    assert details.status_code == 403

    # Manager на чужом КП (без owner_user_id) — 403, как и раньше.
    assert (
        client.get(
            "/api/v1/commercial/archive/search?kp_id=1", cookies=manager_cookie
        ).status_code
        == 403
    )


def test_logistics_kp_search_slim_fields_and_acl_b(
    client, logistics_db, logistics_cookie
) -> None:
    fx.seed_kp_offer(logistics_db, 10, customer_name="Ромашка", status="в работе")
    fx.seed_kp_offer(logistics_db, 11, customer_name="Ромашка Архив", status="в архиве")
    fx.seed_kp_offer(logistics_db, 12, customer_name="Ромашка Готово", status="выполнено")
    fx.seed_kp_offer(logistics_db, 13, customer_name="СГП Клиент", status="На СГП")
    with sqlite3.connect(logistics_db) as conn:
        conn.execute(
            """
            UPDATE KP_offers
            SET discount_percent = 15, subtotal = 5000, vat_amount = 1000, total_amount = 6000
            WHERE kp_id = 10
            """
        )
        conn.commit()

    by_id = client.get(f"{API}/kp-search?kp_id=10", cookies=logistics_cookie)
    assert by_id.status_code == 200, by_id.text
    body = by_id.json()
    assert body["mode"] == "number"
    assert body["total"] == 1
    item = body["items"][0]
    assert item == {
        "kp_id": 10,
        "customer_name": "Ромашка",
        "status": "в работе",
        "product_type": "plates",
    }
    for key in ("discount_percent", "subtotal", "vat_amount", "total_amount", "finance"):
        assert key not in item

    archived = client.get(f"{API}/kp-search?kp_id=11", cookies=logistics_cookie)
    assert archived.status_code == 200
    assert archived.json()["items"] == []
    assert archived.json()["total"] == 0

    done = client.get(f"{API}/kp-search?kp_id=12", cookies=logistics_cookie)
    assert done.status_code == 200
    assert done.json()["items"] == []

    on_sgp = client.get(f"{API}/kp-search?kp_id=13", cookies=logistics_cookie)
    assert on_sgp.status_code == 200
    assert on_sgp.json()["items"][0]["status"] == "На СГП"

    by_customer = client.get(
        f"{API}/kp-search?customer=Ромашка", cookies=logistics_cookie
    )
    assert by_customer.status_code == 200, by_customer.text
    ids = {row["kp_id"] for row in by_customer.json()["items"]}
    assert ids == {10}
    assert 11 not in ids and 12 not in ids


def test_pile_catalog_endpoint(client, logistics_db, admin_cookie) -> None:
    with sqlite3.connect(logistics_db) as conn:
        conn.execute(
            "INSERT INTO pile_catalog (mark, length_m, section_mm, volume_m3, weight_kg, pcs_per_20t) "
            "VALUES ('С60.30', 6.0, 300, 0.216, 1060.0, 18)"
        )
        conn.execute(
            "INSERT INTO pile_catalog (mark, length_m, section_mm, volume_m3, weight_kg, pcs_per_20t) "
            "VALUES ('С137,5.40', 13.75, 400, 0.55, 2750.0, NULL)"
        )
        conn.commit()

    response = client.get(f"{API}/pile-catalog?q=С137", cookies=admin_cookie)
    assert response.status_code == 200
    body = response.json()
    assert body["count"] == 1
    assert body["items"][0]["mark"] == "С137,5.40"
    assert body["items"][0]["weight_kg"] == 2750.0

    empty = client.get(f"{API}/pile-catalog", cookies=admin_cookie)
    assert empty.json()["count"] == 2
