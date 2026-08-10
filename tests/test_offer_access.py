"""Unit tests for commercial offer ACL (admin/manager full archive access)."""

from __future__ import annotations

import pytest
from fastapi import HTTPException

from app.security.offer_access import (
    FORBIDDEN_OFFER_DETAIL,
    assert_offer_read_access,
    assert_offer_write_access,
    can_read_offer,
    can_write_offer,
    list_filters_for_user,
)

ADMIN = {"id": 1, "role": "admin", "username": "admin"}
MANAGER = {"id": 2, "role": "manager", "username": "manager_a"}
OTHER_MANAGER = {"id": 3, "role": "manager", "username": "manager_b"}
PRODUCTION = {"id": 4, "role": "production", "username": "prod"}
LOGISTICS = {"id": 5, "role": "logistics", "username": "logist"}

OWN_OFFER = {"kp_id": 10, "owner_user_id": 2, "status": "в архиве"}
FOREIGN_OFFER = {"kp_id": 11, "owner_user_id": 3, "status": "в архиве"}
UNOWNED_OFFER = {"kp_id": 12, "owner_user_id": None, "status": "в работе"}
IN_PRODUCTION_FOREIGN = {"kp_id": 13, "owner_user_id": 3, "status": "в работе"}


@pytest.mark.parametrize(
    "user,offer",
    [
        (ADMIN, OWN_OFFER),
        (ADMIN, FOREIGN_OFFER),
        (ADMIN, UNOWNED_OFFER),
        (ADMIN, IN_PRODUCTION_FOREIGN),
        (MANAGER, OWN_OFFER),
        (MANAGER, FOREIGN_OFFER),
        (MANAGER, UNOWNED_OFFER),
        (MANAGER, IN_PRODUCTION_FOREIGN),
        (OTHER_MANAGER, OWN_OFFER),
    ],
)
def test_admin_and_manager_can_read_any_offer(user: dict, offer: dict) -> None:
    assert can_read_offer(user, offer) is True


@pytest.mark.parametrize(
    "user,offer",
    [
        (ADMIN, FOREIGN_OFFER),
        (MANAGER, FOREIGN_OFFER),
        (MANAGER, UNOWNED_OFFER),
        (MANAGER, IN_PRODUCTION_FOREIGN),
    ],
)
def test_admin_and_manager_can_write_any_offer(user: dict, offer: dict) -> None:
    assert can_write_offer(user, offer) is True


@pytest.mark.parametrize("user", [PRODUCTION, LOGISTICS])
@pytest.mark.parametrize("offer", [OWN_OFFER, FOREIGN_OFFER, UNOWNED_OFFER])
def test_non_commercial_roles_cannot_read_or_write(user: dict, offer: dict) -> None:
    assert can_read_offer(user, offer) is False
    assert can_write_offer(user, offer) is False


def test_list_filters_empty_for_admin_and_manager() -> None:
    assert list_filters_for_user(ADMIN) == {}
    assert list_filters_for_user(MANAGER) == {}


def test_list_filters_deny_all_for_other_roles() -> None:
    assert list_filters_for_user(PRODUCTION) == {"deny_all": True}
    assert list_filters_for_user(LOGISTICS) == {"deny_all": True}


def test_assert_offer_read_access_allows_manager_on_foreign_offer() -> None:
    assert_offer_read_access(MANAGER, FOREIGN_OFFER)


def test_assert_offer_write_access_allows_manager_on_foreign_offer() -> None:
    assert_offer_write_access(MANAGER, FOREIGN_OFFER)


def test_assert_offer_read_access_forbidden_for_production() -> None:
    with pytest.raises(HTTPException) as exc_info:
        assert_offer_read_access(PRODUCTION, OWN_OFFER)
    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == FORBIDDEN_OFFER_DETAIL


def test_assert_offer_write_access_forbidden_for_production() -> None:
    with pytest.raises(HTTPException) as exc_info:
        assert_offer_write_access(PRODUCTION, OWN_OFFER)
    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == FORBIDDEN_OFFER_DETAIL
