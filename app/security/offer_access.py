from __future__ import annotations

from fastapi import HTTPException, status

FORBIDDEN_OFFER_DETAIL = "Доступ к КП запрещён."

_PRODUCTION_READ_STATUSES = frozenset({"в работе", "выполнено"})


def is_admin(user: dict) -> bool:
    return user.get("role") == "admin"


def get_offer_owner_id(offer: dict) -> int | None:
    raw = offer.get("owner_user_id")
    if raw is None:
        return None
    return int(raw)


def can_read_offer(user: dict, offer: dict) -> bool:
    if is_admin(user):
        return True
    role = user.get("role")
    if role == "manager":
        owner = get_offer_owner_id(offer)
        if owner is None:
            return False
        return owner == int(user["id"])
    if role == "production":
        status_value = offer.get("status") or "в работе"
        return status_value in _PRODUCTION_READ_STATUSES
    return False


def can_write_offer(user: dict, offer: dict) -> bool:
    if is_admin(user):
        return True
    if user.get("role") != "manager":
        return False
    owner = get_offer_owner_id(offer)
    if owner is None:
        return False
    return owner == int(user["id"])


def assert_offer_read_access(user: dict, offer: dict) -> None:
    if not can_read_offer(user, offer):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail=FORBIDDEN_OFFER_DETAIL,
        )


def assert_offer_write_access(user: dict, offer: dict) -> None:
    if not can_write_offer(user, offer):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail=FORBIDDEN_OFFER_DETAIL,
        )


def list_filters_for_user(user: dict) -> dict:
    """Query-level filters for offer list/search (admin → no filter)."""
    if is_admin(user):
        return {}
    role = user.get("role")
    if role == "manager":
        return {"owner_user_id": int(user["id"])}
    if role == "production":
        return {"readable_statuses": tuple(_PRODUCTION_READ_STATUSES)}
    return {"deny_all": True}
