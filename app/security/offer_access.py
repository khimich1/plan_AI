from __future__ import annotations

from fastapi import HTTPException, status

FORBIDDEN_OFFER_DETAIL = "Доступ к КП запрещён."


def is_admin(user: dict) -> bool:
    return user.get("role") == "admin"


def is_admin_or_manager(user: dict) -> bool:
    return user.get("role") in {"admin", "manager"}


def can_read_offer(user: dict, _offer: dict) -> bool:
    # Managers share full commercial archive access with admins (all KP, all sections).
    return is_admin_or_manager(user)


def can_write_offer(user: dict, _offer: dict) -> bool:
    return is_admin_or_manager(user)


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
    """Query-level filters for offer list/search (admin/manager → no filter)."""
    if is_admin_or_manager(user):
        return {}
    return {"deny_all": True}
