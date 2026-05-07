from __future__ import annotations

from fastapi import Depends, HTTPException, status

from app.dependencies.auth import REQUIRE_ADMIN_OR_MANAGER
from app.services.draft_store import DraftStore, UnsafeDraftIdError


def check_draft_ownership(draft_id: str, user: dict) -> None:
    """Ensure the draft exists and was created by the current user."""
    store = DraftStore()
    try:
        raw = store.load_raw_json(draft_id)
    except UnsafeDraftIdError:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from None
    if raw is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.")
    meta = raw.get("metadata") or {}
    owner = meta.get("owner_user_id")
    if owner is None:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Черновик недоступен. Создайте новый черновик.",
        )
    if int(owner) != int(user["id"]):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Доступ к черновику запрещён.",
        )


def verify_draft_ownership(
    draft_id: str,
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> str:
    check_draft_ownership(draft_id, user)
    return draft_id
