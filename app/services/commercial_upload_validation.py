from __future__ import annotations

import threading
import time
from collections import defaultdict
from pathlib import Path
from typing import Literal

from fastapi import HTTPException, UploadFile, status

from app.core.settings import get_settings

UploadFormat = Literal["jpeg", "png", "pdf"]

MSG_EXTERNAL_OCR_DISABLED = (
    "Внешнее распознавание изображений (OCR) отключено. "
    "Введите список плит текстом или включите OCR_EXTERNAL_ENABLED."
)

_READ_CHUNK = 1024 * 1024


class _CommercialOcrUploadLimiter:
    """In-process sliding window (not shared across workers); see COMMERCIAL_OCR_UPLOADS_PER_HOUR."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._events: dict[int, list[float]] = defaultdict(list)

    def check(self, user_id: int, *, max_events: int, window_seconds: float = 3600.0) -> None:
        now = time.monotonic()
        with self._lock:
            events = self._events[user_id]
            cutoff = now - window_seconds
            while events and events[0] < cutoff:
                events.pop(0)
            if len(events) >= max_events:
                raise HTTPException(
                    status_code=status.HTTP_429_TOO_MANY_REQUESTS,
                    detail="Превышен лимит загрузок для распознавания. Попробуйте позже.",
                )
            events.append(now)


_ocr_upload_limiter = _CommercialOcrUploadLimiter()


def reset_commercial_ocr_rate_limiter_for_tests() -> None:
    with _ocr_upload_limiter._lock:
        _ocr_upload_limiter._events.clear()


def check_commercial_ocr_rate_limit(user_id: int) -> None:
    lim = get_settings().commercial_ocr_uploads_per_hour
    if lim <= 0:
        return
    _ocr_upload_limiter.check(user_id, max_events=lim)


def ensure_external_ocr_enabled() -> None:
    """Reject OCR / external image recognition when ``OCR_EXTERNAL_ENABLED`` is false."""
    if not get_settings().ocr_external_enabled:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=MSG_EXTERNAL_OCR_DISABLED,
        )


async def read_upload_file_capped(upload: UploadFile, max_bytes: int | None = None) -> bytes:
    """Read upload in chunks; reject before buffering more than max_bytes (default from settings)."""
    limit = max_bytes if max_bytes is not None else get_settings().commercial_upload_max_bytes
    chunks: list[bytes] = []
    total = 0
    while True:
        chunk = await upload.read(_READ_CHUNK)
        if not chunk:
            break
        total += len(chunk)
        if total > limit:
            raise HTTPException(
                status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                detail="Размер файла превышает допустимый лимит.",
            )
        chunks.append(chunk)
    return b"".join(chunks)


def detect_upload_format(data: bytes) -> UploadFormat | None:
    if len(data) < 4:
        return None
    if data.startswith(b"\xff\xd8\xff"):
        return "jpeg"
    if len(data) >= 8 and data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith(b"%PDF"):
        return "pdf"
    return None


def require_magic_image_or_pdf(data: bytes) -> UploadFormat:
    if not data:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Пустой файл.",
        )
    fmt = detect_upload_format(data)
    if fmt is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Неподдерживаемый формат файла. Допустимы JPEG, PNG или PDF.",
        )
    return fmt


def normalized_ocr_filename(original: str | None, fmt: UploadFormat) -> str:
    """Basename only, extension from verified magic bytes (KP-001-safe, no path segments)."""
    ext_map = {"jpeg": ".jpg", "png": ".png", "pdf": ".pdf"}
    ext = ext_map[fmt]
    raw = (original or "").strip()
    base = Path(raw).name if raw else "upload"
    if ".." in base or "/" in base or "\\" in base:
        base = "upload"
    stem = Path(base).stem
    if not stem:
        stem = "upload"
    return f"{stem}{ext}"


async def prepare_commercial_ocr_upload(
    *,
    image: UploadFile | None,
    user_id: int,
    max_bytes: int | None = None,
) -> tuple[bytes | None, str | None]:
    """
    When a file part is present: enforce per-user OCR rate limit, read once with size cap,
    validate JPEG/PNG/PDF via magic bytes. Empty body → (None, None) without format error.
    """
    if image is None:
        return None, None
    ensure_external_ocr_enabled()
    check_commercial_ocr_rate_limit(user_id)
    raw = await read_upload_file_capped(image, max_bytes=max_bytes)
    if not raw:
        return None, None
    fmt = require_magic_image_or_pdf(raw)
    return raw, normalized_ocr_filename(image.filename, fmt)
