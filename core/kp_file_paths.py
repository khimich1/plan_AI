# -*- coding: utf-8 -*-
"""Safe resolution of KP XLSX paths against configured storage roots."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Sequence


def _is_under_root(resolved: Path, root: Path) -> bool:
    try:
        resolved.relative_to(root)
        return True
    except ValueError:
        return False


def resolve_allowed_kp_xlsx_path(
    raw_path: str,
    *,
    allowed_roots: Sequence[Path],
    max_bytes: int | None = None,
) -> Path | None:
    """
    Resolve ``raw_path`` to a real file under one of ``allowed_roots``.

    Returns ``None`` if the path is missing, escapes allowed roots, or exceeds
    ``max_bytes`` when set.
    """
    if not raw_path or not str(raw_path).strip():
        return None

    candidate = Path(raw_path)
    if ".." in candidate.parts:
        return None

    try:
        resolved = candidate.expanduser().resolve(strict=True)
    except (OSError, RuntimeError):
        return None

    if not resolved.is_file():
        return None

    normalized_roots = [r.expanduser().resolve() for r in allowed_roots]
    if not any(_is_under_root(resolved, root) for root in normalized_roots):
        return None

    if max_bytes is not None:
        try:
            size = resolved.stat().st_size
        except OSError:
            return None
        if size > max_bytes:
            return None

    return resolved


def default_kp_xlsx_allowed_roots() -> list[Path]:
    """Storage roots permitted for KP XLSX reads (lazy settings import)."""
    from core.config.settings import PROJECT_ROOT, get_settings

    settings = get_settings()
    return [
        PROJECT_ROOT,
        settings.drafts_dir,
        settings.outputs_dir,
        settings.prices_dir,
    ]


def resolve_kp_xlsx_path_for_read(raw_path: str | None) -> Path | None:
    """Resolve path using project settings and commercial upload size limit."""
    if not raw_path:
        return None
    from core.config.settings import get_settings

    settings = get_settings()
    return resolve_allowed_kp_xlsx_path(
        raw_path,
        allowed_roots=default_kp_xlsx_allowed_roots(),
        max_bytes=settings.commercial_upload_max_bytes,
    )


def resolve_kp_xlsx_output_path(
    raw_path: str,
    *,
    allowed_roots: Sequence[Path],
) -> Path | None:
    """
    Resolve destination path for writing XLSX bytes.

    Unlike :func:`resolve_allowed_kp_xlsx_path`, the target file may not exist yet;
    the resolved path or its parent must lie under ``allowed_roots``.
    """
    if not raw_path or not str(raw_path).strip():
        return None

    candidate = Path(raw_path)
    if ".." in candidate.parts:
        return None

    try:
        resolved = candidate.expanduser().resolve(strict=False)
    except (OSError, RuntimeError):
        return None

    normalized_roots = [r.expanduser().resolve() for r in allowed_roots]
    for check in (resolved, resolved.parent):
        if any(_is_under_root(check, root) for root in normalized_roots):
            return resolved
    return None


def resolve_kp_xlsx_path_for_write(raw_path: str | None) -> Path | None:
    """Resolve write path using project storage roots (S8 parity with read)."""
    if not raw_path:
        return None
    return resolve_kp_xlsx_output_path(
        raw_path,
        allowed_roots=default_kp_xlsx_allowed_roots(),
    )
