#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Протокол OCR-провайдера (Extract + Verify)."""

from __future__ import annotations

from typing import Any, Dict, List, Protocol, runtime_checkable


@runtime_checkable
class OcrProvider(Protocol):
    """Интерфейс провайдера OCR: извлечение и верификация списка плит."""

    async def extract_plates(
        self,
        *,
        user_text: str,
        image_base64: str | None = None,
        mime_type: str | None = None,
        max_tokens: int = 2500,
    ) -> tuple[List[Dict[str, Any]], float]:
        """Этап Extract: распознать плиты из текста/изображения."""
        ...

    async def verify_plates(
        self,
        *,
        image_base64: str,
        mime_type: str,
        draft_plates: List[Dict[str, Any]],
    ) -> tuple[Dict[str, Any], float]:
        """Этап Verify: сверить черновик с изображением."""
        ...
