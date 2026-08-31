#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пилот OCR: один снимок через recognize_text_smart.

Запуск из корня проекта:
    python scripts/ocr_pilot_compare.py --image path/to/photo.jpg
    python scripts/ocr_pilot_compare.py --image photo.jpg --provider gigachat --verify-mode auto

Требует настроенные credentials в .env (OPENAI_API_KEY или GIGACHAT_CREDENTIALS).
"""

from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

if sys.platform == "win32":
    import io

    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="OCR pilot: compare provider output on one image")
    parser.add_argument("--image", required=True, help="Path to image file (jpg/png/webp/pdf)")
    parser.add_argument(
        "--provider",
        choices=("gigachat", "openai"),
        default=None,
        help="OCR provider override (default: OCR_PROVIDER from env)",
    )
    parser.add_argument(
        "--verify-mode",
        choices=("auto", "always", "never"),
        default=None,
        help="Verify mode override (default: OCR_VERIFY_MODE from env)",
    )
    return parser.parse_args()


def _apply_env_overrides(args: argparse.Namespace) -> None:
    if args.provider:
        os.environ["OCR_PROVIDER"] = args.provider
    if args.verify_mode:
        os.environ["OCR_VERIFY_MODE"] = args.verify_mode

    from core.config.settings import get_settings

    get_settings.cache_clear()


def _format_corrections(corrections: list[dict]) -> str:
    if not corrections:
        return "(нет)"
    lines: list[str] = []
    for idx, item in enumerate(corrections, start=1):
        action = item.get("action") or "changed"
        row_index = item.get("row_index")
        label = f"стр. {row_index}" if row_index is not None else f"#{idx}"
        lines.append(f"  • {label}: {action} — {json.dumps(item, ensure_ascii=False)}")
    return "\n".join(lines)


async def _run(image_path: Path) -> int:
    from core.ocr_gpt import recognize_text_smart

    result = await recognize_text_smart(str(image_path), show_cost=True)
    if not result:
        print("[FAIL] OCR вернул пустой результат.")
        return 1

    plates = list(result.get("plates") or [])
    verify_decision = (
        result.get("ocr_verify_applied_reason")
        or result.get("ocr_verify_skipped_reason")
        or "unknown"
    )

    print("=" * 60)
    print(f"Image:        {image_path}")
    print(f"Method:       {result.get('ocr_method') or result.get('method')}")
    print(f"Plates:       {len(plates)}")
    print(f"API calls:    {result.get('ocr_api_calls', 1)}")
    print(f"Cost (RUB):   {float(result.get('ocr_cost_rub', 0.0) or 0.0):.4f}")
    print(f"Cost (USD):   {float(result.get('cost_usd', 0.0) or 0.0):.4f}")
    print(f"Verify:       applied={bool(result.get('verify_applied'))}, failed={bool(result.get('verify_failed'))}")
    print(f"Decision:     {verify_decision}")
    print(f"Select:       {result.get('ocr_verify_select_reason') or '-'}")
    print(f"Preprocess:   {result.get('ocr_preprocess') or '-'}")
    print("-" * 60)
    print("Recognized text:")
    print(result.get("text") or "(пусто)")
    print("-" * 60)
    print("Corrections:")
    print(_format_corrections(list(result.get("corrections") or [])))
    print("=" * 60)
    return 0


def main() -> int:
    args = _parse_args()
    image_path = Path(args.image).expanduser().resolve()
    if not image_path.is_file():
        print(f"[FAIL] Файл не найден: {image_path}")
        return 1

    _apply_env_overrides(args)
    return asyncio.run(_run(image_path))


if __name__ == "__main__":
    raise SystemExit(main())
