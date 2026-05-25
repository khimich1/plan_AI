#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Распознавание и редактирование списков плит через GPT-4o Vision.
Один вызов GPT на операцию (OCR или инструкция пользователя).
"""

import os
import base64
import json
import logging
import re
from pathlib import Path
from typing import Optional, Dict, List, Any, Literal, Tuple, Callable, Awaitable

from core.plate_format_prompt import build_plate_parser_system_prompt

try:
    from openai import AsyncOpenAI
    GPT_AVAILABLE = True
except ImportError:
    GPT_AVAILABLE = False
    print("[GPT OCR] ⚠️ OpenAI не установлен. Установите: pip install openai")

_logger = logging.getLogger(__name__)

_VERIFY_FAILED_CORRECTION = {
    "action": "verify_failed",
    "row_index": None,
    "before": None,
    "after": None,
    "reason": "Повторная проверка не удалась, использован черновик первого этапа",
}


def _load_image_payload(image_path: str) -> Tuple[bytes, str, str]:
    """Читает файл и возвращает (bytes, base64, mime_type)."""
    with open(image_path, "rb") as f:
        image_data = f.read()
    image_base64 = base64.b64encode(image_data).decode()
    mime_type = _image_mime_type(image_path, image_data)
    return image_data, image_base64, mime_type


def _estimate_cost_usd(tokens_used: int) -> float:
    # GPT-4o: $2.50 за 1M токенов (упрощённая оценка по total_tokens)
    return (tokens_used / 1_000_000) * 2.5


def _plates_to_text(plates: List[Dict[str, Any]]) -> str:
    text_lines: List[str] = []
    for plate in plates:
        candidate = (plate.get("normalized_candidate") or plate.get("raw_name") or "").strip()
        qty = int(plate.get("qty", 1))
        if candidate:
            text_lines.append(f"{candidate} {qty}")
    return "\n".join(text_lines)


def _validate_plate_item(plate: Any) -> Optional[Dict[str, Any]]:
    if not isinstance(plate, dict):
        print(f"[GPT] ⚠️ Пропущена плита (не объект): {plate}")
        return None

    raw_name = plate.get("raw_name") or plate.get("name")
    normalized_candidate = plate.get("normalized_candidate") or raw_name
    if not raw_name or "qty" not in plate:
        print(f"[GPT] ⚠️ Пропущена плита (нет raw_name/name или qty): {plate}")
        return None

    try:
        qty = int(plate["qty"])
    except (ValueError, TypeError):
        print(f"[GPT] ⚠️ Пропущена плита (qty не число): {plate}")
        return None

    confidence = plate.get("confidence", 0.95)
    try:
        confidence = max(0.0, min(1.0, float(confidence)))
    except (ValueError, TypeError):
        confidence = 0.95

    issues = plate.get("issues") if isinstance(plate.get("issues"), list) else []
    return {
        "raw_name": str(raw_name).strip(),
        "normalized_candidate": str(normalized_candidate).strip(),
        "qty": qty,
        "confidence": confidence,
        "issues": issues,
    }


def _build_result_payload(
    *,
    plates: List[Dict[str, Any]],
    draft_plates: List[Dict[str, Any]],
    corrections: List[Dict[str, Any]],
    row_count_on_image: Optional[int],
    method: str,
    verify_applied: bool,
    verify_failed: bool,
    cost_usd: float,
) -> Dict[str, Any]:
    confidences = [float(p.get("confidence", 0.95)) for p in plates]
    avg_confidence = sum(confidences) / len(confidences) if confidences else 0.95
    return {
        "text": _plates_to_text(plates),
        "plates": plates,
        "draft_plates": draft_plates,
        "corrections": corrections,
        "row_count_on_image": row_count_on_image,
        "method": method,
        "verify_applied": verify_applied,
        "verify_failed": verify_failed,
        "confidence": avg_confidence,
        "cost_usd": cost_usd,
    }


_OCR_USER_PROMPT = (
    "Распознай таблицу на изображении. "
    "Верни все строки данных сверху вниз, без заголовков таблицы."
)


def _require_openai_client() -> AsyncOpenAI:
    if not GPT_AVAILABLE:
        raise RuntimeError("GPT недоступен. Установите: pip install openai")
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise ValueError(
            "Для распознавания по фото задайте OPENAI_API_KEY в окружении backend "
            "(docker-compose: сервис backend; локально: .env или экспорт переменной)."
        )
    return AsyncOpenAI(api_key=api_key)


async def _call_gpt_for_plates(
    *,
    user_text: str,
    client: AsyncOpenAI,
    image_base64: str | None = None,
    mime_type: str | None = None,
    max_tokens: int = 2500,
) -> tuple[List[Dict[str, Any]], float]:
    user_content: List[Dict[str, Any]] = [{"type": "text", "text": user_text}]
    if image_base64 and mime_type:
        user_content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{mime_type};base64,{image_base64}",
                    "detail": "high",
                },
            }
        )

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": build_plate_parser_system_prompt()},
            {"role": "user", "content": user_content},
        ],
        max_tokens=max_tokens,
        temperature=0.0,
    )

    result_text = response.choices[0].message.content or ""
    plates = parse_gpt_response(result_text)
    tokens_used = response.usage.total_tokens if response.usage else 0
    cost_usd = _estimate_cost_usd(tokens_used)
    return plates, cost_usd


async def recognize_text_smart(
    image_path: str,
    force_gpt: bool = False,
    show_cost: bool = True,
    mode: Literal["full_gpt", "hybrid"] = "full_gpt",
    verify_enabled: Optional[bool] = None,
    on_status: Optional[Callable[[str], Awaitable[None]]] = None,
) -> Optional[Dict]:
    """
    Распознавание таблицы через один вызов GPT-4o Vision.

    verify_enabled и on_status оставлены для обратной совместиости и игнорируются.
    """
    _ = (force_gpt, verify_enabled, on_status)
    if mode == "hybrid":
        print("[OCR] ℹ️ Режим hybrid отключен: используется только GPT-4o")

    if not GPT_AVAILABLE:
        print("[OCR] ❌ GPT недоступен. Установите: pip install openai")
        return None

    try:
        client = _require_openai_client()
        image_data, image_base64, mime_type = _load_image_payload(image_path)
        image_size_kb = len(image_data) / 1024
        print(f"[GPT] Размер изображения: {image_size_kb:.1f} КБ")

        print("[OCR] GPT-4o Vision (один вызов)...")
        plates, cost_usd = await _call_gpt_for_plates(
            user_text=_OCR_USER_PROMPT,
            client=client,
            image_base64=image_base64,
            mime_type=mime_type,
        )

        if not plates:
            return None

        if show_cost:
            rub_cost = cost_usd * 75
            print(f"[OCR] 💰 Стоимость: ${cost_usd:.4f} (~{rub_cost:.2f}₽)")

        print(f"[OCR] ✅ Итого {len(plates)} строк(и), method=GPT-4o")
        return _build_result_payload(
            plates=plates,
            draft_plates=plates,
            corrections=[],
            row_count_on_image=len(plates),
            method="GPT-4o",
            verify_applied=False,
            verify_failed=False,
            cost_usd=cost_usd,
        )

    except Exception as e:
        print(f"[OCR] ❌ Ошибка GPT: {e}")
        import traceback
        traceback.print_exc()

    return None


async def apply_plates_with_ai(
    *,
    current_plates_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """
    Применяет инструкцию пользователя к списку плит (опционально с изображением).
    Один вызов GPT-4o, temperature=0.
    """
    instruction = (user_instruction or "").strip()
    if not instruction:
        raise ValueError("Инструкция для ИИ не может быть пустой.")

    client = _require_openai_client()
    current_text = (current_plates_text or "").strip() or "(пусто)"
    user_text = (
        f"Текущий список плит:\n{current_text}\n\n"
        f"Инструкция пользователя:\n{instruction}"
    )

    image_base64: str | None = None
    mime_type: str | None = None
    if image_path:
        _, image_base64, mime_type = _load_image_payload(image_path)

    print("[AI] GPT-4o (один вызов, инструкция пользователя)...")
    plates, cost_usd = await _call_gpt_for_plates(
        user_text=user_text,
        client=client,
        image_base64=image_base64,
        mime_type=mime_type,
    )

    if not plates:
        return None

    if show_cost:
        rub_cost = cost_usd * 75
        print(f"[AI] 💰 Стоимость: ${cost_usd:.4f} (~{rub_cost:.2f}₽)")

    print(f"[AI] ✅ Итого {len(plates)} строк(и), method=GPT-4o+ai")
    return _build_result_payload(
        plates=plates,
        draft_plates=plates,
        corrections=[],
        row_count_on_image=len(plates),
        method="GPT-4o+ai",
        verify_applied=False,
        verify_failed=False,
        cost_usd=cost_usd,
    )


def _image_mime_type(image_path: str, image_data: bytes) -> str:
    """MIME для data: URL Vision API (PNG раньше ошибочно слали как image/jpeg)."""
    if len(image_data) >= 8 and image_data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if len(image_data) >= 3 and image_data.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if len(image_data) >= 12 and image_data.startswith(b"RIFF") and image_data[8:12] == b"WEBP":
        return "image/webp"
    if len(image_data) >= 6 and image_data.startswith((b"GIF87a", b"GIF89a")):
        return "image/gif"
    if len(image_data) >= 4 and image_data.startswith(b"%PDF"):
        return "application/pdf"
    suffix = Path(image_path).suffix.lower()
    return {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".gif": "image/gif",
        ".pdf": "application/pdf",
    }.get(suffix, "image/jpeg")


async def recognize_with_gpt_vision(
    image_path: str,
    *,
    client: Optional[AsyncOpenAI] = None,
    image_base64: Optional[str] = None,
    mime_type: Optional[str] = None,
) -> tuple[List[Dict], float]:
    """Legacy wrapper: один вызов GPT-4o Vision для OCR."""
    if client is None:
        client = _require_openai_client()
    if image_base64 is None or mime_type is None:
        _, image_base64, mime_type = _load_image_payload(image_path)
    plates, cost_usd = await _call_gpt_for_plates(
        user_text=_OCR_USER_PROMPT,
        client=client,
        image_base64=image_base64,
        mime_type=mime_type,
    )
    return plates, cost_usd


async def verify_plates_with_gpt_vision(
    *,
    image_base64: str,
    mime_type: str,
    draft_plates: List[Dict[str, Any]],
    client: AsyncOpenAI,
) -> tuple[Dict[str, Any], float]:
    """
    Этап 2 Verify: сверка черновика с изображением.
    """
    draft_json = json.dumps(draft_plates, ensure_ascii=False, indent=2)

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": get_verification_prompt(draft_json)},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:{mime_type};base64,{image_base64}",
                            "detail": "high",
                        },
                    },
                ],
            }
        ],
        max_tokens=2500,
        temperature=0.0,
    )

    result_text = response.choices[0].message.content or ""
    verify_result = parse_verify_response(result_text)

    tokens_used = response.usage.total_tokens if response.usage else 0
    cost_usd = _estimate_cost_usd(tokens_used)
    print(
        f"[GPT] Verify: токенов {tokens_used}, "
        f"строк {len(verify_result.get('plates') or [])}, "
        f"исправлений {len(verify_result.get('corrections') or [])}"
    )

    return verify_result, cost_usd


def get_recognition_prompt() -> str:
    """Deprecated: используйте build_plate_parser_system_prompt()."""
    return build_plate_parser_system_prompt()


def get_verification_prompt(draft_json: str) -> str:
    """Промпт этапа Verify — аудитор черновика по изображению."""
    return f"""Ты аудитор OCR для таблицы железобетонных плит.

На изображении — исходная таблица (колонки: наименование | количество).
Ниже — результат ПЕРВОГО распознавания (может содержать ошибки):

{draft_json}

🎯 ЗАДАЧА: сверить КАЖДУЮ строку черновика с изображением и вернуть ИСПРАВЛЕННЫЙ список.

Работай как корректор, не как составитель:
1. Посчитай строки данных на изображении (без заголовков «Наименование», «Кол-во», «Итого»).
2. Если в черновике строк меньше или больше — добавь пропущенные / убери лишние.
3. Для каждой строки проверь ПОСИМВОЛЬНО:
   - normalized_candidate (марка плиты)
   - qty (число в правой колонке)
4. Особое внимание:
   - похожие марки: 66,2-2,6-8п vs 66,2-9,2-8п vs 66,2-9,2-6п
   - запятые и ",0": 52,0 ≠ 52
   - нагрузка 6п vs 8п — разные изделия
   - qty 1 vs 2 в узкой колонке
5. НЕ упрощай числа, НЕ округляй, НЕ группируй одинаковые строки.
6. Порядок строк — сверху вниз как на изображении.

Формат ответа — ТОЛЬКО JSON объект:
{{
  "row_count_on_image": 26,
  "plates": [
    {{
      "raw_name": "...",
      "normalized_candidate": "...",
      "qty": 1,
      "confidence": 0.98,
      "issues": []
    }}
  ],
  "corrections": [
    {{
      "action": "added|removed|changed_mark|changed_qty|reordered",
      "row_index": 5,
      "before": {{"normalized_candidate": "...", "qty": 1}},
      "after": {{"normalized_candidate": "...", "qty": 2}},
      "reason": "на изображении qty=2, в черновике было 1"
    }}
  ]
}}

Если черновик полностью верен — plates совпадают, corrections = [].
Верни ТОЛЬКО JSON, без текста до и после."""


def _extract_json_from_response(response_text: str) -> str:
    text = (response_text or "").strip()
    object_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if object_match:
        return object_match.group(1)
    array_match = re.search(r"```(?:json)?\s*(\[.*?\])\s*```", text, re.DOTALL)
    if array_match:
        return array_match.group(1)
    if text.startswith("{") or text.startswith("["):
        return text
    brace_match = re.search(r"(\{.*\})", text, re.DOTALL)
    if brace_match:
        return brace_match.group(1)
    bracket_match = re.search(r"(\[.*\])", text, re.DOTALL)
    if bracket_match:
        return bracket_match.group(1)
    return text


def parse_gpt_response(response_text: str) -> List[Dict[str, Any]]:
    """Извлекает JSON-массив плит из ответа GPT (этап Extract)."""
    json_text = _extract_json_from_response(response_text)

    try:
        parsed = json.loads(json_text)
        if isinstance(parsed, dict) and "plates" in parsed:
            parsed = parsed["plates"]
        if not isinstance(parsed, list):
            print(f"[GPT] ⚠️ Ожидался JSON-массив, получено: {type(parsed)}")
            return []

        validated_plates = []
        for plate in parsed:
            item = _validate_plate_item(plate)
            if item:
                validated_plates.append(item)
        return validated_plates

    except json.JSONDecodeError as e:
        print(f"[GPT] ❌ Ошибка парсинга JSON: {e}")
        print(f"[GPT] Ответ GPT (первые 200 символов):")
        print(response_text[:200])
        return []


def parse_verify_response(response_text: str) -> Dict[str, Any]:
    """
    Парсит ответ этапа Verify.
    Поддерживает полный объект {{plates, corrections}} и legacy-массив plates.
    """
    json_text = _extract_json_from_response(response_text)

    try:
        parsed = json.loads(json_text)

        if isinstance(parsed, list):
            plates_raw = parsed
            corrections: List[Dict[str, Any]] = []
            row_count_on_image = len(plates_raw)
        elif isinstance(parsed, dict):
            plates_raw = parsed.get("plates") or []
            corrections = parsed.get("corrections") if isinstance(parsed.get("corrections"), list) else []
            row_count_raw = parsed.get("row_count_on_image")
            try:
                row_count_on_image = int(row_count_raw) if row_count_raw is not None else None
            except (ValueError, TypeError):
                row_count_on_image = None
        else:
            print(f"[GPT] ⚠️ Verify: неожиданный тип JSON: {type(parsed)}")
            return {"plates": [], "corrections": [], "row_count_on_image": None}

        validated_plates = []
        for plate in plates_raw:
            item = _validate_plate_item(plate)
            if item:
                validated_plates.append(item)

        return {
            "plates": validated_plates,
            "corrections": corrections,
            "row_count_on_image": row_count_on_image,
        }

    except json.JSONDecodeError as e:
        print(f"[GPT] ❌ Verify: ошибка парсинга JSON: {e}")
        print(f"[GPT] Ответ GPT (первые 200 символов):")
        print(response_text[:200])
        return {"plates": [], "corrections": [], "row_count_on_image": None}


def format_corrections_for_user(
    corrections: List[Dict[str, Any]],
    *,
    max_items: int = 8,
) -> str:
    """Краткий текст исправлений для Telegram."""
    actionable = [
        c for c in corrections
        if c.get("action") != "verify_failed"
    ]
    if not actionable:
        return ""

    lines = [f"⚠️ Автоисправлено {len(actionable)} строк(и):"]
    for idx, item in enumerate(actionable[:max_items], start=1):
        action = item.get("action") or "changed"
        row_index = item.get("row_index")
        row_label = f"стр. {row_index}" if row_index is not None else f"#{idx}"

        before = item.get("before") or {}
        after = item.get("after") or {}
        before_mark = before.get("normalized_candidate") or before.get("raw_name") or "—"
        after_mark = after.get("normalized_candidate") or after.get("raw_name") or "—"
        before_qty = before.get("qty")
        after_qty = after.get("qty")

        if action == "added":
            mark = after_mark
            qty = after_qty if after_qty is not None else "?"
            lines.append(f"• {row_label}: добавлено «{mark} {qty}»")
        elif action == "removed":
            lines.append(f"• {row_label}: удалено «{before_mark}»")
        elif action == "changed_qty":
            lines.append(
                f"• {row_label}: «{after_mark}» qty {before_qty} → {after_qty}"
            )
        elif action == "changed_mark":
            lines.append(
                f"• {row_label}: «{before_mark}» → «{after_mark}»"
            )
        elif action == "reordered":
            lines.append(f"• {row_label}: изменён порядок")
        else:
            reason = item.get("reason") or action
            lines.append(f"• {row_label}: {reason}")

    if len(actionable) > max_items:
        lines.append(f"• … и ещё {len(actionable) - max_items}")

    return "\n".join(lines)


def estimate_monthly_cost(photos_per_month: int) -> Dict[str, float]:
    """Оценка месячных затрат на OCR (один вызов GPT-4o)."""
    avg_cost_per_photo = 0.002

    return {
        "gpt_only": photos_per_month * avg_cost_per_photo,
        "hybrid": photos_per_month * avg_cost_per_photo,
        "photos": photos_per_month,
    }


if __name__ == "__main__":
    print("💰 Оценка месячных затрат на OCR (один вызов GPT-4o):")
    print("=" * 50)

    for count in [100, 500, 1000, 5000]:
        costs = estimate_monthly_cost(count)
        print(f"\n📊 {count} фото в месяц:")
        print(f"  • GPT-4o: ${costs['gpt_only']:.2f} (~{costs['gpt_only'] * 75:.0f}₽)")
