#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Промпты для OCR (Extract и Verify)."""

from __future__ import annotations

from core.plate_format_prompt import build_plate_parser_system_prompt

OCR_USER_PROMPT = (
    "Распознай таблицу на изображении. "
    "Верни все строки данных сверху вниз, без заголовков таблицы."
)


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
