#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Промпты для OCR (Extract и Verify)."""

from __future__ import annotations

import json

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


CONTRACT_CARD_EXTRACT_PROMPT = """Прочитай карточку организации на изображении и верни реквизиты покупателя.

Это не таблица железобетонных плит и не марки изделий. Не выдумывай поля, которых нет на изображении.

Верни ТОЛЬКО JSON-объект:
{
  "legal_form": "ooo|ao|ip|kfh|person",
  "full_name": "",
  "short_name": "",
  "signatory_position": "",
  "signatory_name": "",
  "signatory_verb": "действующего|действующей",
  "authority_basis": "устав|доверенность",
  "poa_number": "",
  "poa_date": "",
  "inn": "",
  "kpp": "",
  "ogrn": "",
  "legal_address": "",
  "postal_address": "",
  "phone": "",
  "email": "",
  "bank_name": "",
  "account": "",
  "corr_account": "",
  "bik": "",
  "edo_operator": "",
  "edo_id": "",
  "okved": ""
}

Пустое поле — пустая строка. Цифры ИНН, КПП, ОГРН, БИК и счетов пиши без пробелов.
ОКВЭД — только основной код, например 23.61, без названия деятельности и без кодов после «Доп.». Если подписи нет — пустая строка.

Шапка «ИНН КПП ОГРН» обязательна, даже если ОГРН нет в таблице.
Счёт и корсчёт копируй по цифрам, повторы 0 и 7 не сжимай.
Ячейку «БИК … в банке» раздели на БИК и название банка.
"""


def get_contract_card_reread_prompt(field_names: list[str]) -> str:
    """Повтор только по полям, которые контроль не принял."""
    listed = ", ".join(field_names)
    return (
        "На изображении карточка контрагента. "
        f"Прочитай заново только эти поля: {listed}. "
        "Верни ТОЛЬКО JSON-объект с этими ключами. "
        "Цифры копируй как на изображении, повторы 0 и 7 не сжимай. "
        "Если запрошены и БИК, и банк, раздели ячейку «БИК в банке» на два значения."
    )


def get_contract_card_holes_prompt(
    *,
    holes: list[str],
    frozen: dict[str, str],
    card_text: str,
) -> str:
    """Текстовый добор только по дырам. Картинки в этом промпте нет."""
    keys = ", ".join(holes)
    frozen_json = json.dumps(frozen, ensure_ascii=False)
    banks = ""
    if "banks" in holes:
        banks = (
            "Ключ banks — список связок, а не один БИК. "
            "Каждая связка: bank_name, account, corr_account, bik.\n"
        )
    return (
        "Заполни по тексту карточки только запрошенные поля. "
        "Не выдумывай того, чего нет в тексте. "
        "Цифры копируй как в тексте, повторы 0 и 7 не сжимай.\n"
        f"{banks}"
        f"Нужны только эти ключи: {keys}.\n"
        f"Замороженные поля не меняй: {frozen_json}.\n"
        "Верни ТОЛЬКО JSON-объект с запрошенными ключами.\n\n"
        f"Текст карточки:\n{card_text}"
    )


def get_contract_card_verify_prompt(draft_json: str) -> str:
    """Сверка реквизитов карточки, не марок плит."""
    return f"""Ты сверяешь карточку организации с черновиком реквизитов покупателя.

На изображении — карточка контрагента, не таблица плит и не марки изделий.
Ниже — результат первого чтения (может содержать ошибки):

{draft_json}

Верни исправленные поля. Подменяй значение только если на изображении видно другое.
Если черновик верен — corrections пустой, fields совпадают с черновиком.
Если изображение не прочиталось — верни пустые fields и пустые corrections.

Формат ответа — ТОЛЬКО JSON:
{{
  "fields": {{
    "legal_form": "",
    "full_name": "",
    "short_name": "",
    "signatory_position": "",
    "signatory_name": "",
    "signatory_verb": "",
    "authority_basis": "",
    "poa_number": "",
    "poa_date": "",
    "inn": "",
    "kpp": "",
    "ogrn": "",
    "legal_address": "",
    "postal_address": "",
    "phone": "",
    "email": "",
    "bank_name": "",
    "account": "",
    "corr_account": "",
    "bik": "",
    "edo_operator": "",
    "edo_id": "",
    "okved": ""
  }},
  "corrections": [
    {{
      "field": "inn",
      "before": "",
      "after": "",
      "reason": ""
    }}
  ]
}}
"""
