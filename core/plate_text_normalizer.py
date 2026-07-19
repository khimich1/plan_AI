#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Нормализация текста заказа плит перед парсингом.

Конвертирует каталожные марки вида «ПБ 59.12-8Вр1400-25» в канонический
формат парсера «ПБ 59-12-8п».  Исправляет OCR-ошибки в префиксах (ПВ→ПБ),
нормализует тире и пробелы.

Точка входа: normalize_order_text(text) → NormalizeResult
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from .plate_line_parser import match_bare_plate_line

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Константы
# ---------------------------------------------------------------------------

# OCR-ошибки в префиксе плиты
_OCR_PREFIX_FIXES: dict[str, str] = {"пв": "пб", "пг": "пб", "пе": "пб"}

# Юникодные тире → ASCII дефис
_DASH_CHARS = "–—‒−"

# Знаки умножения → x
_MULT_CHARS = "×х"

# ---------------------------------------------------------------------------
# Регулярные выражения
# ---------------------------------------------------------------------------

# Каталожный формат марки: ПБ L.W-load[фабричный суффикс]
#
# Структура в каталожной нотации:
#   «ПБ 59.12-8Вр1400-25»  →  L=59дм (5.9 м), W=12дм (1.2 м), нагрузка=8
#   «ПБ56.05-10»            →  L=56дм (5.6 м), W=5дм  (0.5 м), нагрузка=10
#
# Ключевое отличие от обычного формата «ПБ 59-12-8п»:
#   - DOT разделяет длину и ширину в дм (а не десятичная часть числа)
#   - Нет суффикса «п» после нагрузки
#   - Может следовать суффикс типа «Вр1400-25»
#
# Группы:
#   G1 – префикс (ПБ / ПК / ПВ=OCR-ошибка)
#   G2 – L_dm  : целое число, длина в дм
#   G3 – W_dm  : 1–2 цифры, ширина в дм
#   G4 – load  : нагрузка (целое или дробное)
#
# Необязательный суффикс (Вр1400-25 и т.п.) потребляется жадно.
_CATALOG_CORE_RE = re.compile(
    r"(?i)"
    r"\b(п[бвк])\s*"                                   # G1: префикс
    r"(\d+)\."                                          # G2: L_dm + точка
    r"(\d{1,2})"                                        # G3: W_dm (1–2 цифры)
    r"\s*-\s*"
    r"(\d+(?:[,.]\d+)?)"                                # G4: нагрузка
    r"(?:\s*[А-Яа-яёЁA-Za-z]\w*(?:\s*-\s*\d+)*)?",   # суффикс (опц.)
    re.UNICODE,
)

# Количество после марки: «5», «5 шт», «5шт», «5 штук», «5 шт.»
_QTY_RE = re.compile(r"^(\d+)\s*[а-яёА-ЯЁa-zA-Z]*\.?\s*$")


# ---------------------------------------------------------------------------
# Публичный тип данных
# ---------------------------------------------------------------------------

@dataclass
class NormalizeResult:
    """Результат нормализации текста заказа плит."""

    normalized_text: str
    normalized_lines: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    unrecognized_lines: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Вспомогательные функции (строительные блоки)
# ---------------------------------------------------------------------------

def basic_text_cleanup(text: str) -> str:
    """Нормализует неразрывные пробелы, тире и знаки умножения."""
    text = text.replace("\u00a0", " ")       # неразрывный пробел
    for ch in _DASH_CHARS:
        text = text.replace(ch, "-")
    for ch in _MULT_CHARS:
        text = text.replace(ch, "x")
    text = text.replace("\u00d7", "x")       # ×
    return text


def normalize_plate_prefixes(line: str) -> str:
    """
    Исправляет OCR-ошибки в префиксе (ПВ/ПГ → ПБ) и добавляет пробел
    после ПБ/ПК, если он слит с цифрой: «ПБ56» → «ПБ 56».
    """
    # ПВ/ПГ/ПЕ → ПБ (включая варианты с пробелом/точкой/запятой перед цифрой)
    fixed = re.sub(
        r"\b(пв|пг|пе)(?=[\s\.,]*\d)",
        lambda m: "ПБ" if m.group(1).isupper() else "пб",
        line,
        flags=re.IGNORECASE,
    )
    # Унификация разделителя перед цифрой: «ПБ.56», «ПБ,56», «ПБ56» → «ПБ 56»
    fixed = re.sub(r"\b(п[бк])[\s\.,]*(?=\d)", r"\1 ", fixed, flags=re.IGNORECASE)
    # Финальная защита от дубля пробелов
    fixed = re.sub(r"\s{2,}", " ", fixed)
    return fixed.strip()


def parse_catalog_mark(
    line: str,
) -> Optional[Tuple[str, int, int, float, int]]:
    """
    Разбирает каталожную марку вида «ПБ L.W-load[суффикс][кол-во]».

    Args:
        line: одна строка (регистр не важен, базовая чистка уже применена).

    Returns:
        (prefix, L_dm, W_dm, load, qty) или None, если строка не является
        каталожной маркой.

    Условия применения:
        - L_dm >= 20  (плита длиннее 2 м = 20 дм)
        - 1 <= W_dm <= 15  (ширина 0.1–1.5 м = 1–15 дм)
    """
    s = line.strip()
    m = _CATALOG_CORE_RE.search(s)
    if not m:
        return None

    raw_prefix = m.group(1)
    L_dm = int(m.group(2))
    W_dm = int(m.group(3))          # int("05") == 5
    load_str = m.group(4)

    # Фильтры реалистичности
    if L_dm < 20:
        # Слишком маленькое число для длины в дм → вероятно, десятичная запись
        return None
    if W_dm == 0 or W_dm > 15:
        # Ширина вне диапазона реальных плит
        return None

    try:
        load = float(load_str.replace(",", "."))
    except ValueError:
        load = 8.0

    # Кол-во: ищем цифру в остатке строки после сматченной части
    remainder = s[m.end():].strip()
    qty = 1
    if remainder:
        qty_m = _QTY_RE.match(remainder)
        if qty_m:
            qty = max(1, int(qty_m.group(1)))

    # Нормализуем префикс: ПВ → ПБ, сохраняем ПК
    prefix_lower = raw_prefix.lower()
    prefix_fixed = _OCR_PREFIX_FIXES.get(prefix_lower, prefix_lower)
    prefix = prefix_fixed.upper()          # «ПБ» или «ПК»

    return prefix, L_dm, W_dm, load, qty


def _format_width_dm(W_dm: int) -> str:
    """
    Форматирует целое значение ширины (в дм) как строку для канонической марки.

    Однозначные значения получают «,0» для явности: 5 → «5,0», 2 → «2,0».
    Двузначные остаются как есть: 12 → «12», 10 → «10».

    Это необходимо: последующий normalize_dimension() правильно интерпретирует
    «12» как 12 дм и «5,0» как 5.0 дм.
    """
    if W_dm < 10:
        return f"{W_dm},0"
    return str(W_dm)


def _format_load(load: float) -> str:
    """8.0 → «8п», 10.0 → «10п», 12.5 → «12,5п»."""
    if abs(load - round(load)) < 1e-6:
        return f"{int(round(load))}п"
    return f"{load:.1f}п".replace(".", ",")


# ---------------------------------------------------------------------------
# Основные публичные функции
# ---------------------------------------------------------------------------

def canonicalize_plate_line(line: str) -> Tuple[str, Optional[str]]:
    """
    Нормализует одну строку заказа плит.

    Если строка содержит каталожную марку — преобразует в канонический вид
    «ПБ L-W-Nп [qty]».  Иначе возвращает строку после базовой чистки.

    Returns:
        (нормализованная_строка, предупреждение_или_None)
        Предупреждение описывает выполненное преобразование.
    """
    stripped = line.strip()
    if not stripped:
        return stripped, None

    # Базовая чистка: тире, умножение, неразрывные пробелы
    cleaned = basic_text_cleanup(stripped)
    # Исправляем OCR-ошибки в префиксе и добавляем пробел
    cleaned = normalize_plate_prefixes(cleaned)

    # Пробуем каталожный формат
    result = parse_catalog_mark(cleaned)
    if result is not None:
        prefix, L_dm, W_dm, load, qty = result
        W_str = _format_width_dm(W_dm)
        load_str = _format_load(load)
        canonical = f"{prefix} {L_dm}-{W_str}-{load_str}"
        if qty > 1:
            canonical += f" {qty}"
        warning = f"{stripped!r} → {canonical!r}"
        logger.debug("Каталожная марка нормализована: %s", warning)
        return canonical, warning

    # Не каталожная марка — возвращаем строку после базовой чистки
    return cleaned, None


def get_wide_plate_lines(text: str) -> List[Tuple[str, int]]:
    """
    Возвращает строки из текста заказа, у которых ширина в дм > 12.

    Поддерживает форматы:
    - Каталожная/стандартная марка: ПБ L-W-Nп, ПБ L.W-load и т.п.
      (разбор через parse_catalog_mark по нормализованной строке)
    - Канонический/OCR: ПБ L-W-N [qty] без «п» (например, вывод OCR/GPT)
    - Марка без префикса: «71-15-8 2», «65,6-12-12,5 2»
    - Размерный формат W×L в метрах: «1.5×6.3 — 2 шт», «0.32×4.0 3»
      (ширина 1.5 м = 15 дм > 12)

    Args:
        text: исходный текст заказа.

    Returns:
        Список пар (исходная_строка, qty) для каждой строки с шириной > 12 дм.
    """
    if not text or not text.strip():
        return []

    raw_lines = [l.strip() for l in re.split(r"[\n;]+", text) if l.strip()]
    result: List[Tuple[str, int]] = []

    # Regex для формата WxL (метры): два числа с разделителем x/х/×
    _DIM_RE = re.compile(
        r"(\d+(?:[.,]\d+)?)\s*[xхх×]\s*(\d+(?:[.,]\d+)?)",
        re.IGNORECASE | re.UNICODE,
    )
    # Regex для количества в конце строки
    _QTY_END_RE = re.compile(r"(\d+)\s*(?:шт\.?|штук)?\s*$", re.IGNORECASE)
    # Канонический формат: ПБ L-W-N[п] [qty] (дефис между L и W; «п» опциональна для OCR)
    # Поддерживает дробные L/W: "68,6-15-8", "68,6-1,5-8", "66,2-6,65-8п".
    _CANONICAL_L_W_RE = re.compile(
        r"(?i)\bп[бк]\s*(\d+(?:[,.]\d+)?)\s*-\s*(\d+(?:[,.]\d+)?)\s*-\s*\d+(?:[,.]\d+)?п?(?:\s+(\d+))?\s*$",
        re.UNICODE,
    )

    def _width_token_to_dm(token: str) -> float:
        """
        Приводит ширину из марки ПБ/ПК к дм для проверки >12 дм.

        Поддерживаемые варианты в марке:
        - "15"   -> 15 дм
        - "6,65" -> 6.65 дм
        - "1,5"  -> 15 дм (если ширина записана в метрах)
        """
        s = (token or "").strip().replace(",", ".")
        try:
            val = float(s)
        except ValueError:
            return 0.0
        # Если число с разделителем и <= 2, трактуем как метры.
        if "." in s and val <= 2.0:
            return val * 10.0
        return val

    for raw in raw_lines:
        cleaned = basic_text_cleanup(raw)
        cleaned = normalize_plate_prefixes(cleaned)

        # 1) Попробовать каталожную марку ПБ L.W-load (с точкой)
        parsed = parse_catalog_mark(cleaned)
        if parsed is not None:
            _prefix, _L, W_dm, _load, qty = parsed
            if W_dm > 12:
                result.append((raw, qty))
            continue

        # 2) Канонический формат ПБ L-W-N[п] [qty] (нормализованные строки и OCR)
        m_can = _CANONICAL_L_W_RE.match(cleaned)
        if m_can:
            W_dm = _width_token_to_dm(m_can.group(2))
            qty = int(m_can.group(3)) if m_can.group(3) else 1
            if W_dm > 12:
                result.append((raw, qty))
            continue

        # 3) Марка без префикса ПБ: L-W-load [qty]
        bare = match_bare_plate_line(cleaned)
        if bare:
            _ldm, wdm_raw, _load, qty = bare
            W_dm = _width_token_to_dm(wdm_raw)
            if W_dm > 12:
                result.append((raw, qty))
            continue

        # 4) Попробовать размерный формат W×L (значения в метрах или дм)
        m = _DIM_RE.search(cleaned.replace(",", "."))
        if m:
            try:
                first = float(m.group(1))
                second = float(m.group(2))
            except ValueError:
                continue

            # Определяем ширину: если первое значение <= 2.0 и второе > первого — это W×L
            # если первое > 2.0 — скорее всего L×W, меняем местами
            if first > 2.0 and second <= 2.0:
                width_m = second
            else:
                width_m = first

            width_dm = width_m * 10.0
            if width_dm > 12.0:
                # Определяем количество из остатка строки
                qty = 1
                remainder = cleaned[m.end():]
                qty_m = _QTY_END_RE.search(remainder)
                if qty_m:
                    qty = max(1, int(qty_m.group(1)))
                result.append((raw, qty))

    return result


def normalize_order_text(text: str) -> NormalizeResult:
    """
    Нормализует полный текст заказа плит перед передачей в парсер.

    Конвертирует каталожные марки (ПБ L.W-loadВр…) в канонический вид
    (ПБ L-W-loadп).  Исправляет OCR-ошибки в префиксах.  Нормализует
    пробелы и тире.

    Args:
        text: исходный текст заказа (многострочный или однострочный).

    Returns:
        NormalizeResult с нормализованными строками и диагностикой.
    """
    if not text or not text.strip():
        return NormalizeResult(normalized_text=text or "")

    raw_lines = [l.strip() for l in re.split(r"[\n;]+", text) if l.strip()]

    normalized_lines: List[str] = []
    warnings: List[str] = []

    for raw in raw_lines:
        canonical, warn = canonicalize_plate_line(raw)
        normalized_lines.append(canonical)
        if warn:
            warnings.append(warn)

    normalized_text = "\n".join(normalized_lines)

    if warnings:
        logger.info(
            "Нормализатор: %d строк преобразовано из каталожного формата",
            len(warnings),
        )

    return NormalizeResult(
        normalized_text=normalized_text,
        normalized_lines=normalized_lines,
        warnings=warnings,
        unrecognized_lines=[],
    )
