"""Единая точка работы с именами плит (canonical/display/make).

Имя плиты в проекте появляется в трёх «диалектах»:

- ``"Плиты ПБ 60-12-8п"`` — как пишут менеджеры в исходных КП и как
  сохраняется в ``kp_plates.plate_name``.
- ``"ПБ 60-12-8п"`` — короткая форма, которую генерирует визуализация
  (``_make_plate_name`` в ``app/services/day_view_service.py``).
- ``"ПБ 60-12-8п"`` без префикса — старый ``_normalize_plate_name``
  в ``core/kp_db.py``.

При сравнении имён в разных слоях (агрегация ``plates_info`` в day_view,
поиск строки в ``find_one_row``) разный формат давал две «разные» плиты с
одинаковой геометрией. Этот модуль — единственный источник истины:
вся бизнес-логика сравнивает плиты только по ``canonical(name)``.

Не зависит от ``app/`` и ``bot/``: только стандартная библиотека.
"""
from __future__ import annotations

PREFIX_RU = "Плиты "


def canonical(name: str | None) -> str:
    """Каноническая форма имени плиты для сравнений.

    - снимает префикс ``"Плиты "`` (case-insensitive);
    - убирает повторяющиеся пробелы;
    - сохраняет регистр оригинала после префикса (``"ПБ 60-12-8п"``).

    Пустые / ``None`` имена возвращаются как ``""``.
    """
    if not name:
        return ""
    cleaned = str(name).strip()
    if cleaned.lower().startswith(PREFIX_RU.lower()):
        cleaned = cleaned[len(PREFIX_RU):].strip()
    # Сжимаем любые пробельные последовательности в один пробел —
    # «ПБ  60-12-8п» и «ПБ 60-12-8п» должны сравниваться равно.
    cleaned = " ".join(cleaned.split())
    return cleaned


def display(name: str | None) -> str:
    """Имя для отображения и записи в БД (``"Плиты ПБ 60-12-8п"``).

    Если входное имя уже содержит префикс — он не дублируется.
    Пустые имена возвращаются как ``""``.
    """
    if not name:
        return ""
    body = canonical(name)
    if not body:
        return ""
    return f"{PREFIX_RU}{body}"


def equal(left: str | None, right: str | None) -> bool:
    """True, если две строки представляют одну и ту же плиту."""
    return canonical(left) == canonical(right)


def make(
    length_m: float,
    width_mm: int,
    load_code: int,
    *,
    length_dm_raw: str | None = None,
) -> str:
    """Единая фабрика «короткого» имени плиты из геометрии.

    Возвращает короткую форму без префикса ``"Плиты "``: ``"ПБ 60-12-8п"``.
    Используется ``day_view_service`` и ``rescue_tracks`` для генерации имени,
    когда в lookup-таблицах нет точной записи КП.

    ``length_dm_raw``: если задано (в формате «60», «59,8», «59.8»), используется
    как есть, чтобы не терять разницу 59,8 vs 59,9 из-за float-округления.
    """
    if length_dm_raw and str(length_dm_raw).strip():
        length_str = str(length_dm_raw).strip().replace(".", ",")
    else:
        length_dm = float(length_m) * 10
        if abs(length_dm - round(length_dm)) < 0.01:
            length_str = str(int(round(length_dm)))
        else:
            length_str = (
                f"{length_dm:.1f}".rstrip("0").rstrip(".").replace(".", ",")
            )

    width_mm_int = int(width_mm)
    if width_mm_int == 1200:
        width_str = "12"
    else:
        width_dm = width_mm_int / 100.0
        if abs(width_dm - int(width_dm)) < 0.01:
            width_str = str(int(width_dm))
        else:
            width_str = str(width_dm).replace(".", ",")

    return f"ПБ {length_str}-{width_str}-{int(load_code)}п"


__all__ = ["canonical", "display", "equal", "make", "PREFIX_RU"]
