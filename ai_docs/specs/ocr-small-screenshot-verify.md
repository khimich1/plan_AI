# Spec: OCR мелких скриншотов — Verify по короткой стороне

> **Тип:** feature-spec (SDD Phase: SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅)  
> **Дата:** 2026-08-22  
> **Статус:** код этапа B влит в рабочее дерево; пилот S9 вручную  
> **Источник идеи:** [`ai_docs/ideas/ocr-small-screenshot-verify.md`](../ideas/ocr-small-screenshot-verify.md)  
> **План реализации:** [`ai_docs/develop/plans/2026-08-22-ocr-small-screenshot-verify.md`](../develop/plans/2026-08-22-ocr-small-screenshot-verify.md)  
> **Связанные документы:** [`gigachat-vision-ocr.md`](./gigachat-vision-ocr.md), [`pile-ocr-reliability.md`](./pile-ocr-reliability.md)

---

## ASSUMPTIONS I'M MAKING

1. Этот документ — **только этап B**: не менять пиксели. Этап A (2× Lanczos + контраст) — **другая спека**, только если пилот B не снизит ошибки марок/пропусков строк.
2. Канал — **веб wizard КП**. Telegram-бот вне scope.
3. Типичный плохой вход — **скриншот** Excel / PDF / переписки, не размытое фото и не «низкий DPI».
4. Успех для менеджера — **заметно меньше правок**, не 90% и не «никогда не ошибаться». Человек остаётся в контуре.
5. Один порог на **все шесть** OCR-пайплайнов (плиты, сваи, ступени, марши, мостовые сваи, ФБС) — проблема общая, политика уже копируется.
6. Default порог короткой стороны — **1000 px** (середина 800–1200 из идеи). `0` выключает эвристику. Калибровка после пилота — Ask first.
7. Сравнение: `short_side < min_short_side` → Verify. Равенство порогу **не** триггерит (как у `max_rows` / `max_bytes`).
8. Если размер кадра **не прочитать** (PDF, битый файл) → Verify, reason `auto_image_size_unknown`. Безопаснее пропустить эвристику.
9. `OCR_VERIFY_MODE=never` и `OCR_MAX_API_CALLS=1` по-прежнему **безусловно** глушат Verify.
10. Новых зависимостей нет (Pillow уже есть). UI, промпты, JSON-контракт OCR, схема БД — без изменений.
11. Объём 50–200 фото/мес; лишний Verify на мелком скрине удваивает вызов, latency цель **< 15 с p95** не ослабляем.

→ Поправь допущения до PLAN / IMPLEMENT.

---

## Decisions locked (предложение)

| # | Тема | Решение |
|---|------|---------|
| D1 | Этап | **B only.** A (апскейл) — out of scope |
| D2 | Метрика кадра | Короткая сторона в px, не байты и не DPI |
| D3 | Default | `OCR_VERIFY_AUTO_MIN_SHORT_SIDE=1000`; `0` = выкл |
| D4 | Scope изделий | Все 6 `should_run_*_verify` |
| D5 | Нечитаемый размер | Verify (`auto_image_size_unknown`) |
| D6 | Порядок reason | Мелкий кадр **до** row/confidence, чтобы пилот видел `auto_small_image` |
| D7 | Общий код | Один helper на 6 функций, не шестая копия if-цепочки |
| D8 | UI | Не меняем; reason уже уходит в `ocr_verify_applied_reason` |

---

## Objective

Закрыть дыру в `auto`-Verify: мелкий «уверенный» скриншот с коротким списком сейчас пропускает второй проход, и модель тихо путает марку или теряет строку.

### Проблема

`should_run_*_verify` в режиме `auto` включает второй вызов, когда extract «грязный» или файл **большой** (`> OCR_VERIFY_AUTO_MAX_BYTES`, ~800 КБ). Мелкий скрин 5–8 строк с confidence ≥ 0.92 и валидным parse даёт `auto_all_checks_passed` → 1 вызов. Vision-модель при этом ошибается на похожих марках и пропускает строки.

DPI и слепой апскейл эту дыру не закрывают: модель видит пиксели, не метку dpi.

### Пользователь

Менеджер по продажам на шаге ввода состава КП: загружает скрин таблицы, сверяет черновик, жмёт «Список верен».

### User Stories

- Как **менеджер**, я загружаю мелкий скриншот таблицы и получаю тот же UX, но реже правую марки и пропавшие строки — система сама делает второй проход по фото.
- Как **менеджер**, я загружаю крупный чёткий скрин / фото и по-прежнему могу получить 1 вызов, если extract чистый.
- Как **оператор**, я меняю порог короткой стороны через env без правки кода; `0` выключает эвристику.
- Как **оператор**, я вижу в metadata `ocr_verify_applied_reason=auto_small_image` (или `auto_image_size_unknown`) и в логах `short_side_px`.

### Success Criteria (измеримые)

| # | Критерий | Метод проверки |
|---|----------|----------------|
| S1 | Чистый extract + `short_side < 1000` → Verify, reason `auto_small_image` | `pytest tests/test_ocr_verify_policy.py` |
| S2 | Чистый extract + `short_side >= 1000` и остальные пороги OK → skip, `auto_all_checks_passed` | тот же файл |
| S3 | `short_side is None` → Verify, `auto_image_size_unknown` | тот же файл |
| S4 | `min_short_side=0` не форсит Verify на мелком кадре | тот же файл |
| S5 | `never` / `max_api_calls=1` глушат эвристику | тот же файл |
| S6 | Все 6 функций политики ведут себя одинаково на мелком кадре | unit на каждой или общий helper-тест + по одной на функцию |
| S7 | Пайплайн передаёт `short_side_px` в политику; лог `[OCR]` содержит его | mock-тест пайплайна (плиты + хотя бы сваи) |
| S8 | Регрессия: `pytest tests/test_ocr_verify_policy.py tests/test_recognition_pipeline.py tests/test_pile_ocr_pipeline.py tests/test_commercial_ocr_policy.py -q` — green | CI |
| S9 | Пилот 10–15 реальных плохих скринов: ошибки марок **или** пропуски строк не хуже baseline `auto` и хотя бы на части кейсов лучше; крупные чистые кадры не деградируют | ручной прогон, вне CI |

S9 решает, писать ли спеку этапа A. Если B не двигает ошибки — гипотеза «ломается пропуск Verify» опровергнута.

---

## Tech Stack

| Компонент | Технология |
|-----------|------------|
| Политика Verify | `core/ocr/verify_policy.py` |
| Пайплайн | `core/ocr/pipeline.py` (6 entrypoints) |
| Размер кадра | Pillow (`PIL.Image`), уже в зависимостях |
| Settings | `core/config/settings.py` → env `OCR_VERIFY_AUTO_MIN_SHORT_SIDE` |
| Провайдер | без изменений (GigaChat / OpenAI) |
| Frontend | без изменений |

Новых пакетов нет.

---

## Commands

```bash
# из корня репозитория, venv при необходимости
source .venv/bin/activate   # или venv/bin/activate

pytest tests/test_ocr_verify_policy.py -q
pytest tests/test_ocr_verify_policy.py tests/test_recognition_pipeline.py tests/test_pile_ocr_pipeline.py tests/test_commercial_ocr_policy.py -q

# backend (не обязателен для этой фичи)
uvicorn app.main:app --reload
```

Пилот (вне CI): те же 10–15 скринов до/после на staging с `OCR_VERIFY_MODE=auto`. Сравнить `ocr_verify_applied_reason`, правки менеджера, `ocr_api_calls`.

---

## Project Structure

```
core/ocr/verify_policy.py          # OcrVerifySettings.min_short_side + helper + 6 should_run_*
core/ocr/pipeline.py               # прочитать short_side, передать в политику, залогировать
core/ocr/providers/openai.py       # опционально: helper image_short_side_px рядом с load_image_payload
core/config/settings.py            # ocr_verify_auto_min_short_side
tests/test_ocr_verify_policy.py    # новые кейсы + поправить DEFAULT_SETTINGS / вызовы
tests/test_recognition_pipeline.py # если моки/вызовы политики разъедутся
tests/test_pile_ocr_pipeline.py    # то же
.env.example                       # документировать OCR_VERIFY_AUTO_MIN_SHORT_SIDE, если файл есть в репо
ai_docs/specs/ocr-small-screenshot-verify.md
```

Не трогаем: `frontend/`, промпты, parser gates, провайдеры extract/verify, upload validation.

---

## Code Style

- Политика тестируется без API и без файлов: на вход `short_side_px: int | None`.
- Не копировать новый `if` в шесть функций. Вынести общее решение по mode / bytes / short side.

```python
@dataclass(frozen=True)
class OcrVerifySettings:
    max_rows: int = 10
    min_confidence: float = 0.92
    max_bytes: int = 819_200
    min_short_side: int = 1000  # 0 = эвристика выключена


def _auto_image_decision(
    *,
    image_size_bytes: int,
    short_side_px: int | None,
    settings: OcrVerifySettings,
) -> tuple[bool, str] | None:
    """None — продолжить проверки строк. Иначе (run_verify, reason)."""
    if image_size_bytes > settings.max_bytes:
        return True, "auto_file_too_large"
    if short_side_px is None:
        return True, "auto_image_size_unknown"
    if settings.min_short_side > 0 and short_side_px < settings.min_short_side:
        return True, "auto_small_image"
    return None
```

Порядок в `should_run_*_verify` после `never` / `always` / unknown mode / empty rows:

1. `_auto_image_decision` (bytes → unknown size → small image)
2. `len(rows) > max_rows`
3. per-row confidence / parse / issues

Пайплайн: открыть изображение только для `size`, **не** ресайзить и не пересохранять. Нечитаемый кадр → `short_side_px=None`.

Лог: к существующему `[OCR] image_size_kb=...` добавить `short_side_px`.

---

## Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| Unit | S1–S6, все mode, порог 0, equality на 1000 | `tests/test_ocr_verify_policy.py` |
| Unit | существующие skip-тесты передают `short_side_px >= min_short_side` (иначе они начнут ждать Verify) | тот же файл |
| Integration (mock) | pipeline читает fixture PNG с известным размером и прокидывает short side | `tests/test_recognition_pipeline.py` и/или `tests/test_pile_ocr_pipeline.py` |
| Policy guard | `OCR_EXTERNAL_ENABLED` без регрессий | `tests/test_commercial_ocr_policy.py` |
| E2E manual | S9, 10–15 скринов | вне CI, без live API в тестах |

Покрытие: новая ветка политики 100% ветвей (`<`, `==`, `>`, `None`, `min_short_side=0`). Не гоняем реальный GigaChat/OpenAI в CI.

Существующие вызовы `should_run_*_verify` в тестах **обязаны** начать передавать `short_side_px` (явный kwarg, без «тихого default skip»). Default `None` = unknown = Verify — чтобы прод не забыл прокинуть размер.

---

## Boundaries

### Always

- Не менять пиксели, MIME, dpi metadata.
- Сохранять JSON-контракт OCR и ручную сверку менеджером.
- Покрывать новую ветку unit-тестами до merge.
- Логировать `short_side_px` и писать reason в metadata.
- `never` / `max_api_calls=1` важнее мелкого кадра.

### Ask first

- Менять default `1000` после пилота.
- Включать этап A (2× / контраст).
- Считать «мелким» огромный кадр с мелким шрифтом (высота глифа).
- Разные пороги по `product_type`.
- Выключать эвристику в prod (`=0`) без записи в эту спеку.

### Never

- Слепой апскейл / запись DPI «для точности».
- Автокроп, deskew, Real-ESRGAN / text SR.
- Отдельный table-OCR (Paddle, Azure DI).
- Третий LLM-вызов на одно фото.
- Live vision API в CI.
- Коммит секретов / живых БД.
- Telegram OCR.

---

## Configuration

```env
OCR_VERIFY_MODE=auto
OCR_MAX_API_CALLS=2
OCR_VERIFY_AUTO_MAX_ROWS=10
OCR_VERIFY_AUTO_MIN_CONFIDENCE=0.92
OCR_VERIFY_AUTO_MAX_BYTES=819200
OCR_VERIFY_AUTO_MIN_SHORT_SIDE=1000   # NEW; 0 = выкл
```

Поле в `Settings`: `ocr_verify_auto_min_short_side: int = Field(default=1000, alias="OCR_VERIFY_AUTO_MIN_SHORT_SIDE", ge=0)`.

Пайплайн кладёт его в `OcrVerifySettings.min_short_side`.

---

## Open Questions

- [ ] Default 1000 px ок, или сразу 800 / 1200?
- [ ] Все 6 типов сразу, или сначала плиты (остальные — тот же helper, но wiring позже)?
- [ ] Нечитаемый размер: Verify (как в D5) или skip эвристики?
- [ ] Есть ли уже пачка 10–15 плохих скринов для S9, или пилот после merge на staging?

---

## Out of scope (этап A и дальше)

- Условный 2× Lanczos + контраст.
- DPI, автокроп, нейросетевой SR, table-OCR.
- Смена copy wizard («вставьте текст»).
- Изменение `OCR_VERIFY_AUTO_MAX_BYTES` / `MAX_ROWS` (это другая калибровка; см. P2/P3 в pile-ocr spec).
