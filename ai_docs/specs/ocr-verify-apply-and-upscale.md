# Spec: OCR — Apply Verify только при corrections + 2× на мелких

> **Тип:** feature-spec (SDD Phase: SPECIFY ✅ → PLAN ✅ → IMPLEMENT ✅)  
> **Дата:** 2026-08-22  
> **Статус:** код AU-201…209 влит в рабочее дерево; пилот S9 (AU-210) вручную  
> **План:** [`ai_docs/develop/plans/2026-08-22-ocr-verify-apply-and-upscale.md`](../develop/plans/2026-08-22-ocr-verify-apply-and-upscale.md)  
> **Источник идеи:** [`ai_docs/ideas/ocr-verify-apply-and-upscale.md`](../ideas/ocr-verify-apply-and-upscale.md)  
> **Handoff:** [`ai_docs/develop/handoffs/2026-08-22-ocr-verify-apply-and-upscale.md`](../develop/handoffs/2026-08-22-ocr-verify-apply-and-upscale.md)  
> **Предшественник (этап B, в working tree):** [`ocr-small-screenshot-verify.md`](./ocr-small-screenshot-verify.md)  
> **Связанные документы:** [`gigachat-vision-ocr.md`](./gigachat-vision-ocr.md), [`pile-ocr-reliability.md`](./pile-ocr-reliability.md)

---

## ASSUMPTIONS I'M MAKING

1. Этап B (`auto_small_image`, `image_short_side_px`, `OCR_VERIFY_AUTO_MIN_SHORT_SIDE`) **остаётся**. Эту спеку не откатываем и не выключаем второй вызов.
2. Канал — **веб wizard КП**. Telegram OCR вне scope. Путь «ИИ-инструкция» (`apply_plates_with_ai` и аналоги) **не** трогаем.
3. Регресс пилота — **подмена `plates`**, не UI-строка про омоглифы. Серверный лог того прогона: `corrections=0`, `plates = parser_gate(verified)` безусловно.
4. `corrections=[]` (или отсутствует), даже если `plates` Verify **другие** и даже если `row_count_on_image ≠ len(extract)` → оставить Extract. **Не** синтезировать add/remove.
5. Контраст по умолчанию — `PIL.ImageOps.autocontrast(..., cutoff=1)`. Без бинаризации, без unsharp. Смена cutoff / линейный contrast — Ask first.
6. Порог «мелкий» для **и** Verify, **и** 2× — один и тот же: короткая сторона **исходника** `< OCR_VERIFY_AUTO_MIN_SHORT_SIDE` (default 1000). `0` выключает оба. Равенство порогу **не** триггерит. После 2× (416→832) Verify всё ещё идёт.
7. Политика Verify (`should_run_*`, `image_size_bytes`, `short_side_px`) считает **исходник**. В провайдер уходят уже обработанные байты (PNG). Новый env не заводим.
8. Нечитаемый кадр / PDF (`short_side_px is None`) → препроцесс **no-op**, политика B без изменений (`auto_image_size_unknown` → Verify).
9. Metadata для пилота S9: `ocr_preprocess` (`null` | `2x_lanczos`) и `ocr_verify_select_reason`. `ocr_verify_applied_reason` **не** затираем — там остаётся причина *запуска* Verify (`auto_small_image` и т.д.).
10. Новых зависимостей нет (Pillow уже есть). UI wizard, промпты Extract/Verify, parser_gate, upload validation, схема БД — без изменений в MVP. Долг тестов этапа B **не** входит в этот PR, пока пользователь явно не попросит.
11. Обрезанную 10-ю строку не обещаем восстановить. Таблица омоглифов `u`→`и` — out. 2× PNG типичного скрина < 1000 px не упрётся в лимит GigaChat; если encode даст файл **> 8 MiB** — отправить исходник и залогировать skip.

→ Допущения приняты (2026-08-22). PLAN: не код до «план ок».

---

## Decisions locked

| # | Тема | Решение |
|---|------|---------|
| D1 | Этап B | **Оставляем.** Не откат, не `min_short_side=0` «чтобы починить хвост» |
| D2 | Apply-policy | Verify-список только если `verified` не пустой **и** `corrections` не пустой. Иначе Extract |
| D3 | Не угадывать diff | Разные plates / другой `row_count_on_image` при пустом `corrections` ≠ основание подменить список |
| D4 | Reason запуска vs выбора | `ocr_verify_applied_reason` = почему вызвали Verify. `ocr_verify_select_reason` = что взяли (`applied` / `kept_extract_empty_corrections` / `empty_verified_plates`) |
| D5 | `verify_applied` / `verify_failed` | Вызов состоялся → `verify_applied=true`. `verify_failed=true` только пустой `verified` или exception. Пустые corrections — **не** fail |
| D6 | Препроцесс | Исходник `short < min_short_side` → RGB → 2× `LANCZOS` → `autocontrast(cutoff=1)` → PNG в API. Диск исходника не писать |
| D7 | Метрика «мелкий» | Только исходник. 2× не двигает порог Verify |
| D8 | Байты для политики | `image_size_bytes` / `auto_file_too_large` — **исходный** файл, не раздутый PNG |
| D9 | Общий код | Один `select_ocr_items` + один `preprocess_image_for_ocr` на все 6 runners. Не 6 копий if |
| D10 | Env | Новых переменных нет. Тот же `OCR_VERIFY_AUTO_MIN_SHORT_SIDE` |
| D11 | UI / промпты / БД | Не меняем. Новые ключи только в OCR payload + mapper черновика (без copy) |
| D12 | Out | Омоглифы, автокроп, DPI, SR, отказ «переснимите», разные пороги по типу, откат B |

---

## Objective

Не давать второму проходу по мелкому скрину **молча заменить** более верный Extract, и дать обоим проходам чуть больше пикселей на символ — без DPI, без кропа, без обещания восстановить обрезанные строки.

### Проблема

Этап B на пилотном Excel-скрине (`short_side_px=416`, `verify_decision=auto_small_image`, 2 вызова) отработал как задумано: Verify включился. Verify вернул `rows=10`, `corrections=0`. Пайплайн всё равно сделал `plates = parser_gate(verified_plates)`. Тот же файл **до B** (1 вызов) читался лучше: хвост `ПБ36` → `ПБ 63`, кириллическая **и** → латинская **u**. Обрезанную последнюю строку ни апскейл, ни Verify не вернут.

Два независимых механизма, оба нужны:

```
исходник
  ├── image_short_side_px(original)     → политика auto_small_image (этап B, без изменений)
  └── preprocess if short < N            → 2× Lanczos + autocontrast → PNG
           │
           ├── Extract
           └── Verify? (B, порог по исходнику)
                    │
                    └── select_ocr_items(extract, verify_result, gate)
                         ├── no/empty plates     → extract, verify_failed
                         ├── empty corrections   → extract, kept_extract
                         └── else                → parser_gate(verified)
```

B без правила apply на этом фото вреден. 2× без правила apply снова отдаст Verify право затереть Extract.

### Пользователь

Менеджер по продажам на шаге ввода состава КП: загружает скрин таблицы, сверяет черновик, жмёт «Список верен».

### User Stories

- Как **менеджер**, я загружаю мелкий скрин и не теряю верные марки Extract, если второй проход не назвал ни одной правки.
- Как **менеджер**, я всё ещё получаю правки Verify, когда модель **явно** вернула непустой `corrections` (и непустой список).
- Как **менеджер**, на мелком кадре оба прохода смотрят чуть более крупную сетку (2×), без смены UX.
- Как **оператор**, в metadata/логах я вижу и `auto_small_image` (зачем Verify), и `kept_extract_empty_corrections` (что взяли), и `ocr_preprocess=2x_lanczos`.
- Как **оператор**, исходный файл на диске / во временном upload не перезаписывается.

### Success Criteria (измеримые)

| # | Критерий | Метод проверки |
|---|----------|----------------|
| S1 | `corrections=[]` + другой `plates` → в результате Extract, `verify_applied=true`, `verify_failed=false`, `ocr_verify_select_reason=kept_extract_empty_corrections` | unit `select_ocr_items` + mock pipeline (плиты) |
| S2 | непустой `corrections` + непустой `plates` → результат = `parser_gate(verified)`, `select_reason=applied` | тот же pack |
| S3 | пустой `verified` (`plates`/`[]`) → extract, `verify_failed=true`, `select_reason=empty_verified_plates` | unit + существующий fail-path |
| S4 | `never` / `max_api_calls=1` — без Verify, без подмены, `select_reason` нет / не пишем | существующие skip-тесты |
| S5 | исходник `short >= min_short_side` (и `min_short_side>0`) → препроцесс no-op, в провайдер исходные байты, `ocr_preprocess is None` | `tests/test_ocr_image_preprocess.py` |
| S6 | исходник `short < min_short_side` → в провайдер PNG (`image/png`), размер 2×, исходный файл на диске байт-в-байт тот же | тот же файл + mock pipeline проверяет mime/bytes провайдера |
| S7 | все 6 runners зовут один `select_ocr_items` и один `preprocess_image_for_ocr` (не копипаста if) | code review + минимум плиты + сваи в mock pipeline |
| S8 | pytest pack этапа B + новые тесты — green; без live GigaChat/OpenAI в CI | команда ниже |
| S9 | тот же скрин 416 px: хвост марок **не хуже**, чем Extract-only до B; `corrections=0` не затирает Extract; `ocr_preprocess=2x_lanczos` | ручной `scripts/ocr_pilot_compare.py`, вне CI |

S9 не обещает восстановить обрезанную строку и не требует, чтобы 2× исправил `36`/`63`. Минимум — нет регресса подменой. Улучшение хвоста за счёт 2× — наблюдаемый бонус, не gate.

---

## Tech Stack

| Компонент | Технология |
|-----------|------------|
| Apply-policy | новый `core/ocr/verify_apply.py` (`select_ocr_items`) |
| Препроцесс | новый `core/ocr/image_preprocess.py` (Pillow: `Image.Resampling.LANCZOS`, `ImageOps.autocontrast`) |
| Размер исходника | существующий `core/ocr/image_meta.py` (только чтение `size`) |
| Политика «когда Verify» | `core/ocr/verify_policy.py` — **без обязательных правок** |
| Пайплайн | `core/ocr/pipeline.py` (6 entrypoints) |
| Payload | `core/ocr/result.py` → `build_result_payload` |
| Черновик КП | `commercial_draft_service._map_ocr_result_metadata` + optional поля в `CommercialDraftMetadata` |
| Settings | тот же `ocr_verify_auto_min_short_side` |
| Провайдеры / промпты / frontend UI | без изменений |

Новых пакетов нет.

---

## Commands

```bash
# из корня репозитория
source .venv/bin/activate   # или venv/bin/activate

pytest tests/test_ocr_verify_apply.py tests/test_ocr_image_preprocess.py -q
pytest tests/test_ocr_verify_policy.py tests/test_ocr_image_meta.py \
  tests/test_recognition_pipeline.py tests/test_pile_ocr_pipeline.py \
  tests/test_ocr_verify_apply.py tests/test_ocr_image_preprocess.py \
  tests/test_commercial_ocr_policy.py -q

# backend не обязателен для этой фичи
uvicorn app.main:app --reload
```

Пилот S9 (вне CI, live credentials):

```bash
python scripts/ocr_pilot_compare.py --image path/to/pilot-416.png --verify-mode auto
```

Сравнить текст хвоста с прогоном Extract-only до B; в выводе должны быть `ocr_preprocess` и `ocr_verify_select_reason`. Скрипт пилота дописать печать этих полей (без смены CLI).

---

## Project Structure

```
core/ocr/verify_apply.py                 # NEW: select_ocr_items
core/ocr/image_preprocess.py             # NEW: preprocess_image_for_ocr
core/ocr/image_meta.py                   # без изменения контракта (только size)
core/ocr/pipeline.py                     # 6×: preprocess → extract/verify на payload; select после verify
core/ocr/result.py                       # прокинуть select_reason + preprocess
core/ocr/verify_policy.py                # не обязан меняться
core/config/settings.py                  # без новых полей
app/services/commercial_draft_service.py # mapper: ocr_verify_select_reason, ocr_preprocess
app/schemas/commercial.py                # optional str | None поля, default None
scripts/ocr_pilot_compare.py             # печать новых полей
tests/test_ocr_verify_apply.py           # NEW: S1–S3, row_count mismatch
tests/test_ocr_image_preprocess.py       # NEW: S5–S6, None size, min_short_side=0, no rewrite
tests/test_recognition_pipeline.py       # поправить моки: непустой corrections, если ждут подмену
tests/test_pile_ocr_pipeline.py          # то же (_mock_provider сегодня всегда corrections=[])
tests/test_ocr_verify_policy.py          # регрессия B, без новых веток
ai_docs/specs/ocr-verify-apply-and-upscale.md
```

Не трогаем: `frontend/` UI/copy, `core/ocr/prompts.py`, parser gates, провайдеры extract/verify, `commercial_upload_validation`, Telegram, путь AI-инструкции.

---

## Code Style

Политика apply тестируется **без API и без файлов**. Препроцесс — на fixture PNG в `tmp_path`, исходник не мутировать.

```python
@dataclass(frozen=True)
class OcrSelectDecision:
    items: list[dict[str, Any]]
    verify_failed: bool
    select_reason: str  # applied | kept_extract_empty_corrections | empty_verified_plates


def select_ocr_items(
    extract_items: list[dict[str, Any]],
    verify_result: dict[str, Any],
    apply_gate: Callable[[list[dict[str, Any]]], list[dict[str, Any]]],
) -> OcrSelectDecision:
    verified = verify_result.get("plates") or []
    corrections = verify_result.get("corrections") or []
    if not verified:
        return OcrSelectDecision(extract_items, True, "empty_verified_plates")
    if not corrections:
        return OcrSelectDecision(extract_items, False, "kept_extract_empty_corrections")
    return OcrSelectDecision(apply_gate(verified), False, "applied")
```

Ключ Verify-списка везде `"plates"` — как сейчас в шести runners (сваи/ступени тоже кладут ответ Verify в `plates`).

`corrections` «непустой» = `bool(list)` после `or []`. Не фильтруем «настоящие» vs мусорные элементы. Не сравниваем plates поэлементно.

```python
@dataclass(frozen=True)
class OcrPreprocessResult:
    image_data: bytes
    mime_type: str  # "image/png"
    applied: bool   # True → 2x_lanczos


def preprocess_image_for_ocr(
    image_path: str,
    *,
    min_short_side: int,
) -> OcrPreprocessResult | None:
    """None — не удалось прочитать кадр (как short_side None): пайплайн шлёт исходник.
    applied=False — порог не взят / min_short_side==0 / PNG>8MiB: исходные байты.
    """
```

Порядок в каждом runner:

1. `short_side_px = image_short_side_px(path)` — исходник.
2. Прочитать исходные байты (для политики и fallback).
3. `preprocess_image_for_ocr(path, min_short_side=verify_settings.min_short_side)`.
4. В Extract/Verify — payload препроцесса, если `applied`; иначе исходник.
5. `should_run_*_verify(..., image_size_bytes=len(original_bytes), short_side_px=short_side_px)`.
6. После Verify — `select_ocr_items(draft, verify_result, apply_*_parser_gate)`.

Лог `[OCR]`: к существующим `short_side_px` / `verify_decision` добавить `preprocess=` и `select=`.

`image_meta.py` не начинает ресайзить. Препроцесс — sibling.

---

## Testing Strategy

| Уровень | Что | Где |
|---------|-----|-----|
| Unit | S1–S3, в т.ч. другой `row_count_on_image` + пустой corrections → extract | `tests/test_ocr_verify_apply.py` |
| Unit | S5–S6; `short == min_short_side` no-op; `min_short_side=0` no-op; PDF/битые байты → None/no-op; файл не перезаписан; RGB/P/RGBA не падают | `tests/test_ocr_image_preprocess.py` |
| Integration (mock) | плиты + сваи: silent rewrite не применяется; непустой diff применяется; провайдер на мелком кадре получает `image/png` | `tests/test_recognition_pipeline.py`, `tests/test_pile_ocr_pipeline.py` |
| Регрессия B | pack политики short-side | `tests/test_ocr_verify_policy.py`, `tests/test_ocr_image_meta.py` |
| Policy guard | `OCR_EXTERNAL_ENABLED` | `tests/test_commercial_ocr_policy.py` |
| E2E manual | S9, тот же 416 px скрин | `ocr_pilot_compare.py`, без live API в pytest |

Покрытие apply: 100% трёх веток. Покрытие preprocess: `<` / `==` / `>` / `0` / unreadable.

**Ловушка существующих моков:** `_mock_provider` в `test_pile_ocr_pipeline.py` всегда отдаёт `corrections=[]`. Тест `test_pile_pipeline_runs_verify_on_parser_rejected` сегодня ждёт, что Verify **подменит** `???` на хорошие сваи. После S1 этот тест обязан передать непустой `corrections`, иначе останется extract — и это правильное новое поведение, не баг теста.

Не гоняем реальный GigaChat/OpenAI в CI.

---

## Boundaries

### Always

- Оставлять этап B включённым (default порог 1000).
- Применять Verify-список только при непустых `plates` **и** непустых `corrections`.
- Считать «мелкий» по исходнику; не перезаписывать исходный файл.
- Один helper apply + один helper preprocess на 6 пайплайнов.
- Покрыть S1–S8 pytest до merge.
- Логировать `short_side_px`, `ocr_preprocess`, `ocr_verify_applied_reason`, `ocr_verify_select_reason`.
- `never` / `max_api_calls=1` важнее и Verify, и (для вызова) — препроцесс всё равно может сработать на Extract.

### Ask first

- Менять default 1000 / cutoff autocontrast / линейный contrast.
- Жёсткий лимит PNG ≠ 8 MiB или отдельный env на препроцесс.
- Закрывать долг тестов этапа B в том же PR.
- Разные пороги по `product_type`.
- Показывать `ocr_preprocess` / select reason в UI wizard.
- Менять промпт Verify (чтобы модель чаще заполняла `corrections`).
- Включать препроцесс на PDF (рендер страницы) — отдельная спека.

### Never

- Откатывать `auto_small_image` «чтобы хвост стал как вчера».
- Синтезировать `corrections` из diff plates.
- Автокроп, deskew, Real-ESRGAN / text SR, запись DPI «для точности».
- Третий LLM-вызов, table-OCR (Paddle, Azure DI).
- Таблица омоглифов `u`→`и`, отказ «переснимите».
- Обещать восстановить обрезанную строку.
- Live vision API в CI.
- Коммит секретов / живых БД / этапа B «заодно» без просьбы.
- Telegram OCR.

---

## Configuration

Новых переменных нет.

```env
OCR_VERIFY_MODE=auto
OCR_MAX_API_CALLS=2
OCR_VERIFY_AUTO_MAX_ROWS=10
OCR_VERIFY_AUTO_MIN_CONFIDENCE=0.92
OCR_VERIFY_AUTO_MAX_BYTES=819200
OCR_VERIFY_AUTO_MIN_SHORT_SIDE=1000   # уже есть; 0 = нет auto_small_image и нет 2×
```

`min_short_side=0`: эвристика B выкл **и** препроцесс выкл (`short < 0` никогда). Verify по-прежнему может включиться по rows/confidence/issues/`always`.

---

## Metadata (контракт payload)

Добавить в `build_result_payload` и прокинуть в mapper черновика (UI не рисует):

| Поле | Значения | Смысл |
|------|----------|--------|
| `ocr_verify_applied_reason` | как сейчас (`auto_small_image`, …) | почему вызвали Verify |
| `ocr_verify_select_reason` | `applied` / `kept_extract_empty_corrections` / `empty_verified_plates` / `None` если Verify не вызывали | что попало в `plates` |
| `ocr_preprocess` | `2x_lanczos` / `None` | что ушло в API |

`verify_applied` / `verify_failed` / `corrections` / `draft_plates` — семантика D5. `corrections` из ответа модели **сохраняем как есть** даже когда список не применяем (пустой массив на пилотном регрессе).

Frontend types / copy **не** обязательны в MVP: лишние ключи API менеджер не видит.

---

## Open Questions

- [ ] `corrections=[]` + другой `row_count_on_image` → всё равно Extract? (**в спеке: да**, A4)
- [ ] Autocontrast `cutoff=1` ок, или сразу 0 / линейный contrast? (**в спеке: cutoff=1**)
- [ ] Писать `ocr_preprocess` + `ocr_verify_select_reason` в draft metadata? (**в спеке: да**, без UI)
- [ ] Закрыть долг тестов B (Settings `=0`, PDF → `auto_image_size_unknown`, пин `OCR_VERIFY_MODE=auto`) **в том же PR**? (**в спеке: нет**, Ask first)
- [ ] Есть ли тот же файл 416 px под рукой для S9, или путь уточним в PLAN?

---

## Out of scope

- Откат этапа B.
- Омоглифы, автокроп, DPI, нейросетевой SR, отказ переснять.
- Смена copy wizard и промптов Extract/Verify.
- Рендер PDF в растр.
- Калибровка `OCR_VERIFY_AUTO_MAX_BYTES` / `MAX_ROWS`.
- Коммит uncommitted этапа B без явной просьбы.
