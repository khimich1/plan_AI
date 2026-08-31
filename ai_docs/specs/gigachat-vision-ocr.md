# Spec: GigaChat Vision OCR — распознавание списков плит

> **Тип:** feature-spec (SDD Phase: SPECIFY)  
> **Дата:** 2026-07-08  
> **Статус:** утверждено к реализации  
> **Источник идеи:** [`ai_docs/ideas/gigachat-vision-ocr.md`](../ideas/gigachat-vision-ocr.md)  
> **План реализации:** [`ai_docs/develop/plans/2026-07-08-gigachat-vision-ocr.md`](../develop/plans/2026-07-08-gigachat-vision-ocr.md)  
> **Связанные документы:** [`ux-wizard-step-plates.md`](./ux-wizard-step-plates.md), [`project-baseline.md`](./project-baseline.md)

---

## ASSUMPTIONS I'M MAKING

1. Основной канал — **веб wizard КП** (`PlateInputStep`); Telegram-бот вне scope.
2. API оплачивается как **физлицо** (`GIGACHAT_SCOPE=GIGACHAT_API_PERS`).
3. Модель по умолчанию — **GigaChat-2-Max** (Extract и Verify).
4. Менеджер **всегда** может проверить и исправить список вручную; цель — минимизировать правки, не убрать человека из контура.
5. Объём — **50–200 фото/мес**; latency одного фото **< 15 с** (p95).
6. **GPT-4o** остаётся как fallback (`OCR_PROVIDER=openai`) на период миграции.
7. Промпт и JSON-контракт из `core/plate_format_prompt.py` сохраняются; меняется только провайдер и пайплайн.

→ Поправь допущения до начала IMPLEMENT.

---

## Objective

Заменить распознавание списков ЖБ-плит с фото с **GPT-4o Vision** на **GigaChat 2 Max Vision** с **адаптивным числом API-вызовов**: на чётких небольших снимках — один вызов (Extract), на сложных — два (Extract + Verify).

### Проблема

- GPT-4o: оплата в USD, данные за рубежом, ошибки на похожих марках (`66,2` vs `66`, `6п` vs `8п`).
- Verify-этап в коде есть, но **отключён** в `recognize_text_smart`.
- Лишний Verify на каждом фото удваивает стоимость без выигрыша на простых таблицах.

### Пользователь

Менеджер по продажам — шаг 1 wizard КП: загружает фото таблицы плит, получает черновик списка, правит при необходимости, переходит к расчёту.

### User Stories

- Как **менеджер**, я загружаю чёткое фото короткой таблицы, чтобы система распознала плиты **быстро и дёшево** (1 вызов API).
- Как **менеджер**, я загружаю длинную или нечёткую таблицу, чтобы система **автоматически** провела вторую проверку по фото.
- Как **менеджер**, я вижу **corrections** и предупреждения (`ocr_verify_failed`, `parser_rejected`), чтобы понять, что сверить вручную.
- Как **оператор**, я могу задать режим verify (`auto` / `always` / `never`) и пороги через env без правки кода.

### Success Criteria (измеримые)

| # | Критерий | Метод проверки |
|---|----------|----------------|
| S1 | На пилоте 30–50 реальных фото ≥ **90%** строк без ручной правки марки/qty | Сравнение менеджером |
| S2 | В режиме `auto` ≥ **50%** фото обрабатываются **1 вызовом** без роста ошибок vs `always` | Логи `ocr_api_calls` |
| S3 | Средняя стоимость фото в `auto` **на 30%+ ниже**, чем при `always`, при сопоставимой точности | Логи `ocr_cost_rub` |
| S4 | `pytest tests/test_recognition_pipeline.py tests/test_commercial_ocr_policy.py -q` — green | CI |
| S5 | Публичный API `recognize_text_smart()` не ломает `commercial_draft_service` | Существующие тесты КП |

---

## Tech Stack

| Компонент | Технология |
|-----------|------------|
| Vision OCR (новый) | GigaChat API, модель `GigaChat-2-Max`, SDK `gigachat` |
| Vision OCR (fallback) | OpenAI `gpt-4o`, `openai` SDK |
| Парсер (локальный gate) | `core/plate_line_parser.parse_line` |
| Промпты | `core/plate_format_prompt.py` → `core/ocr/prompts.py` |
| Backend entry | `app/services/commercial_draft_service.py` |
| Frontend | `PlateInputStep.tsx` (без изменений в MVP, кроме новых metadata-полей при необходимости) |

### Зависимость (добавить)

```
gigachat>=0.1.39   # уточнить актуальную версию при IMPLEMENT
```

---

## Commands

```powershell
Set-Location "c:\Users\Роман\Desktop\Шишов"
.\.venv\Scripts\Activate.ps1

# Тесты OCR-пайплайна
pytest tests/test_recognition_pipeline.py tests/test_commercial_ocr_policy.py -q

# Тесты КП (регрессия wizard)
pytest tests/test_commercial_web_flow.py -q

# Backend dev
uvicorn app.main:app --reload

# Пилот на реальном фото (ручной скрипт — добавить при IMPLEMENT)
# python scripts/ocr_pilot_compare.py --image path/to/photo.jpg --provider gigachat
```

---

## Project Structure

```
core/ocr/                          # NEW — OCR subsystem
  __init__.py                      # recognize_text_smart, apply_plates_with_ai (re-export)
  pipeline.py                      # extract → parser_gate → verify_policy → verify?
  verify_policy.py                 # should_run_verify()
  parser_gate.py                   # apply_parser_gate()
  prompts.py                       # system + verify prompts
  parsing.py                       # parse_gpt_response, parse_verify_response (из ocr_gpt)
  providers/
    base.py                        # Protocol OcrProvider
    gigachat.py                    # GigaChat Vision
    openai.py                      # текущий GPT-4o (из ocr_gpt)

core/ocr_gpt.py                    # DEPRECATED shim → core.ocr (обратная совместимость импортов)
core/plate_format_prompt.py        # без изменений (источник промпта)
core/config/settings.py            # новые поля OCR/GigaChat
app/services/commercial_draft_service.py  # без изменений импорта (через shim)

tests/
  test_recognition_pipeline.py     # обновить моки под провайдер
  test_ocr_verify_policy.py        # NEW — unit-тесты эвристик auto
  test_ocr_gigachat_provider.py    # NEW — мок SDK
```

---

## Functional Design

### Пайплайн

```
Фото
  → Extract (GigaChat-2-Max)
  → Parser Gate (plate_line_parser, 0 токенов)
  → verify_policy.should_run_verify()
       ├─ skip → результат (ocr_api_calls=1)
       └─ run  → Verify (GigaChat-2-Max) → Parser Gate → результат (ocr_api_calls=2)
```

### Parser Gate

Для каждой строки после Extract/Verify:

1. Вызвать `parse_line(f"{normalized_candidate} {qty}")`.
2. Если `parsed == false` → `issues += ["parser_rejected"]`, `confidence = min(confidence, 0.5)`.

### Режимы verify (`OCR_VERIFY_MODE`)

| Значение | Поведение |
|----------|-----------|
| `never` | Только Extract (1 вызов) |
| `always` | Extract + Verify (2 вызова) |
| `auto` | Verify по эвристикам ниже |

Жёсткий потолок: `OCR_MAX_API_CALLS` ∈ `{1, 2}`.

### Эвристики `auto` — Verify **пропускается**, если **все** true

| Условие | Env | Default |
|---------|-----|---------|
| Размер файла ≤ порога | `OCR_VERIFY_AUTO_MAX_BYTES` | `819200` (800 КБ) |
| Число строк ≤ порога | `OCR_VERIFY_AUTO_MAX_ROWS` | `15` |
| У каждой строки `confidence` ≥ порога | `OCR_VERIFY_AUTO_MIN_CONFIDENCE` | `0.92` |
| Parser Gate: все строки `parsed` | — | — |
| `issues` пуст у всех строк | — | — |
| Extract вернул непустой список | — | — |

Verify **всегда** при любом нарушении выше или при `OCR_VERIFY_MODE=always`.

`verify_policy.should_run_verify()` возвращает `(run: bool, reason: str)` — `reason` пишется в metadata (`ocr_verify_skipped_reason` или `ocr_verify_applied_reason`).

### Ответ `recognize_text_smart()` — расширенные поля

Сохранить существующие поля + добавить:

| Поле | Тип | Описание |
|------|-----|----------|
| `ocr_api_calls` | `int` | 1 или 2 |
| `ocr_cost_rub` | `float` | Суммарная стоимость в ₽ |
| `ocr_cost_usd` | `float` | Deprecated; 0 для GigaChat |
| `ocr_verify_skipped_reason` | `str \| null` | Почему verify не вызывался |
| `ocr_method` | `str` | Напр. `GigaChat-2-Max`, `GigaChat-2-Max+verify` |

### Провайдер GigaChat

- Scope: `GIGACHAT_API_PERS`
- Credentials: `GIGACHAT_CREDENTIALS` (не коммитить)
- SDK синхронный → `asyncio.to_thread()` в async-пайплайне
- Изображение: `data:{mime};base64,...` в `ChatContentPart` (как в документации SDK)
- `temperature=0`, `max_tokens=2500`

### Тарификация (физлицо)

| Модель | ₽ / 1000 токенов | Пакет |
|--------|------------------|-------|
| GigaChat 2 Max | 0,65 | 3 млн = **1 950 ₽** / 12 мес. |
| Freemium | — | 50 000 токенов Max / 12 мес. |

Оценка на фото:

| Вызовов | Токенов | ₽ (Max) |
|---------|---------|---------|
| 1 | ~4 000–6 000 | 2,6–4 |
| 2 | ~8 000–12 000 | 5–8 |

При 50–200 фото/мес и `auto` (~35% verify): **~200–1 400 ₽/мес**.

---

## Configuration

```env
# Провайдер
OCR_PROVIDER=gigachat              # gigachat | openai
OCR_EXTERNAL_ENABLED=true
GIGACHAT_CREDENTIALS=              # ключ из Studio, не коммитить
GIGACHAT_MODEL=GigaChat-2-Max
GIGACHAT_SCOPE=GIGACHAT_API_PERS

# Адаптивный verify
OCR_VERIFY_MODE=auto               # auto | always | never
OCR_MAX_API_CALLS=2                # 1 | 2
OCR_VERIFY_AUTO_MAX_ROWS=15
OCR_VERIFY_AUTO_MIN_CONFIDENCE=0.92
OCR_VERIFY_AUTO_MAX_BYTES=819200

# Fallback (если OCR_PROVIDER=openai)
OPENAI_API_KEY=
```

Поля добавить в `core/config/settings.py` с Pydantic alias.

---

## Code Style

- Провайдеры — через `Protocol` в `providers/base.py`, без наследования от конкретных SDK.
- Бизнес-логика verify **только** в `verify_policy.py` (тестируемо без моков API).
- Логирование: `[OCR]` prefix, INFO для `api_calls`, `cost_rub`, `verify_decision`.
- Обратная совместимость: `from core.ocr_gpt import recognize_text_smart` работает через shim.

Пример `verify_policy`:

```python
def should_run_verify(
    *,
    mode: str,
    max_api_calls: int,
    image_size_bytes: int,
    plates: list[dict],
    settings: OcrVerifySettings,
) -> tuple[bool, str]:
    if max_api_calls <= 1 or mode == "never":
        return False, "max_api_calls_or_never"
    if mode == "always":
        return True, "mode_always"
    # auto: проверка порогов...
    return False, "auto_all_checks_passed"
```

---

## Testing Strategy

| Уровень | Что | Файлы |
|---------|-----|-------|
| Unit | `should_run_verify` — все комбинации mode/порогов | `test_ocr_verify_policy.py` |
| Unit | `apply_parser_gate` — parser_rejected, confidence | `test_ocr_parser_gate.py` |
| Unit | parsing JSON ответов | существующие в `test_recognition_pipeline.py` |
| Integration (mock) | pipeline 1 vs 2 вызова | `test_recognition_pipeline.py` |
| Integration (mock) | GigaChat provider | `test_ocr_gigachat_provider.py` |
| Policy | `OCR_EXTERNAL_ENABLED` guard | `test_commercial_ocr_policy.py` |
| E2E manual | 30–50 реальных фото | пилот, вне CI |

**Не вызывать реальный GigaChat API в CI.**

```bash
pytest tests/test_recognition_pipeline.py tests/test_ocr_verify_policy.py tests/test_commercial_ocr_policy.py -q
```

---

## Boundaries

### Always

- Сохранять JSON-контракт плит (`raw_name`, `normalized_candidate`, `qty`, `confidence`, `issues`).
- `temperature=0` для OCR.
- Не коммитить `GIGACHAT_CREDENTIALS` / `OPENAI_API_KEY`.
- Покрывать `verify_policy` unit-тестами.

### Ask first

- Удаление `core/ocr_gpt.py` shim (после миграции всех импортов).
- Изменение порогов `auto` по умолчанию после пилота.
- Переход на юрлицо (B2B scope).

### Never

- GigaChat Lite для vision OCR плит.
- Более 2 LLM-вызовов на одно фото в MVP.
- Три и более провайдеров одновременно в проде.
- Убирать ручную проверку менеджером из UX.

---

## Implementation Plan

| # | Задача | Зависимости | Verify |
|---|--------|-------------|--------|
| 1 | `core/ocr/providers/openai.py` — вынести текущий GPT-код | — | pytest recognition |
| 2 | `core/ocr/parsing.py`, `prompts.py` — перенос из `ocr_gpt.py` | 1 | pytest parsing |
| 3 | `parser_gate.py` + тесты | — | pytest parser_gate |
| 4 | `verify_policy.py` + тесты | — | pytest verify_policy |
| 5 | `providers/gigachat.py` + settings + requirements | 2 | mock tests |
| 6 | `pipeline.py` — wiring extract/verify/auto | 1–5 | pytest pipeline |
| 7 | Shim `core/ocr_gpt.py` → `core.ocr` | 6 | import tests |
| 8 | `.env.example`, metadata в draft | 6 | manual |
| 9 | Пилот 30–50 фото, калибровка порогов | 8 | S1–S3 |

---

## Not Doing (MVP)

| Исключено | Причина |
|-----------|---------|
| Препроцессинг изображений | Фаза 2; сначала провайдер + verify |
| Pro + Max hybrid | Сложность; Max на обоих этапах |
| GigaChat async mode | Медленно для интерактивного wizard |
| Telegram-бот | `bot_archived`, отдельная задача |
| A/B GPT vs GigaChat в проде | Только ручной пилот |
| Юрлицо / договор | Выбрано физлицо |
| Метрика резкости (Laplacian) | Фаза 2 для `auto` |

---

## Open Questions

| # | Вопрос | Решение по умолчанию |
|---|--------|----------------------|
| Q1 | Freemium физлица покрывает коммерческое использование завода? | Уточнить в ToS Studio; пилот на freemium |
| Q2 | Staging = `always`, prod = `auto`? | Да, через env |
| Q3 | Логировать статистику 1 vs 2 вызовов? | Да, structured log |
| Q4 | При `OCR_MAX_API_CALLS=1` и сомнительном Extract — warning в UI? | Да, если `parser_rejected` или low confidence |

---

## Tasks (IMPLEMENT checklist)

- [ ] **T1:** Создать `core/ocr/` package и OpenAI provider  
  - Acceptance: `recognize_text_smart` работает через новый пакет с `OCR_PROVIDER=openai`  
  - Files: `core/ocr/**`, shim `core/ocr_gpt.py`  
  - Verify: `pytest tests/test_recognition_pipeline.py -q`

- [ ] **T2:** `verify_policy` + `parser_gate`  
  - Acceptance: unit-тесты покрывают `auto`/`always`/`never` и parser_rejected  
  - Files: `verify_policy.py`, `parser_gate.py`, `tests/test_ocr_verify_policy.py`  
  - Verify: `pytest tests/test_ocr_verify_policy.py -q`

- [ ] **T3:** GigaChat provider + settings  
  - Acceptance: `OCR_PROVIDER=gigachat` вызывает mock SDK, считает `cost_rub`  
  - Files: `providers/gigachat.py`, `settings.py`, `requirements.txt`, `.env.example`  
  - Verify: `pytest tests/test_ocr_gigachat_provider.py -q`

- [ ] **T4:** Pipeline wiring + metadata  
  - Acceptance: `ocr_api_calls`, `ocr_verify_skipped_reason`, `ocr_cost_rub` в ответе draft  
  - Files: `pipeline.py`, `commercial_draft_service.py` (metadata mapping)  
  - Verify: `pytest tests/test_commercial_web_flow.py -q`

- [ ] **T5:** Пилот на реальных фото  
  - Acceptance: S1–S3 подтверждены менеджером; пороги `auto` зафиксированы  
  - Verify: отчёт в `ai_docs/develop/reports/`
