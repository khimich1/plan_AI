# Plan: GigaChat Vision OCR

**Created:** 2026-07-08  
**Orchestration:** `orch-2026-07-08-17-01-gigachat-vision-ocr`  
**Status:** 🟡 OCR-009 pending (manual pilot)  
**Spec:** [`ai_docs/specs/gigachat-vision-ocr.md`](../../specs/gigachat-vision-ocr.md)  
**Idea:** [`ai_docs/ideas/gigachat-vision-ocr.md`](../../ideas/gigachat-vision-ocr.md)

## Goal

Заменить GPT-4o Vision на **GigaChat 2 Max** с **адаптивным verify** (1 или 2 API-вызова на фото) для распознавания списков ЖБ-плит в wizard КП. Сохранить обратную совместимость `recognize_text_smart()` и fallback на OpenAI.

## Текущее состояние кодовой базы

| Компонент | Статус |
|-----------|--------|
| `core/ocr_gpt.py` | Shim → `core/ocr/` |
| `core/ocr/` | ✅ Пакет: providers, pipeline, gate, policy |
| `core/config/settings.py` | ✅ `OCR_VERIFY_MODE`, GigaChat fields |
| `app/services/commercial_draft_service.py` | ✅ OCR metadata + Q4 warnings |
| `tests/test_recognition_pipeline.py` | Моки `core.ocr_gpt.AsyncOpenAI` |
| `tests/test_commercial_ocr_policy.py` | Guard `OCR_EXTERNAL_ENABLED` |
| `requirements.txt` | `openai>=1.0.0`; `gigachat` — **добавить** |

## Architecture Decisions

- **Провайдеры через Protocol** (`providers/base.py`) — без наследования от SDK.
- **Verify policy** — единственная точка решения в `verify_policy.py`; тестируется без моков API.
- **Parser Gate** — локально через `core.plate_line_parser.parse_line`, 0 токенов.
- **Shim** — `core/ocr_gpt.py` реэкспортирует из `core.ocr` для обратной совместимости.
- **Синхронный GigaChat SDK** — обёртка `asyncio.to_thread()` в async-пайплайне.
- **Staging = `always`, prod = `auto`** — через env, без правки кода.

## Task List

### Phase 1: Foundation (рефакторинг без смены провайдера)

- [x] **OCR-001:** Создать `core/ocr/` package и вынести OpenAI provider (✅ Done)
- [x] **OCR-002:** Перенести parsing и prompts из `ocr_gpt.py` (✅ Done)

**Checkpoint: Foundation**
- [x] `pytest tests/test_recognition_pipeline.py -q` — green с `OCR_PROVIDER=openai`
- [x] `from core.ocr_gpt import recognize_text_smart` работает через shim

### Phase 2: Локальная логика (без API)

- [x] **OCR-003:** `parser_gate.py` + unit-тесты (✅ Done)
- [x] **OCR-004:** `verify_policy.py` + unit-тесты (✅ Done)

**Checkpoint: Local logic**
- [x] `pytest tests/test_ocr_parser_gate.py tests/test_ocr_verify_policy.py -q` — green

### Phase 3: GigaChat + конфигурация

- [x] **OCR-005:** GigaChat provider + settings + зависимости (✅ Done)
- [x] **OCR-006:** `pipeline.py` — wiring extract → gate → verify → metadata (✅ Done)

**Checkpoint: Pipeline**
- [x] `pytest tests/test_recognition_pipeline.py tests/test_ocr_gigachat_provider.py tests/test_commercial_ocr_policy.py -q` — green
- [x] Ответ содержит `ocr_api_calls`, `ocr_cost_rub`, `ocr_verify_skipped_reason`

### Phase 4: Интеграция и конфиг

- [x] **OCR-007:** Shim, `.env.example`, metadata в draft (✅ Done)
- [x] **OCR-008:** Скрипт пилота `scripts/ocr_pilot_compare.py` (✅ Done)

**Checkpoint: Integration**
- [x] `pytest tests/test_commercial_web_flow.py -q` — green (56 passed, 1 skipped)
- [ ] Ручной прогон на 1 тестовом фото с реальным GigaChat (вне CI)

### Phase 5: Пилот и калибровка

- [ ] **OCR-009:** Пилот 30–50 реальных фото, отчёт S1–S3 (⏳ Pending)

**Checkpoint: Complete**
- [ ] Критерии S1–S4 из спеки подтверждены
- [ ] Пороги `auto` зафиксированы в спеке / `.env.example`

---

## Task Details

### OCR-001: Package + OpenAI provider

**Description:** Создать `core/ocr/` с `providers/base.py` (Protocol) и `providers/openai.py` — вынести `_call_gpt_for_plates`, `verify_plates_with_gpt_vision`, `_load_image_payload` из `ocr_gpt.py`. Публичный API пока через shim.

**Acceptance criteria:**
- [ ] `core/ocr/__init__.py` экспортирует `recognize_text_smart`, `apply_plates_with_ai`
- [ ] OpenAI provider реализует Extract и Verify
- [ ] Поведение с `OCR_PROVIDER=openai` идентично текущему (1 вызов)

**Verification:**
- [ ] `pytest tests/test_recognition_pipeline.py -q`

**Dependencies:** None  
**Files:** `core/ocr/**`, начало shim в `core/ocr_gpt.py`  
**Scope:** M (4–6 файлов)

---

### OCR-002: Parsing + prompts

**Description:** Перенести `parse_gpt_response`, `parse_verify_response`, `get_verification_prompt`, `_validate_plate_item` в `core/ocr/parsing.py` и `prompts.py` (источник — `core/plate_format_prompt.py`).

**Acceptance criteria:**
- [ ] Все существующие тесты parsing проходят без изменения контракта
- [ ] `core/ocr_gpt.py` реэкспортирует parsing-функции для тестов

**Verification:**
- [ ] `pytest tests/test_recognition_pipeline.py -q -k "parse or validate"`

**Dependencies:** OCR-001  
**Files:** `core/ocr/parsing.py`, `core/ocr/prompts.py`  
**Scope:** S

---

### OCR-003: Parser Gate

**Description:** `apply_parser_gate(plates)` — для каждой строки вызывает `parse_line(f"{normalized_candidate} {qty}")`; при `parsed == false` → `issues += ["parser_rejected"]`, `confidence = min(confidence, 0.5)`.

**Acceptance criteria:**
- [ ] Все строки parsed → без изменений issues/confidence
- [ ] Невалидная марка → `parser_rejected` + пониженный confidence
- [ ] Пустой список → без ошибок

**Verification:**
- [ ] `pytest tests/test_ocr_parser_gate.py -q`

**Dependencies:** None (параллельно с OCR-001)  
**Files:** `core/ocr/parser_gate.py`, `tests/test_ocr_parser_gate.py`  
**Scope:** S

---

### OCR-004: Verify Policy

**Description:** `should_run_verify(mode, max_api_calls, image_size_bytes, plates, settings) -> (bool, str)` — режимы `never` / `always` / `auto` + эвристики из спеки.

**Acceptance criteria:**
- [ ] `never` / `max_api_calls=1` → всегда skip
- [ ] `always` → всегда run (если max ≥ 2)
- [ ] `auto` — все 5 условий skip; любое нарушение → run
- [ ] `reason` заполняется для metadata

**Verification:**
- [ ] `pytest tests/test_ocr_verify_policy.py -q`

**Dependencies:** None (параллельно с OCR-003)  
**Files:** `core/ocr/verify_policy.py`, `tests/test_ocr_verify_policy.py`  
**Scope:** S

---

### OCR-005: GigaChat provider + settings

**Description:** `providers/gigachat.py` — Vision через SDK `gigachat`, scope `GIGACHAT_API_PERS`, `asyncio.to_thread`. Добавить поля в `core/config/settings.py` и `requirements.txt`, `.env.example`.

**Acceptance criteria:**
- [ ] Mock SDK в тестах — Extract и Verify возвращают JSON
- [ ] `cost_rub` считается по 0,65 ₽/1000 токенов
- [ ] Credentials не в коде; ошибка при отсутствии ключа

**Verification:**
- [ ] `pytest tests/test_ocr_gigachat_provider.py -q`

**Dependencies:** OCR-002  
**Files:** `core/ocr/providers/gigachat.py`, `core/config/settings.py`, `requirements.txt`, `.env.example`  
**Scope:** M

**Settings (новые поля):**
```env
OCR_PROVIDER=gigachat
GIGACHAT_CREDENTIALS=
GIGACHAT_MODEL=GigaChat-2-Max
GIGACHAT_SCOPE=GIGACHAT_API_PERS
OCR_VERIFY_MODE=auto
OCR_MAX_API_CALLS=2
OCR_VERIFY_AUTO_MAX_ROWS=15
OCR_VERIFY_AUTO_MIN_CONFIDENCE=0.92
OCR_VERIFY_AUTO_MAX_BYTES=819200
```

**Миграция:** `OCR_VERIFY_ENABLED` → deprecated; при наличии маппить `false` → `never`, `true` → `always` (лог warning).

---

### OCR-006: Pipeline wiring

**Description:** `pipeline.py` — оркестрация: Extract → Parser Gate → `should_run_verify` → [Verify → Parser Gate] → `_build_result_payload` с новыми полями metadata.

**Acceptance criteria:**
- [ ] 1 вызов при `auto` + все checks pass
- [ ] 2 вызова при `always` или сомнительном Extract
- [ ] `ocr_api_calls`, `ocr_cost_rub`, `ocr_method`, `ocr_verify_skipped_reason` в ответе
- [ ] Логи `[OCR]` с `api_calls`, `cost_rub`, `verify_decision`

**Verification:**
- [ ] `pytest tests/test_recognition_pipeline.py -q`
- [ ] `pytest tests/test_commercial_ocr_policy.py -q`

**Dependencies:** OCR-001, OCR-003, OCR-004, OCR-005  
**Files:** `core/ocr/pipeline.py`, `core/ocr/__init__.py`  
**Scope:** M

---

### OCR-007: Shim + env + draft metadata

**Description:** Завершить shim `core/ocr_gpt.py`; прокинуть metadata в `commercial_draft_service` (если нужно для wizard); обновить `.env.example`.

**Acceptance criteria:**
- [ ] Все импорты `core.ocr_gpt` работают без изменений в app/
- [ ] Draft preview содержит OCR metadata для UI warnings (Q4 из спеки)
- [ ] Staging env: `OCR_VERIFY_MODE=always`

**Verification:**
- [ ] `pytest tests/test_commercial_web_flow.py -q`

**Dependencies:** OCR-006  
**Files:** `core/ocr_gpt.py`, `app/services/commercial_draft_service.py`, `.env.example`  
**Scope:** S

---

### OCR-008: Пилотный скрипт

**Description:** `scripts/ocr_pilot_compare.py --image PATH [--provider gigachat|openai] [--verify-mode auto|always|never]` — выводит plates, api_calls, cost, corrections.

**Acceptance criteria:**
- [ ] Работает на одном фото без изменения pytest
- [ ] Сравнение режимов `auto` vs `always` на одном файле

**Verification:**
- [ ] Ручной запуск на тестовом фото

**Dependencies:** OCR-006  
**Files:** `scripts/ocr_pilot_compare.py`  
**Scope:** S

---

### OCR-009: Пилот на реальных фото

**Description:** 30–50 снимков из архива менеджеров; сравнение с GPT-4o; калибровка порогов `auto`.

**Acceptance criteria:**
- [ ] S1: ≥ 90% строк без ручной правки марки/qty
- [ ] S2: ≥ 50% фото с 1 вызовом в `auto`
- [ ] S3: средняя стоимость в `auto` на 30%+ ниже vs `always`
- [ ] Отчёт в `ai_docs/develop/reports/2026-07-08-gigachat-ocr-pilot.md`

**Verification:**
- [ ] Подтверждение менеджером + structured logs

**Dependencies:** OCR-007, OCR-008  
**Files:** отчёт  
**Scope:** L (ручная работа)

---

## Dependency Graph

```
OCR-001 ──┬── OCR-002 ── OCR-005 ──┐
          │                         ├── OCR-006 ── OCR-007 ── OCR-009
OCR-003 ──┤                         │              │
OCR-004 ──┘                         └── OCR-008 ─────┘
```

**Параллелизация:** OCR-003 и OCR-004 можно делать параллельно с OCR-001/002.

## Commands

```powershell
Set-Location "c:\Users\Роман\Desktop\Шишов"
.\.venv\Scripts\Activate.ps1

# После каждого checkpoint
pytest tests/test_recognition_pipeline.py tests/test_ocr_verify_policy.py tests/test_commercial_ocr_policy.py -q

# Полная регрессия перед пилотом
pytest tests/test_commercial_web_flow.py -q

# Пилот (после OCR-008)
python scripts/ocr_pilot_compare.py --image path/to/photo.jpg --provider gigachat --verify-mode auto
```

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| GigaChat хуже GPT на реальных КП | High | Пилот OCR-009 до prod; fallback `OCR_PROVIDER=openai` |
| Freemium не покрывает коммерческое использование | Med | Уточнить ToS; пакет 1950 ₽ |
| Latency > 15 с | Med | `asyncio.to_thread`; таймаут + понятная ошибка в UI |
| Ломаются моки в тестах при рефакторинге | Med | Shim + поэтапная миграция patch-путей |
| Эвристики `auto` пропускают ошибки | High | Консервативные дефолты; staging = `always` |

## Open Questions

| # | Вопрос | Default |
|---|--------|---------|
| Q1 | Freemium физлица для завода? | Пилот на freemium, уточнить ToS |
| Q2 | Warning в UI при `OCR_MAX_API_CALLS=1` + low confidence? | Да |
| Q3 | Логировать статистику 1 vs 2 вызовов? | Да, structured log |
| Q4 | Удалять `OCR_VERIFY_ENABLED` сразу? | Deprecated + маппинг на `OCR_VERIFY_MODE` |

## Success Criteria (из спеки)

| # | Критерий | Задача |
|---|----------|--------|
| S1 | ≥ 90% строк без правки на пилоте | OCR-009 |
| S2 | ≥ 50% фото с 1 вызовом в `auto` | OCR-009 |
| S3 | −30% cost vs `always` | OCR-009 |
| S4 | pytest green | OCR-006, OCR-007 |
| S5 | `commercial_draft_service` не сломан | OCR-007 |
