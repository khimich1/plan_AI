# Plan: КП — правка архива через конструктор

**Created:** 2026-09-02 12:57  
**Orchestration:** orch-2026-09-02-09-57-kp-archive-edit  
**Спека:** [../../specs/kp-archive-edit-in-constructor.md](../../specs/kp-archive-edit-in-constructor.md)  
**Идея:** [../../ideas/kp-archive-edit-in-constructor.md](../../ideas/kp-archive-edit-in-constructor.md)  
**Goal:** Из drawer КП «в архиве» открывать конструктор (допись / правка), сохранять в тот же `kp_id` со статусом «в архиве»; Итоги read-only; nav «Конструктор КП».  
**Total Tasks:** 8  
**Priority:** High  
**Status:** PLAN ✅ · IMPLEMENT ✅

## Overview

Переворот статус-гейта R2: resume/update только для «в архиве» (три слоя — hydrate, `save_offer` update path, `KpPersistenceService.update_kp_from_order_data`). FE: две CTA с разным landing, убрать resume для «в работе», read-only «Итоги», rename nav. Без новых endpoint’ов и зависимостей. TDD на каждый срез.

## Architecture Decisions

- **Три гейта, один контракт.** Spec перечисляет lifecycle; в коде есть ещё гейт в `core/kp_persistence_service.py` (`update_kp_from_order_data` → «только в работе»). Без его флипа archive-save после resume упадёт на persistence. Все три: allow только «в архиве».
- **Статус после update.** `update_kp_from_order_data` **не** меняет `kp_meta.status`. На update-path в `save_offer` писать `saved_offer.status = existing_status` («в архиве»), не доверять дефолту `status="в работе"`.
- **Один FE handler.** `openInConstructor(landing: "append" | "result")` — общий resume; append дополнительно `start-append-cycle`.
- **Итоги.** Для `status === "в архиве"` — только `FinanceCard`/текст (вес, НДС, рейс, доставка, скидка, итого). Инпуты/OK не рендерить; mutations не звать. Для прочих статусов текущие finance-поля можно оставить (API hardening out of scope).
- **Сообщения ошибок (RU):** в духе существующих → «…только в статусе «в архиве»» / «Обновление КП разрешено только в статусе «в архиве»».

## Tasks Overview

1. **ARC-001** Persistence gate flip `(feat-be)` — dependsOn: []
2. **ARC-002** Hydrate gate flip `(feat-be)` — dependsOn: []
3. **ARC-003** save_offer update path + preserve «в архиве» `(feat-be)` — dependsOn: [ARC-001]
4. **ARC-004** HTTP + integration resume/save tests `(api)` — dependsOn: [ARC-002, ARC-003]
5. **ARC-005** Drawer dual CTA + remove «в работе» resume `(feat-fe)` — dependsOn: [ARC-004]
6. **ARC-006** Read-only «Итоги» for «в архиве» `(ui)` — dependsOn: [ARC-005]
7. **ARC-007** Nav rename «Конструктор КП» `(ui)` — dependsOn: []
8. **ARC-008** Focused verify + docs status `(chore)` — dependsOn: [ARC-004, ARC-006, ARC-007]

## Dependencies Graph

```
ARC-001 ──────────────┐
                      ├──► ARC-003 ──┐
ARC-002 ──────────────┘              ├──► ARC-004 ──► ARC-005 ──► ARC-006 ──┐
                                     │                                       ├──► ARC-008
ARC-007 (parallelSafe) ──────────────┴───────────────────────────────────────┘
```

`ARC-001` ∥ `ARC-002` ∥ `ARC-007` (разные файлы).  
`ARC-005`/`ARC-006` после backend, чтобы manual smoke совпадал с API.

---

## Task List

### Phase 1 — Backend status gate (TDD)

#### Task ARC-001: Persistence `update_kp_from_order_data` → только «в архиве»

**Type:** `feat-be`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** []  
**parallelSafe:** true (с ARC-002, ARC-007)  
**needsExplore:** true  
**securitySensitive:** false  

**Description:** RED→GREEN: гейт в `KpPersistenceService.update_kp_from_order_data` разрешает sync только при `kp_meta.status == «в архиве»`; «в работе» и прочие — ValueError. Обновить фикстуры/тесты в `test_kp_persistence_mixed.py` (сейчас seed default «в работе»). Убедиться, что после update статус в `kp_meta` **не** меняется.

**Acceptance criteria:**
- [x] Allow: status «в архиве» → update succeeds, same `kp_id`, status остаётся «в архиве»
- [x] Reject: «в работе», «На СГП», «выполнено», … → ValueError с текстом про «в архиве»
- [x] Существующие sync-тесты (append line_id / preserve id) переведены на seed «в архиве»

**Verification:**
```bash
pytest tests/test_kp_persistence_mixed.py -q -k "update_kp_from_order_data"
```

**Files likely touched:**
- `core/kp_persistence_service.py`
- `tests/test_kp_persistence_mixed.py`

---

#### Task ARC-002: Hydrate gate → только «в архиве»

**Type:** `feat-be`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** []  
**parallelSafe:** true  
**needsExplore:** true  

**Description:** RED→GREEN: `CommercialDraftLifecycle.hydrate_draft_from_saved_kp` (и docstring `ArchiveService.resume_as_draft`) принимают только «в архиве». Переписать MNA-601 unit-тесты: happy path на «в архиве»; reject «в работе» и прочих. Сообщение: «Дополнить КП можно только в статусе «в архиве».»

**Acceptance criteria:**
- [x] Hydrate «в архиве» → 200-shaped draft, `resume_kp_id`, `saved_offer.status == «в архиве»`, `current_step=result`
- [x] Hydrate «в работе» / СГП / … → ValueError
- [x] Docstrings больше не говорят «в работе only»

**Verification:**
```bash
pytest tests/test_commercial_draft_append.py -q -k hydrate
```

**Files likely touched:**
- `app/services/commercial_draft_lifecycle.py` (hydrate gate + docstring)
- `app/services/archive_service.py` (docstring)
- `tests/test_commercial_draft_append.py` (hydrate block ~MNA-601)

---

#### Task ARC-003: `save_offer` update path — allow «в архиве», не флипать статус

**Type:** `feat-be`  
**Priority:** Critical  
**Complexity:** Moderate  
**dependsOn:** [ARC-001]  
**parallelSafe:** false  
**needsExplore:** true  

**Description:** RED→GREEN: в `save_offer` при `existing_kp_id` / `resume_kp_id` гейт `existing_status == «в архиве»`. После `update_offer_from_order_data` в `saved_offer` писать status **«в архиве»** (из existing / явно), даже если параметр `status` дефолтный «в работе». Новый/обновлённый тест: resume-draft + `save_draft(mode="archive")` → тот же `kp_id`, status «в архиве». Старый reject-тест: «в работе» / «выполнено» отклоняются.

**Acceptance criteria:**
- [x] Update при saved_offer.status «в архиве» → ok, тот же kp_id
- [x] Update при «в работе» → ValueError
- [x] `saved_offer.status` после archive-save = «в архиве» (не «в работе»)
- [x] Persistence не меняет `kp_meta.status` (регресс через ARC-001)

**Verification:**
```bash
pytest tests/test_commercial_draft_append.py -q -k "save_offer or saved_kp or resume"
```

**Files likely touched:**
- `app/services/commercial_draft_lifecycle.py` (`save_offer` update branch)
- `tests/test_commercial_draft_append.py` (MNA-304 save/update tests)

---

#### Task ARC-004: HTTP resume + multi-append integration under new gate

**Type:** `api`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [ARC-002, ARC-003]  
**parallelSafe:** false  
**needsExplore:** true  

**Description:** Обновить `tests/test_archive_endpoints.py` (MNA-601 HTTP: mock/message «в архиве»; reject non-archive). Переписать `tests/test_commercial_multi_append_flow.py`: save в архив → resume → append → archive-save → тот же `kp_id`, status «в архиве»; resume blocked для «в работе». При необходимости точечно `test_commercial_web_flow.py -k "resume or archive or save"`.

**Acceptance criteria:**
- [x] `POST .../archive/{id}/resume` ok для «в архиве»; 4xx для «в работе»
- [x] SC-6 flow: archive → resume → append → archive save → same kp_id + «в архиве»
- [x] Старые ожидания «в работе» для resume/update убраны

**Verification:**
```bash
pytest tests/test_archive_endpoints.py -q -k resume
pytest tests/test_commercial_multi_append_flow.py -q
pytest tests/test_commercial_web_flow.py -q -k "resume or archive or save"
```

**Files likely touched:**
- `tests/test_archive_endpoints.py`
- `tests/test_commercial_multi_append_flow.py`
- `tests/test_commercial_web_flow.py` (только если краснеет)

---

### Checkpoint: Backend

- [x] Persistence + hydrate + save_offer gates согласованы на «в архиве»
- [x] Integration resume/save зелёный
- [x] Human: ok to proceed FE

---

### Phase 2 — Frontend drawer

#### Task ARC-005: Dual CTA landing + remove «в работе» resume

**Type:** `feat-fe`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [ARC-004]  
**parallelSafe:** false  
**needsExplore:** true  

**Description:** TDD в `OfferDetailsDrawer.test.tsx`: для «в архиве» кнопки «(+ Добавить)» и «Редактировать»; для «в работе» **нет** «Добавить другое наименование» и нет новых CTA. Implement: один `openInConstructor(landing)`; append → `hydrate-draft` + `start-append-cycle`; result → только hydrate (metadata уже `current_step=result`). Удалить старый CTA block `status === "в работе"`.

**Acceptance criteria:**
- [x] S1–S3: две CTA; append → picker path; edit → result path (navigate `/new?draft=…`)
- [x] S6: «в работе» без resume/edit состава
- [x] Pending/error поведение сохранено на обоих CTA
- [x] «В производство» / PDF / XLSX не трогаем

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer
```

**Files likely touched:**
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`

---

#### Task ARC-006: Read-only «Итоги» for «в архиве»

**Type:** `ui`  
**Priority:** High  
**Complexity:** Moderate  
**dependsOn:** [ARC-005]  
**parallelSafe:** false  
**needsExplore:** true  

**Description:** RED→GREEN: при `status === "в архиве"` секция «Итоги» без inputs/OK для целевой суммы, скидки, стоимости рейса (и pile logistics / pending override inputs). Показать сводку текстом/`FinanceCard` (вес, НДС, рейс/доставка, скидка %, итого). Не вызывать discount/logistics mutations из drawer.

**Acceptance criteria:**
- [x] S5: нет finance OK/inputs в «в архиве»
- [x] Видны вес / НДС / доставка / итого (и скидка как текст, если была)
- [x] RTL покрывает отсутствие spin/textbox+OK для скидки/рейса/цели

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer
```

**Files likely touched:**
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.tsx`
- `frontend/src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx`

---

### Phase 3 — Nav (parallel-safe early)

#### Task ARC-007: Nav label «Конструктор КП»

**Type:** `ui`  
**Priority:** Medium  
**Complexity:** Simple  
**dependsOn:** []  
**parallelSafe:** true  
**needsExplore:** false  

**Description:** В `AppHeader` заменить «Создать КП» → «Конструктор КП». Добавить/обновить тест под `src/app/layout` (сейчас отдельного test-файла может не быть — создать `AppHeader.test.tsx`). Комментарии с «Создать КП» в bridge — по желанию, не блокер UI.

**Acceptance criteria:**
- [x] S7: видимый nav text = «Конструктор КП»
- [x] Vitest находит label

**Verification:**
```bash
cd frontend && npm run test -- src/app/layout
```

**Files likely touched:**
- `frontend/src/app/layout/AppHeader.tsx`
- `frontend/src/app/layout/AppHeader.test.tsx` (new)

---

### Phase 4 — Verify

#### Task ARC-008: Focused suites + typecheck + docs status

**Type:** `chore`  
**Priority:** High  
**Complexity:** Simple  
**dependsOn:** [ARC-004, ARC-006, ARC-007]  
**parallelSafe:** false  

**Description:** Прогнать команды из спеки; починить оставшиеся красные ожидания «в работе» для resume. Обновить idea/spec/plan → IMPLEMENT ✅ после зелёного прогона (documenter/orchestrator). Не коммитить. Не убивать `./run+logs.sh`.

**Acceptance criteria:**
- [x] S8–S9: focused pytest + vitest + typecheck green
- [x] Docs статусы согласованы

**Verification:**
```bash
cd frontend && npm run test -- src/features/commercial-archive/components/OfferDetailsDrawer
cd frontend && npm run test -- src/app/layout
cd frontend && npm run typecheck
pytest tests/test_archive_endpoints.py -q -k resume
pytest tests/test_commercial_web_flow.py -q -k "resume or archive or save"
pytest tests/test_kp_persistence_mixed.py -q -k "update_kp_from_order_data"
pytest tests/test_commercial_draft_append.py -q -k "hydrate or save_offer or saved_kp"
pytest tests/test_commercial_multi_append_flow.py -q
```

**Files likely touched:**
- `ai_docs/specs/kp-archive-edit-in-constructor.md`
- `ai_docs/ideas/kp-archive-edit-in-constructor.md`
- `ai_docs/develop/plans/2026-09-02-kp-archive-edit-in-constructor.md`

---

## Progress (orchestrator)

- ✅ ARC-001: Persistence gate `(feat-be)` (Done)
- ✅ ARC-002: Hydrate gate `(feat-be)` (Done)
- ✅ ARC-003: save_offer update `(feat-be)` (Done)
- ✅ ARC-004: HTTP + integration `(api)` (Done)
- ✅ ARC-005: Dual CTA + remove в работе `(feat-fe)` (Done)
- ✅ ARC-006: Read-only Итоги `(ui)` (Done)
- ✅ ARC-007: Nav rename `(ui)` (Done)
- ✅ ARC-008: Verify + docs `(chore)` (Done)

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Третий гейт в `update_kp_from_order_data` забыт → save падает после resume | High | ARC-001 первым; integration ARC-004 ловит E2E |
| `saved_offer.status` флипается в «в работе» из-за default `save_offer(..., status="в работе")` | High | ARC-003: на update path писать existing «в архиве» |
| Каскад тестов MNA/SC-6 на «в работе» | Med | Явно переписать в ARC-002/003/004; не «чинить молча» |
| Finance UI для не-архивных статусов сломать при правке Итоги | Med | Условный рендер только `status === "в архиве"` |
| Ручной smoke при живом `./run+logs.sh` | Low | Не убивать процесс; FE HMR / pytest отдельно |

## Out of scope (remind)

- Hardening PATCH discount/logistics API
- Inline rename/delete в таблице drawer
- Edit после производства
- MoveToProduction / readiness

## Open Questions

_Нет блокирующих._ Spec assumptions locked. Единственное уточнение для human checkpoint: подтвердить, что persistence-гейт тоже только «в архиве» (не «в архиве **или** в работе») — иначе производство снова можно обновить через `update_kp_from_order_data`. Plan assumes **только «в архиве»** per D-status / S8.

## Implementation Notes for workers

- Inject `plan-web-context` skill.
- TDD: failing test first on every behavior change.
- Prefer ≤5 files per task; do not expand scope.
- No commit unless user asks; no new deps.
- After plan approval checkpoint → execute tasks in DAG order.
