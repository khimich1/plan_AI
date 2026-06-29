# PRD: Онбординг новых пользователей системы «Шишов»

> **Тип:** Product Requirements Document (feature-spec)  
> **Фаза SDD:** SPECIFY  
> **Дата:** 2026-06-19  
> **Статус:** черновик на ревью  
> **Родительский анализ:** [`product-analysis-swot-ost-assumptions.md`](./product-analysis-swot-ost-assumptions.md)  
> **Baseline:** [`project-baseline.md`](./project-baseline.md)

---

## ASSUMPTIONS I'M MAKING

1. **Целевая аудитория MVP** — новые менеджеры (`manager`) и admin; production получает только role-based redirect, без wizard onboarding.
2. **MVP — преимущественно frontend** (welcome, hints, demo prefill); персистентность «onboarding completed» — localStorage + опционально поле в `users` (v2).
3. **Desired outcome:** первое корректное КП за ≤15 мин без помощи коллеги.
4. **Не менееём** текущий wizard contract (`wizard_state.can_proceed_to`); упрощение CTA на шаге 1 — P1, не MVP.
5. **OCR/AI** остаются опциональными; onboarding не требует `OPENAI_API_KEY`.
6. **Аналитика** — минимальные frontend events (console/custom hook); полноценный analytics — v2.

→ Поправь допущения сейчас или подтверди.

---

## 1. Objective & Problem Statement

### Objective

Сократить время и когнитивную нагрузку при первом создании КП, чтобы новый менеджер мог самостоятельно пройти wizard от входа до сохранения корректного коммерческого предложения.

### Problem Statement

**Сейчас:** после авторизации пользователь без различия роли перенаправляется на `/new` (`LoginPage.tsx`, константа `DEFAULT_REDIRECT = "/new"`). Wizard (`CommercialOfferWizard.tsx`) открывается сразу на шаге `plates` без контекста. Шаг 1 (`PlateInputStep.tsx`) требует знания формата марок ПБ, понимания двухфазного flow («Распознать» → «Обработать») и доменных терминов (нормализация, широкие плиты, OCR corrections).

**Последствия:**
- Высокий drop-off на шаге 1 (оценка: 40–60% первых сессий без save — требует baseline замера).
- Запросы к коллегам «как пользоваться» → масштабирование обучения не работает.
- Production-пользователи попадают на wizard вместо `/production`.
- Низкое доверие к результату из-за непрозрачного preview и расхождений pricing (см. strategic roadmap Phase 1).

**Success definition:** median time-to-first-KP ≤ 15 мин; first-session completion rate ≥ 60% на pilot-группе (≥5 новых менеджеров, 2 недели).

---

## 2. Users & Personas

### Primary: Менеджер по продажам (`manager`)

| Атрибут | Описание |
|---------|----------|
| **Цель** | Быстро выдать клиенту КП с правильными ценами и PDF |
| **Контекст** | Знает марки ПБ из Excel/переписки; не знает UI системы |
| **Боли** | Непонятный формат ввода; страх ошибиться; «Распознать» vs «Обработать» |
| **Критерий успеха** | Сохранил КП в архив, может найти его и отправить клиенту |

### Secondary: Производство (`production`)

| Атрибут | Описание |
|---------|----------|
| **Цель** | Планирование смен, календарь, документы дня |
| **Контекст** | Не создаёт КП; заходит редко, по задаче |
| **Боли** | После login попадает на wizard `/new` — лишние клики |
| **Критерий успеха** | Login → сразу `/production` |

### Tertiary: Администратор (`admin`)

| Атрибут | Описание |
|---------|----------|
| **Цель** | Всё выше + управление БД, восстановление плит |
| **Контекст** | Технически подкован; может обучать других |
| **Боли** | Нет «режима демо» для показа новичкам |
| **Критерий успеха** | Может запустить demo flow за 1 мин для обучения |

---

## 3. Success Metrics

| Метрика | Baseline (оценка) | Target MVP | Target v2 | Измерение |
|---------|-------------------|------------|-----------|-----------|
| **Time-to-first-KP** | 30–45 мин (с помощью) | Median ≤ 20 мин | Median ≤ 15 мин | `login_at` → `first_save_at` |
| **First-session completion** | ~30% | ≥ 50% | ≥ 60% | wizard_start → save |
| **Step 1 drop-off** | ~50% | ≤ 35% | ≤ 25% | plates mount → handleProcess success |
| **Welcome → wizard start** | N/A | ≥ 80% click-through | ≥ 90% | welcome CTA click |
| **Demo draft usage (first session)** | 0% | ≥ 40% | ≥ 30% | demo button click |
| **D7 repeat (new managers)** | Unknown | ≥ 40% | ≥ 50% | second KP within 7d |
| **Support tickets «как создать КП»** | Baseline TBD | −50% | −70% | helpdesk |
| **Production wrong landing** | ~100% on `/new` | 0% | 0% | role=production on `/new` |

---

## 4. Scope

### In Scope (MVP)

| # | Feature | Priority |
|---|---------|----------|
| 1 | Welcome screen после первого login (или пока не dismissed) | P0 |
| 2 | Role-based landing redirect | P0 |
| 3 | Demo draft — prefill типового заказа в textarea шага 1 | P0 |
| 4 | Inline format hints на `PlateInputStep` | P0 |
| 5 | Collapsible «Справка по формату марок ПБ» | P0 |
| 6 | Preview table improvements — акцент на unparsed_lines | P0 |
| 7 | Minimal onboarding events (hook для будущей analytics) | P1 |
| 8 | «Пропустить обучение» / dismiss welcome | P0 |

### In Scope (v2)

| # | Feature | Priority |
|---|---------|----------|
| 9 | Unified primary CTA «Продолжить» на шаге 1 | P1 |
| 10 | Coach marks / guided tour (5 шагов wizard) | P2 |
| 11 | Success celebration после первого save | P2 |
| 12 | Server-side `onboarding_completed_at` в `users` | P2 |
| 13 | Admin «режим демо» с watermarked preview | P2 |
| 14 | Price breakdown hints (после Phase 1 pricing) | P1 |

### Out of Scope

- Переделка всего wizard UI / design system
- Онboarding Telegram-бота (отдельная spec)
- Очное обучающее видео / LMS
- Изменение backend wizard contract (MVP)
- Phase 1 pricing refactor (параллельный трек)
- Мультиязычность

---

## 5. User Stories & Acceptance Criteria

### Epic 1: Welcome & Orientation

#### US-1.1: Welcome screen для менеджера

**Как** новый менеджер,  
**я хочу** увидеть краткое объяснение после входа,  
**чтобы** понять, что система делает и с чего начать.

**Acceptance Criteria:**
- [ ] После login `manager`/`admin` без флага `onboarding_dismissed` видит `/welcome` (или modal на `/new`)
- [ ] Экран содержит: заголовок, 3 bullet «что будет» (ввод плит → условия → PDF), две CTA: «Создать КП» и «Попробовать на примере»
- [ ] «Пропустить» сохраняет `onboarding_dismissed` в localStorage и больше не показывает welcome
- [ ] «Создать КП» → `/new` (wizard step `plates`)
- [ ] «Попробовать на примере» → `/new?demo=1` с prefill demo order

#### US-1.2: Role-based landing

**Как** пользователь производства,  
**я хочу** попадать сразу на страницу производства,  
**чтобы** не видеть wizard КП.

**Acceptance Criteria:**
- [ ] `production` после login redirect → `/production`
- [ ] `manager` / `admin` → `/welcome` (если не dismissed) или `/new`
- [ ] Authenticated user на `/login` redirect по роли (не всегда `/new`)
- [ ] Deep link `/new` для `production` → redirect `/production` с info toast (optional)

**Код-референс:** `frontend/src/pages/login/LoginPage.tsx` — заменить единый `DEFAULT_REDIRECT` на функцию `getDefaultRoute(role)`.

---

### Epic 2: Demo Draft

#### US-2.1: Типовой заказ одной кнопкой

**Как** новый менеджер,  
**я хочу** загрузить пример типового заказа,  
**чтобы** увидеть, как выглядит правильный ввод.

**Acceptance Criteria:**
- [ ] Константа `DEMO_PLATE_ORDER_TEXT` в frontend (5–7 позиций, mix форматов: `ПБ 78-12-8п 2`, `71-12-8 3`)
- [ ] Кнопка «Заполнить примером» на шаге 1 и на welcome
- [ ] При demo prefill показывается info banner «Это демо-заказ — можно изменить или распознать»
- [ ] Demo не auto-submit; пользователь сам нажимает «Распознать»
- [ ] Query param `?demo=1` триггерит prefill при mount wizard

---

### Epic 3: Inline Hints (Step 1)

#### US-3.1: Подсказка формата на шаге plates

**Как** новый менеджер,  
**я хочу** видеть примеры формата строк рядом с полем ввода,  
**чтобы** не гадать, как писать марки.

**Acceptance Criteria:**
- [ ] Под textarea — блок «Примеры формата» с 3–4 строками и пояснениями
- [ ] Collapsible «Подробная справка» с форматами: `ПБ L-W-loadп qty`, bare `L-W-load`, L×W×H мм (warning про 8п default)
- [ ] Placeholder textarea обновлён (уже есть в коде — сохранить)
- [ ] Hint не перекрывает OCR upload на mobile (stack vertically)

**Детали UX:** см. [`ux-wizard-step-plates.md`](./ux-wizard-step-plates.md)

#### US-3.2: Понятные ошибки на шаге 1

**Как** новый менеджер,  
**я хочу** получать actionable ошибки,  
**чтобы** исправить ввод без звонка коллеге.

**Acceptance Criteria:**
- [ ] Пустой submit → «Введите текст или загрузите изображение» (уже есть)
- [ ] После recognize: если `unparsed_lines.length > 0` — alert с списком и hint «Проверьте формат — см. справку»
- [ ] Если `order_data.length === 0` — блокирующее сообщение, CTA disabled на «Обработать»
- [ ] OCR verify failed — сохранить текущий warning + добавить link «Открыть справку по формату»

---

### Epic 4: Preview & Progress

#### US-4.1: Preview table с акцентом на проблемы

**Как** новый менеджер,  
**я хочу** сразу видеть, какие строки не распознались,  
**чтобы** исправить до перехода дальше.

**Acceptance Criteria:**
- [ ] `KpPlatePreviewPanel` — unparsed_lines визually prominent (красный/amber блок)
- [ ] Summary card «Предпросмотр обработанного списка» — link «Что означают предупреждения?»
- [ ] Wide plates warning сохраняется с текстом «проверка на шаге 2»

#### US-4.2: Progress visibility

**Как** новый менеджер,  
**я хочу** видеть, на каком шаге я и сколько осталось,  
**чтобы** не чувствовать себя потерянным.

**Acceptance Criteria:**
- [ ] `WizardProgress` sidebar — добавить subtitle «Шаг X из 5» в main content header (MVP)
- [ ] First visit tooltip на sidebar «Нажмите на шаг, чтобы вернуться» (optional P1)

---

### Epic 5: Analytics & Completion (v2)

#### US-5.1: Success после первого save

**Как** новый менеджер,  
**я хочу** получить подтверждение успеха,  
**чтобы** знать, что всё прошло правильно.

**Acceptance Criteria (v2):**
- [ ] Modal «Первое КП готово!» с kp_id, ссылкой на архив, CTA «Создать ещё»
- [ ] Показывается один раз (localStorage `first_kp_saved`)

---

## 6. Functional Requirements

### FR-1: Welcome Screen

| ID | Requirement |
|----|-------------|
| FR-1.1 | Route `/welcome` protected, roles: `admin`, `manager` |
| FR-1.2 | Content: hero, 3-step explainer, CTA primary/secondary, skip link |
| FR-1.3 | State: `localStorage.shishov_onboarding_dismissed = "1"` |
| FR-1.4 | Redirect guard: если dismissed → `/new` |

### FR-2: Role-based Redirect

| ID | Requirement |
|----|-------------|
| FR-2.1 | `getDefaultRoute(user.role)`: `production` → `/production`, `manager`/`admin` → `/welcome` or `/new` |
| FR-2.2 | Update `LoginPage.tsx` Navigate and post-login navigate |
| FR-2.3 | Update `ProtectedRoute` root redirect if applicable |

### FR-3: Demo Draft

| ID | Requirement |
|----|-------------|
| FR-3.1 | `DEMO_PLATE_ORDER_TEXT` constant in `frontend/src/features/commercial-offer/lib/demoOrder.ts` |
| FR-3.2 | `CommercialOfferWizard` reads `?demo=1` search param on mount |
| FR-3.3 | Prefill dispatches `set-source` with demo text |

### FR-4: PlateInputStep Hints

| ID | Requirement |
|----|-------------|
| FR-4.1 | New component `PlateFormatHint.tsx` colocated in `components/steps/` |
| FR-4.2 | Collapsible section with format examples from `plate_line_parser` conventions |
| FR-4.3 | Demo fill button triggers `onTextChange(DEMO_PLATE_ORDER_TEXT)` |

### FR-5: Error & Preview Enhancements

| ID | Requirement |
|----|-------------|
| FR-5.1 | Enhanced unparsed_lines display in `KpPlatePreviewPanel` |
| FR-5.2 | Step header «Шаг 1 из 5» in `StepLayout` or wizard shell |

### FR-6: Events (P1)

| ID | Requirement |
|----|-------------|
| FR-6.1 | `useOnboardingEvents()` hook: `track(event, payload)` → console in dev, stub for prod |
| FR-6.2 | Events: `welcome_view`, `welcome_cta_create`, `welcome_cta_demo`, `welcome_skip`, `demo_prefill`, `plates_recognize`, `plates_process`, `first_save` |

---

## 7. Non-Functional Requirements

| ID | Category | Requirement |
|----|----------|-------------|
| NFR-1 | Performance | Welcome screen LCP ≤ 1.5s; no additional API calls on welcome |
| NFR-2 | Accessibility | Welcome and hints keyboard-navigable; collapsible aria-expanded |
| NFR-3 | Security | Demo text static client-side only; no bypass auth |
| NFR-4 | Compatibility | Chrome/Edge latest (завод); 1280px+ primary target |
| NFR-5 | i18n | Russian only |
| NFR-6 | Maintainability | Onboarding components in `frontend/src/features/onboarding/` |
| NFR-7 | Testability | Unit tests for `getDefaultRoute`, demo prefill; Vitest |
| NFR-8 | Privacy | localStorage only; no PII in events |

---

## 8. UX Flows

### 8.1. First login — manager (MVP)

```mermaid
flowchart TD
    A[Login /login] --> B{Role?}
    B -->|production| C[/production]
    B -->|manager/admin| D{onboarding_dismissed?}
    D -->|No| E[/welcome]
    D -->|Yes| F[/new wizard]
    E --> G[CTA: Создать КП]
    E --> H[CTA: Попробовать пример]
    E --> I[Skip]
    G --> F
    H --> J[/new?demo=1]
    J --> K[PlateInputStep prefill]
    I --> L[Set localStorage]
    L --> F
    K --> M[User clicks Распознать]
    M --> N[Draft created]
    N --> O[User clicks Обработать]
    O --> P{can_proceed_to}
    P -->|wide-plates| Q[Step 2]
    P -->|manager| R[Step 3]
    Q --> R
    R --> S[Step 4 client]
    S --> T[Step 5 result + save]
```

### 8.2. Returning user

```mermaid
flowchart TD
    A[Login] --> B{onboarding_dismissed?}
    B -->|Yes| C[/new directly]
    B -->|No| D[/welcome]
```

### 8.3. Code touchpoints

| Flow step | File |
|-----------|------|
| Login redirect | `frontend/src/pages/login/LoginPage.tsx` |
| Welcome page | `frontend/src/pages/welcome/WelcomePage.tsx` (new) |
| Wizard entry | `frontend/src/pages/commercial-offer-create/CommercialOfferCreatePage.tsx` |
| Step 1 UI | `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx` |
| Step order | `frontend/src/features/commercial-offer/lib/wizardStepOrder.ts` |
| Wizard orchestration | `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx` |
| Router | `frontend/src/app/router.tsx` |

---

## 9. Phased Rollout

### Phase MVP (2–3 спринта)

**Deliverables:**
- Welcome screen + skip
- Role-based redirect
- Demo prefill + format hints
- Enhanced unparsed_lines UX
- `useOnboardingEvents` stub
- Vitest for redirect logic

**Exit criteria:**
- 3 internal users complete first KP without help
- Production role never lands on `/new` after login
- No regression in `test_commercial_web_flow.py` / frontend build

### Phase v2 (1–2 спринта после MVP metrics)

**Deliverables:**
- Unified CTA on step 1
- First-save celebration modal
- Coach marks (optional library.py library or custom)
- Server-side onboarding flag (migration `users.onboarding_completed_at`)

**Exit criteria:**
- Median time-to-first-KP ≤ 15 min on pilot (n≥5)
- First-session completion ≥ 60%

---

## 10. Dependencies & Risks

### Dependencies

| Dependency | Owner | Status |
|------------|-------|--------|
| Auth role in `/auth/me` response | Backend | ✅ Exists |
| Wizard step contract | Backend `CommercialWorkflowService` | ✅ Stable |
| Demo order text approved by domain expert | Product/завод | ⏳ Need validation |
| Pilot group of new managers | HR/руководство | ⏳ |
| Phase 1 pricing for breakdown hints | Engineering | 🔜 v1.1 |

### Risks

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| Demo order не репрезентативен | Medium | Medium | Согласовать с 1 senior manager |
| Unified CTA ломает server flow | Medium | High | Defer to v2; spike first |
| localStorage cleared → welcome repeats | Low | Low | Acceptable MVP; v2 server flag |
| Production redirect breaks bookmark | Low | Medium | Allow `/new` with redirect + toast |
| Metrics not collected | High | Medium | Manual observation in pilot if no analytics |

---

## 11. Open Questions

1. **Welcome как отдельная страница vs modal на `/new`?** Recommendation: отдельная `/welcome` — чище analytics, меньше clutter wizard.
2. **Показывать welcome admin каждый раз или только «новым»?** MVP: localStorage dismiss; v2: server flag + created_at heuristic.
3. **Кто утверждает текст demo order?** Нужен sample от завода (5–7 строк реального типового заказа).
4. **Нужен ли onboarding для Telegram-бота в v1?** Recommendation: out of scope; ссылка «полная версия в web» в `/help`.
5. **Интегрировать с Phase 1 pricing warnings в preview?** Recommendation: v1.1 после `core/pricing`.
6. **Baseline замер «как сейчас» — кто организует?** Product + 3 volunteer managers до начала dev.

---

## 12. Appendix: Demo Order Text (draft)

```
ПБ 78-12-8п 2
71-12-8 3
ПБ 66-12-8п 4
59-12-8 1
ПБ 87-12-10п 2
```

*Требует валидации domain expert перед релизом.*

---

## Связанные документы

- [`product-analysis-swot-ost-assumptions.md`](./product-analysis-swot-ost-assumptions.md)
- [`ux-wizard-step-plates.md`](./ux-wizard-step-plates.md)
- [`project-baseline.md`](./project-baseline.md)
