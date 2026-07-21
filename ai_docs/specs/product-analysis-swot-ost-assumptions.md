# Product Analysis: SWOT, OST, Risky Assumptions — система «Шишов»

> **Тип:** продуктовый анализ  
> **Дата:** 2026-06-19  
> **Статус:** черновик на ревью  
> **Связанные документы:** [`project-baseline.md`](./project-baseline.md), [`strategic-roadmap-pb-pricing-optimizer-1c.md`](./strategic-roadmap-pb-pricing-optimizer-1c.md), [`prd-onboarding.md`](./prd-onboarding.md)

---

## 1. Краткое описание продукта

**«Шишов»** — внутренняя веб-система автоматизации завода железобетонных изделий (ЖБИ), в первую очередь многопустотных плит перекрытия (ПБ). Продукт закрывает цепочку от коммерческого предложения (КП) до производственного планирования и документов смены.

### Основные возможности

| Модуль | Назначение | Пользователи |
|--------|------------|--------------|
| Wizard КП | Создание КП: ввод/OCR/AI → расчёт → PDF/XLSX | `manager`, `admin` |
| Архив КП | Поиск, скидки, логистика, перенос в производство | `manager`, `admin` |
| Производство | Планы, календарь, треки, завершение смены | `production`, `admin` |
| Telegram-бот | Паритет ключевых сценариев для менеджеров | `manager`, `admin` |
| Оптимизация | 2D ILP-раскладка плит (PuLP + CBC) | backend, производство |

### Технический контур

- **Frontend:** React SPA (Vite, TypeScript), маршруты `/login`, `/new`, `/archive`, `/production`
- **Backend:** FastAPI, SQLite (`plita.db`, `pb.db`), HMAC session cookie
- **Wizard КП:** 5 шагов — `plates` → `wide-plates` → `manager` → `client` → `result`

### Текущая проблема онбординга

После входа (`LoginPage.tsx`, `DEFAULT_REDIRECT = "/new"`) новый менеджер сразу попадает в wizard без приветствия, подсказок или демо-данных. Первый шаг (`PlateInputStep`) требует знания формата марок ПБ (`ПБ 78-12-8п 2`), понимания разницы между «Распознать» и «Обработать», а также доменных понятий (широкие плиты, нормализация). Сложность домена и отсутствие guided tour создают высокий порог входа.

---

## 2. SWOT-анализ

### Strengths (сильные стороны)

| # | Сила | Обоснование |
|---|------|-------------|
| S1 | **End-to-end покрытие процесса** | КП → архив → производство → документы смены в одной системе; меньше ручного переноса данных между Excel и цехом |
| S2 | **Гибкий ввод заказа** | Текст, OCR (GPT Vision), AI-инструкции, paste изображения из буфера — снижают ручной ввод для опытных пользователей |
| S3 | **Оптимизация раскладки** | ILP-оптимизатор с `verify_coverage`, audit trail — конкурентное преимущество vs Excel-калькуляторы |
| S4 | **Telegram-бот** | Менеджеры могут работать вне офиса; единый доменный слой `core/` |
| S5 | **Черновики и восстановление** | DraftStore + wizard state в браузере; можно прервать и продолжить |
| S6 | **Обширное тестовое покрытие backend** | pytest по парсингу, оптимизации, commercial flow — база для рефакторинга |

### Weaknesses (слабые стороны)

| # | Слабость | Обоснование |
|---|----------|-------------|
| W1 | **Нет онбординга** | Login → `/new` без welcome, demo, tooltips; новый менеджер «брошен» на шаг 1 |
| W2 | **Сложный домен на первом экране** | Форматы строк (`ПБ 78-12-8п`, `71-12-8`, L×W×H мм), широкие плиты, нормализация — неочевидны без обучения |
| W3 | **Размытая UX-модель шага 1** | Кнопки «Распознать», «Распознать (заменить)», «ИИ», «Обработать» — разные mental models; footer «Обработать» появляется только после draft |
| W4 | **Расхождения ценообразования** | Несколько путей расчёта цены (`commercial_offer.py`, `price_db.py`, `price_utils.py`); silent fallback `площадь × 4000 ₽` |
| W5 | **Клиент vs сервер — дублирование логики** | Итоги на фронте и в `calculate_draft()` могут расходиться (аудит A3) |
| W6 | **Архитектурный долг** | Монолитный `CommercialWorkflowService`, inverted dependency `core → app`, глобальное состояние оптимизатора |
| W7 | **Нет design system / Figma** | Inline styles, нет единых паттернов для onboarding-компонентов |

### Opportunities (возможности)

| # | Возможность | Обоснование |
|---|-------------|-------------|
| O1 | **Guided onboarding → быстрый time-to-first-KP** | Цель ≤15 мин для нового менеджера без помощи коллеги — измеримый KPI adoption |
| O2 | **Demo draft с типовым заказом** | Снижает страх «сломать»; показывает полный путь wizard до результата |
| O3 | **Role-based landing** | `production` → `/production`, `manager` → welcome + wizard; меньше путаницы |
| O4 | **Phase 1 roadmap (единый pricing)** | Прозрачная разбивка цены в preview → доверие к КП |
| O5 | **Inline hints на шаге plates** | Примеры формата, live validation, collapsible help — без отдельного обучающего PDF |
| O6 | **Интеграция с 1С (Phase 3)** | Устранение двойного ввода; усиление ценности для руководства |
| O7 | **Расширение на другие ЖБИ (Phase 4)** | Рост TAM внутри холдинга / других заводов |

### Threats (угрозы)

| # | Угроза | Обоснование |
|---|--------|-------------|
| T1 | **Отказ от системы новыми менеджерами** | Высокий порог входа → возврат к Excel/WhatsApp; onboarding failure = churn внутри завода |
| T2 | **Недоверие к суммам КП** | Расхождения pricing paths → менеджеры перепроверяют вручную, обесценивая автоматизацию |
| T3 | **Потери плит в оптимизаторе** | `lost_plates` как warning, не block → конфликт с производством, репутационный удар |
| T4 | **Зависимость от OpenAI для OCR** | Без ключа OCR отключён; rate limit in-memory не масштабируется |
| T5 | **SQLite + single instance** | Нет горизонтального масштабирования; filelock на черновиках → 503 при нагрузке |
| T6 | **Security findings в аудитах** | Default `APP_SECRET_KEY`, `secure=False` cookie, path traversal в DraftStore (частично закрыто) |
| T7 | **Scope creep** | Параллельные фазы (цены, оптимизатор, 1С, новые изделия) без фокуса замедляют UX-улучшения |

---

## 3. Opportunity Solution Tree — улучшение онбординга

### Desired Outcome (желаемый результат)

> **Новый менеджер создаёт первое корректное КП за ≤15 минут без помощи коллеги.**

«Корректное» = распознанный заказ без критических `unparsed_lines`, пройдены шаги wizard, получен preview с ценами, сохранено в архив или БД.

```
                    ┌─────────────────────────────────────────────────────┐
                    │ OUTCOME: Первое корректное КП ≤15 мин без помощи    │
                    └──────────────────────────┬──────────────────────────┘
                                               │
         ┌─────────────────────────────────────┼─────────────────────────────────────┐
         │                                     │                                     │
         ▼                                     ▼                                     ▼
┌─────────────────┐                 ┌─────────────────────┐               ┌─────────────────────┐
│ O1: Понимает    │                 │ O2: Быстро вводит   │               │ O3: Доверяет         │
│ что делать      │                 │ заказ на шаге 1     │               │ результату расчёта   │
│ после login     │                 │                     │               │                     │
└────────┬────────┘                 └──────────┬──────────┘               └──────────┬──────────┘
         │                                     │                                     │
    ┌────┴────┐                           ┌────┴────┐                          ┌────┴────┐
    ▼         ▼                           ▼         ▼                          ▼         ▼
 S1.1      S1.2                        S2.1      S2.2                       S3.1      S3.2
 Welcome   Role-based                  Format    Demo draft                 Price     Success
 screen    landing                     hints     prefill                    breakdown celebration
```

### Opportunities → Solutions → Experiments

| Opportunity | Solution | Experiment | Метрика успеха |
|-------------|----------|------------|----------------|
| **O1:** Понимает, что делать после login | **S1.1:** Welcome screen с 3 шагами «что будет дальше» + CTA «Создать первое КП» / «Посмотреть демо» | A/B: welcome vs текущий редирект на `/new` | % пользователей, начавших wizard в первые 2 мин после login |
| **O1** | **S1.2:** Role-based landing (`manager` → welcome, `production` → `/production`) | Измерить bounce rate production-роли на `/new` до/после | Drop-off production на `/new` → 0 |
| **O2:** Быстро вводит заказ | **S2.1:** Inline format hint + collapsible «Как писать марки ПБ» на `PlateInputStep` | Usability test: 3 новых менеджера, time-on-task | Median time step 1 ≤ 5 мин |
| **O2** | **S2.2:** Demo draft с типовым заказом (5–7 позиций) | «Попробовать на примере» → prefill textarea | % первых сессий с demo; completion rate demo → save |
| **O2** | **S2.3:** Упростить CTA: один primary «Продолжить» вместо «Распознать» + «Обработать» | Prototype test | Ошибки «забыл нажать Обработать» → 0 |
| **O3:** Доверяет результату | **S3.1:** Preview table с breakdown (база + резы) на шаге 1 | Показать 3 менеджерам vs текущий preview | Self-reported confidence ≥ 4/5 |
| **O3** | **S3.2:** Success screen после первого save + ссылка на архив | Track first-save event | Repeat usage D7 ≥ 50% |

### Priority Matrix

| ID | Solution | Priority | Rationale |
|----|----------|----------|-----------|
| S1.1 | Welcome screen | **P0** | Минимальный код, максимальный эффект на orientation; блокирует всё остальное |
| S1.2 | Role-based landing | **P0** | Исправляет явный баг UX для `production`; 1 redirect в `LoginPage` / router |
| S2.2 | Demo draft prefill | **P0** | Снижает friction шага 1; не требует backend changes |
| S2.1 | Format hints на plates | **P0** | Highest-friction step; см. [`ux-wizard-step-plates.md`](./ux-wizard-step-plates.md) |
| S2.3 | Unified CTA «Продолжить» | **P1** | Требует refactor `PlateInputStep` + wizard handler; высокий impact |
| S3.1 | Price breakdown в preview | **P1** | Зависит от Phase 1 pricing; можно stub с текущими данными |
| S3.2 | Success celebration | **P2** | Nice-to-have; после MVP onboarding |
| — | Guided tour (coach marks) | **P2** | v2; после стабилизации MVP flows |
| — | Video / PDF обучение | **P2** | Дополнение, не замена in-app hints |

---

## 4. Risky Assumptions (рискованные допущения)

Категории по типу риска для онбординга и продукта в целом.

### Value (ценность)

| # | Допущение | Риск | Как проверить |
|---|-----------|------|---------------|
| V1 | Новые менеджеры **хотят** создавать КП в системе, а не в Excel | Сопротивление изменениям; «Excel быстрее для простых заказов» | Интервью 5 менеджеров; shadowing 2 смен |
| V2 | Time-to-first-KP ≤15 мин — **достаточный** порог adoption | Может быть слишком амбициозно при OCR-ошибках | Замер baseline у 3 новых без подсказок |
| V3 | Demo draft **не создаёт** ложного ощущения «всё всегда так просто» | Разочарование на реальном заказе с wide plates | A/B + follow-up interview через неделю |

### Usability (юзабилити)

| # | Допущение | Риск | Как проверить |
|---|-----------|------|---------------|
| U1 | Inline hints **достаточны** без очного обучения | Домен слишком сложен для self-service | Usability test 5 участников |
| U2 | Менеджеры **понимают** разницу «исходный текст» vs «нормализованный» | Путаница → некорректные правки | Think-aloud на шаге 1 |
| U3 | OCR/AI **ускоряет** onboarding, а не пугает | Страх «ИИ ошибётся»; зависимость от API key | Track OCR vs text-only first sessions |

### Feasibility (реализуемость)

| # | Допущение | Риск | Как проверить |
|---|-----------|------|---------------|
| F1 | Welcome + hints — **frontend-only** MVP без backend | Demo draft может потребовать server-side template | Spike 1 день: static prefill vs API |
| F2 | Role-based redirect **не ломает** deep links и bot parity | Production user bookmarked `/new` | E2E test + bot role matrix |
| F3 | Unified CTA совместим с `wizard_state.can_proceed_to` | Сервер требует двухфазный ingest | Integration test commercial flow |

### Viability (жизнеспособность)

| # | Допущение | Риск | Как проверить |
|---|-----------|------|---------------|
| Vi1 | ROI онбординга **выше** Phase 1 pricing | Руководство приоритизирует «цифры в КП» | Stakeholder alignment meeting |
| Vi2 | Churn новых менеджеров **измерим** | Нет analytics events сейчас | Добавить minimal event logging |
| Vi3 | 15 мин **коррелирует** с retention | Может быть vanity metric | Cohort: first-KP-time vs D30 usage |

### Go-to-Market (внутреннее внедрение)

| # | Допущение | Риск | Как проверить |
|---|-----------|------|---------------|
| G1 | Завод **выделит** 2 часа на обучение + pilot | «Некогда» → провал adoption | Pilot с 2 менеджерами + champion |
| G2 | Admin **создаст** учётки с правильными ролями | Production попадает на wizard | Audit users table + login logs |
| G3 | Telegram-бот **не нужен** в первом онбординге | Менеджеры начинают с бота, не web | Survey channel preference |

### Top-5 Leap of Faith (самые рискованные)

| Rank | ID | Допущение | Почему leap of faith | Эксперимент (1–2 недели) |
|------|-----|-----------|----------------------|--------------------------|
| 1 | **V1** | Менеджеры перейдут с Excel на wizard | Без этого весь onboarding бессмысленен | 5 интервью + 1 неделя shadowing; metric: % КП через систему vs Excel |
| 2 | **U1** | Self-service onboarding без тренера | Если false — нужен очный onboarding program | Usability test 3 новых; success: 2/3 без помощи |
| 3 | **V2** | ≤15 мин — реалистичный target | Завышенная цель демотивирует команду | Baseline замер 5 сессий «как сейчас» |
| 4 | **Vi1** | Onboarding важнее pricing fix для adoption | Конфликт ресурсов с Phase 1 roadmap | Опрос руководства: «что блокирует использование?» |
| 5 | **F3** | UX-упрощение шага 1 не конфликтует с server wizard | Может потребовать backend contract change | Spike: single-button flow + `test_commercial_web_flow.py` |

---

## 5. Дополнительные рекомендации

### 5.1. Продуктовая roadmap (связь с onboarding)

| Горизонт | Фокус | Связь с onboarding |
|----------|-------|-------------------|
| **Q3 2026 — MVP onboarding** | Welcome, role landing, demo draft, hints на plates | PRD: [`prd-onboarding.md`](./prd-onboarding.md) |
| **Q3–Q4 2026 — Phase 1 pricing** | Единый `core/pricing`, golden tests, warning на fallback | Preview на шаге 1 показывает breakdown → доверие |
| **Q4 2026 — Phase 2 optimizer** | Hard gate на `verify_coverage`, block plan при потерях | Меньше «сюрпризов» на шаге result для новичков |
| **2027 — Phase 3 1С** | Контракт обмена после ответа 1С-специалистов | Onboarding v2: «КП автоматически уходит в 1С» |
| **2027+ — Phase 4–5** | Новые изделия, multi-factory | Пересмотр wizard step 0 «тип изделия» |

### 5.2. Технические риски из аудитов (релевантные onboarding)

| Риск | Источник | Влияние на onboarding | Митигация |
|------|----------|----------------------|-----------|
| Расхождение client/server wizard state | A1 commercial-offer audit | UI переходит без POST → ошибки на result | Опираться на `can_proceed_to`; не skip calculate |
| Дублирование totals client/server | A3 | Новичок видит «не те» суммы → потеря доверия | Показывать только `draft.totals` с сервера |
| Default APP_SECRET_KEY | Full project audit S1 | Security incident → простой системы | Fail fast в prod без секрета |
| OCR rate limiter in-memory | S5 | 503 при pilot onboarding группы | Документировать лимит; queue message |
| Path traversal DraftStore | S-H01 (частично закрыто) | Утечка черновиков при pilot | Verify `verify_draft_ownership` на всех paths |
| Монолитный CommercialWorkflowService | A2 | Медленные итерации onboarding API | Frontend-only MVP где возможно |

### 5.3. Позиционирование (внутреннее)

**Для менеджера:** «Создай КП за 15 минут — система сама посчитает цену, нарежет плиты и подготовит PDF для клиента.»

**Для производства:** «План = заказ. Ни одна плита не теряется.»

**Для руководства:** «Один источник правды: от КП до смены, с аудитом и будущей интеграцией 1С.»

**Anti-positioning (чего не обещать в onboarding):** «Замена 1С», «Работает без интернета», «OCR всегда точен».

### 5.4. Метрики успеха (North Star и input metrics)

| Метрика | Тип | Target (6 мес после MVP onboarding) | Instrumentation |
|---------|-----|-------------------------------------|-----------------|
| **Time-to-first-KP** | Input | Median ≤ 15 мин | Event: login → first `save` |
| **First-session completion rate** | Input | ≥ 60% | Event: wizard start → save/skip |
| **Step 1 drop-off** | Input | ≤ 25% | Event: plates view → plates process |
| **OCR vs text split (first session)** | Diagnostic | Track only | Event: createDraft source |
| **D7 repeat usage (new managers)** | Outcome | ≥ 50% | Event: second KP within 7 days |
| **Support tickets «как создать КП»** | Outcome | −70% vs baseline | Helpdesk tag |
| **КП через систему / всего КП** | North Star proxy | ≥ 80% на pilot-группе | Manual + archive count |

### 5.5. Следующие шаги

1. Ревью этого документа + [`prd-onboarding.md`](./prd-onboarding.md) с product owner и 1–2 менеджерами.
2. Baseline замер time-to-first-KP «как сейчас» (3–5 новых пользователей).
3. MVP onboarding (P0 items) → pilot 2 недели → итерация по метрикам.
4. Параллельно не блокировать Phase 1 pricing — preview breakdown в onboarding v1.1.

---

## Связанные артеfactы

- [`prd-onboarding.md`](./prd-onboarding.md) — детальный PRD
- [`ux-wizard-step-plates.md`](./ux-wizard-step-plates.md) — deep-dive шага 1
- [`project-baseline.md`](./project-baseline.md) — архитектурный baseline
- [`strategic-roadmap-pb-pricing-optimizer-1c.md`](./strategic-roadmap-pb-pricing-optimizer-1c.md) — стратегическая roadmap
