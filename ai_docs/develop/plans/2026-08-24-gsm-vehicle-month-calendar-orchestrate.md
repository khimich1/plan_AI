# Промпт оркестратора: GSM calendar (скопировать в новое Agent-окно)

Ниже — самодостаточный стартовый промпт. Спека и план уже утверждены: не
переписывать скоуп, не «улучшать» дизайн, не идти в planner.

---

```
# Роль

Ты — оркестратор реализации мини-среза «ГСМ: календарь машины в раскрытой строке»
в проекте «Шишов» (FastAPI + React/TS). Phase 4 (Implement) spec-driven
development. Спека и план утверждены: не переписывай, не меняй скоуп,
не «улучшай» дизайн, не запускай planner заново.

# Контекст — прочитай ДО начала работы

1. `.cursor/skills/project-shishov/SKILL.md`
2. `ai_docs/specs/gsm-vehicle-month-calendar.md` — СПЕКА (источник истины)
3. `ai_docs/develop/plans/2026-08-24-gsm-vehicle-month-calendar.md` — ПЛАН: T1–T4
4. `.cursor/skills/test-driven-development/SKILL.md`
5. `.cursor/skills/incremental-implementation/SKILL.md`
6. `.cursor/skills/frontend-ui-engineering/SKILL.md`
7. Для T4: `.cursor/skills/browser-testing-with-devtools/SKILL.md`

# Миссия

Выполнить T1–T4 плана. После каждой задачи:
- её Verification-команда зелёная;
- отметь задачу [x] в файле плана (и acceptance-пункты).

Порядок: T1 → T2 → T3 → Checkpoint → T4.
T2+T3 не параллелить: оба трогают UI журнала / контракт календаря.
TDD: сначала красный тест, потом код.

# Срез

Фронтенд-only. В раскрытой строке обзора над лентой ПЛ — сетка дней периода.
Маркеры tx (`fuel`/`wash`) и ПЛ. Дыра = tx без ПЛ (warning). Красный бак =
`manual_intervention` (danger), это не дыра. ПЛ без заправки не красить.
Пустой период: сетка + «нет движений», без автопрыжка месяца.
Клик по ПЛ → существующий drawer. Клик по дыре → focus + scrollIntoView на
«Сгенерировать», generate не вызывать.

# Правила

- Только файлы из плана. `app/`, `core/`, схема БД, баннер `open_before` — не трогать.
- Даты ISO `YYYY-MM-DD`. Неделя с понедельника; слоты до Пн и до Вс — не даты вне [from, to].
- `useGsmWaybillsQuery` + `useGsmTransactionsQuery` (уже есть). Новых GET нет.
- Незакоммиченные изменения gsm-fleet-overview-ux и gsm-anchor-corridor не откатывать и не коммитить.
- Коммитов, push, git config нет.

# Окружение

- Репозиторий: `/home/roman/project/Шишов`
- Frontend: `cd frontend && npm test -- --run src/features/gsm/` ; `npm run build`
- Backend pytest не обязан меняться; не чинить 11 падений вне GSM.
- `run+logs.sh` / uvicorn на :8000 может быть уже запущен. Generate/export на рабочей `plita.db` в T4 не делать (только чтение UI).

# Ключевые файлы

- `frontend/src/features/gsm/components/VehicleWaybillJournal.tsx`
- `frontend/src/features/gsm/hooks/useGsmQueries.ts` (`useGsmTransactionsQuery`)
- `frontend/src/features/gsm/lib/waybillWarnings.ts` (`isProblematicDay`)
- `frontend/src/features/gsm/lib/fleetStatus.ts` (тона warning/danger)

# Definition of Done

- T1–T4 и Checkpoint [x] в плане.
- `cd frontend && npm test -- --run src/features/gsm/` зелёный.
- `cd frontend && npm run build` зелёный.
- Отчёт `ai_docs/develop/reports/2026-08-24-gsm-vehicle-month-calendar.md`.
- Финальное сообщение: список файлов + сводка тестов. Ничего не закоммичено.

Начни с чтения спеки/плана. Подтверди T1–T4 одной строкой на задачу. Затем T1.
```

---

## Как запустить в отдельном окне

1. **Новый чат Agent** (не продолжать текущий: там коридор бака, idea-refine и
   спека календаря — лишний шум).
2. Режим **Agent**, не Ask/Plan. Если нужен координатор с воркерами —
   включить **Multitask** в том окне (как на срезе коридора).
3. В поле ввода: `@` на три файла, затем вставить блок из тройных кавычек выше.
   - `ai_docs/specs/gsm-vehicle-month-calendar.md`
   - `ai_docs/develop/plans/2026-08-24-gsm-vehicle-month-calendar.md`
   - этот файл (не обязательно, если промпт уже вставлен целиком)
4. Не писать «сначала уточни спеку» — скоуп закрыт. Если агент начнёт planner —
   остановить и напомнить: Phase 4, план approved.
5. Это окно можно оставить для вопросов по ГСМ; реализацию вести только там,
   чтобы два агента не правили одни и те же файлы.
