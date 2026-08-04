# Skills в plan_web

Инженерные навыки агента лежат в **`.cursor/skills/`** (Cursor подхватывает только оттуда). Корневая папка `skills/` — устаревший дамп; не используйте её.

## Как пользоваться

### Автоматически
Агент сам выбирает skill по `description` в `SKILL.md`. Явно подключайте через `@` в чате, например `@plan-web-context`, `@test-driven-development`.

### Slash-команды проекта (предпочтительны для большой работы)

| Команда | Когда |
|---------|--------|
| `/orchestrate` | Большая фича: план → код → тесты → ревью → доки |
| `/implement` | Небольшая задача end-to-end |
| `/review` | Ревью перед коммитом/PR |
| `/refactor` | Упрощение/рефакторинг |
| `/audit` | Архитектура + security + качество |

### Фазы разработки (новые skills)

```
неясно что нужно     → interview-me / idea-refine
нужна спецификация   → spec-driven-development
разбить на задачи    → planning-and-task-breakdown  (или /orchestrate)
писать код           → incremental-implementation
  UI                 → frontend-ui-engineering
  API                → api-and-interface-design
тесты                → test-driven-development
баг                  → debugging-and-error-recovery
ревью                → code-review-and-quality  (или /review)
безопасность         → security-and-hardening
коммит/ветки         → git-workflow-and-versioning
доки/ADR             → documentation-and-adrs
релиз                → shipping-and-launch
```

Стартовый ориентир: `@using-agent-skills`. Контекст репозитория: `@plan-web-context`.

### Примеры промптов

- «Почини двойной учёт отходов в каскаде — через TDD»
- «`/orchestrate` реализация СГП по ТЗ из Task»
- «`@spec-driven-development` спецификация новых цен на плиты +доб»
- «`/review` изменения по прайсу»

## Карта

| Skill | Назначение |
|-------|------------|
| **plan-web-context** | Стек, пути, команды, доменные зоны риска |
| **using-agent-skills** | Роутер: какой skill когда |
| orchestration, simple-workflow, review/refactor/audit-workflow | Slash-workflow проекта |
| architecture-principles, code-quality-standards, security-guidelines, docs, git-helper, task-management | Базовые стандарты проекта |
| interview-me … shipping-and-launch | Фазовые engineering skills (адаптированы под FastAPI/React/pytest) |

## Правила

1. Перед нетривиальной работой — `plan-web-context`.
2. Большие задачи — `/orchestrate`, не ручной набор из 10 skills подряд.
3. Пруф: `pytest` и/или `cd frontend && npm run test` (+ `typecheck` при смене типов).
4. Планы/отчёты — пути из `.cursor/config.json` (`ai_docs/develop/...`).
5. Бот Telegram deprecated — не расширять без явной просьбы.

Подробнее про `.cursor`: `ai_docs/develop/cursor-dot-folder-guide-RU.md`.
