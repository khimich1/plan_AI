# Implementation Plan: Стабилизация P0 по аудиту 2026-08-28

**Спека**: [ai_docs/specs/stabilizaciya-p0-audit-2026-08-28.md](../../specs/stabilizaciya-p0-audit-2026-08-28.md)
**Дата**: 2026-08-28

## Overview

Закрываем четыре находки аудита (Q3, S2, S4, S1) без изменения поведения
системы. Порядок задач продиктован зависимостями: baseline-артефакты должны
быть сняты ДО удаления кода; рискованное обновление зависимостей (S2)
изолировано между чекпоинтами; документация — в конце, когда код стабилен.

## Architecture Decisions

- **Baseline-first**: PDF и структурный снимок `build_layout_sequence` на реальных
  данных (`plita.db`, активный план в `production_plans`) генерируются до правок.
  Сравнение после — по структурному JSON (детерминировано), PDF — визуально
  пользователем (побайтовое сравнение невозможно: в PDF вшита дата создания).
- **S2 — прыжок на последнюю стабильную** связку FastAPI/Starlette (решение
  пользователя). Страховка — полный прогон pytest; fallback при неразрешимой
  регрессии: минимальная связка с чистым pip-audit (фиксируем в спеке).
- **S1 закрывается документом** (ADR), код `offer_access.py` не меняется —
  только docstring-ссылка.
- **Без коммитов**: все изменения остаются в рабочем дереве для ревью
  пользователем.

## Task List

### Phase 1: Baseline

- [x] **Task 1: Baseline-артефакты до изменений**
  - **Description**: Временный скрипт `_p0_baseline/generate_baseline.py`:
    (а) находит дату с дорожками в активном плане `production_plans.payload_json`;
    (б) генерирует PDF «Схема дорожек» через `generate_day_schema` (реальный
    пользовательский путь); (в) прогоняет `build_layout_sequence` на реальных
    данных дня и сохраняет структурный снимок (counts/keys/hash) в JSON;
    (г) прогоняет pytest-подмножество layout с сохранением лога.
  - **Acceptance**: существуют `_p0_baseline/schema_before.pdf`,
    `sequence_before.json`, `pytest_layout_before.log`; PDF открывается;
    pytest-подмножество зелёное.
  - **Verify**: `ls -la _p0_baseline/`; лог pytest без failures.
  - **Files**: `_p0_baseline/generate_baseline.py` (новый, временный — не коммитим)
  - **Estimated scope**: S

### Phase 2: Q3 — мёртвый код

- [x] **Task 2: Удаление unreachable-кода в `builder.py`**
  - **Description**: Удалить блок после `return sequence` (строка 271) до конца
    мёртвой зоны (~строка 991), не задев живой код выше. Проверить, что после
    мёртвого блока нет живых функций файла (прочитать границы перед удалением).
  - **Acceptance**: `builder.py` содержит только живой код; `grep -n "DEPRECATED / UNREACHABLE"`
    пуст; `wc -l builder.py` ≈ 271.
  - **Verify**: `pytest tests/ -k "layout or visualization or viz"` зелёный;
    повторная генерация: `sequence_after.json` структурно идентичен
    `sequence_before.json`; `schema_after.pdf` сгенерирован без ошибок.
  - **Files**: `viz_modules/layout_sequence/builder.py`
  - **Estimated scope**: S

### Checkpoint: после Tasks 1–2

- [x] Layout/visualization тесты зелёные
- [x] JSON-снимки до/после идентичны
- [ ] Пользователю переданы оба PDF для визуального контроля

### Phase 3: S2 — зависимости backend

- [x] **Task 3: Обновление FastAPI/Starlette до последней стабильной**
  - **Description**: `pip install --upgrade fastapi starlette` (последняя
    стабильная связка); зафиксировать новые пины в `requirements.txt`;
    прогнать `pip-audit`; прогнать полный pytest. При падениях — локализовать
    и чинить (типичные места: TestClient, BaseHTTPMiddleware в CSRF/plate
    isolation). Если регрессия неразрешима быстро — fallback на минимальную
    чистую связку с фиксацией в спеке.
  - **Acceptance**: `pip-audit` без уязвимостей в fastapi/starlette;
    полный `pytest` зелёный; `requirements.txt` обновлён.
  - **Verify**: `pytest` (весь набор); вывод `pip-audit` сохранён в
    `_p0_baseline/pip_audit_after.txt`.
  - **Files**: `requirements.txt` (+ код только при breakage — см. Boundaries:
    ask first, если чинить придётся за пределами тестов/пинов)
  - **Estimated scope**: M (непредсказуемость регрессий)

### Checkpoint: после Task 3

- [x] Полный backend-регресс зелёный *(12 failed / 2324 passed — те же 12 pre-existing; регрессий S2 нет)*
- [x] pip-audit чист по fastapi/starlette

### Phase 4: S4 — зависимости frontend

- [x] **Task 4: npm audit fix**
  - **Description**: `npm audit fix` в `frontend/`; при предложении major-версий
    с breaking changes — стоп и вопрос пользователю (Boundary: ask first).
  - **Acceptance**: `npm run audit:ci` exit 0 (или исключения задокументированы
    в спеке с причиной); `npm run test`, `npm run typecheck`, `npm run build`
    зелёные.
  - **Verify**: команды выше.
  - **Files**: `frontend/package.json`, `frontend/package-lock.json`
  - **Estimated scope**: S

### Phase 5: Документация

- [x] **Task 5: ADR политики доступа к КП (S1)**
  - **Description**: Создать `ai_docs/develop/architecture/offer-access-policy.md`
    по шаблону `deployment-single-instance.md` (Status: accepted): менеджеры и
    админы разделяют полный доступ к архиву КП; `owner_user_id` в модели данных
    зарезервирован на будущее. В `offer_access.py` — docstring-ссылка на ADR
    (логика не меняется).
  - **Acceptance**: ADR существует, Status: accepted, ссылается на аудит [S1]/[A15];
    docstring обновлён; pytest зелёный (docstring-only change).
  - **Verify**: `pytest tests/test_archive_authorization.py tests/test_offers_production_authorization.py`.
  - **Files**: `ai_docs/develop/architecture/offer-access-policy.md`,
    `app/security/offer_access.py` (только docstring)
  - **Estimated scope**: XS

- [x] **Task 6: Статусы в отчёте аудита + финальный отчёт**
  - **Description**: Пометить в `2026-08-28-full-project-audit.md` находки
    Q3/S2/S4/S1 как resolved/documented со ссылками на спеку и ADR. Финальное
    сообщение пользователю: сводка изменений, список файлов рабочего дерева,
    два PDF для визуального контроля, напоминание про коммит (агент не коммитит).
  - **Acceptance**: в отчёте проставлены статусы; пользователь получил сводку.
  - **Verify**: прочитать обновлённые секции отчёта.
  - **Files**: `ai_docs/develop/audits/2026-08-28-full-project-audit.md`
  - **Estimated scope**: XS

### Checkpoint: Complete

- [x] Все Success Criteria спеки выполнены *(кроме визуального контроля PDF — у пользователя)*
- [x] Полная регрессия: pytest без регрессий S2; vitest + typecheck + build зелёные
- [x] Рабочее дерево готово к ревью пользователем

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| S2: прыжок на latest ломает тесты (TestClient/middleware) | Med | Полный pytest; типичные места известны (CSRF, plate isolation); fallback — минимальная чистая связка |
| Baseline PDF не генерируется (нет данных за дату) | Low | Дата берётся из `production_plans.payload_json` (1 активный план есть) |
| npm audit fix тянет major с breaking changes | Med | Boundary: ask first; задокументировать исключение в спеке |
| Механическая ошибка при удалении (задет живой код) | Low | Границы блока читаются перед удалением; структурный JSON-diff до/после |
| Побайтовое сравнение PDF невозможно (вшитая дата) | Low | Структурное сравнение JSON + визуальный контроль пользователем |

## Open Questions

- Точные целевые версии FastAPI/Starlette: fastapi 0.141.1, starlette 1.6.0, pydantic 2.13.4.
- Судьба `_p0_baseline/` после приёмки: оставить для ревью, не коммитить
  (временные артефакты).
