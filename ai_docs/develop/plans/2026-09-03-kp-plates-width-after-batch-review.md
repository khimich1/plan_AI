# Plan: KP Plates Width After Batch Review

**Created:** 2026-09-03  
**Spec:** [../../specs/kp-plates-width-after-batch-review.md](../../specs/kp-plates-width-after-batch-review.md)  
**Goal:** Закрыть финальные правки по двухэкранному поведению шага "Плиты" после batch review с обязательным тестовым покрытием до merge.  
**Status:** Ready for multitask execution  
**SDD:** SPECIFY ✅ · PLAN ✅ · IMPLEMENT ⏳

## Overview

Нужно довести изменения по ширинам до merge-ready состояния без расширения scope: подтверждаем корректный backend fallback для invalid width resolve, закрепляем frontend-поведение двух экранов шага "Плиты", фиксируем правила подсветок для batch review и подтверждаем UX вручную по "скриншоту 3".

## Locked Scope

1. Только поведение, описанное в `kp-plates-width-after-batch-review.md`.
2. Без изменения бизнес-логики вне resolve/matching + UI шага "Плиты".
3. Обязательные тесты (backend + frontend) выполняются до merge.
4. Ручной smoke UX проводится по целевому виду "скриншот 3".

## Tasks

### P1 - Backend fallback по размерам для invalid width

- **Type:** `feat-be`
- **Priority:** Critical
- **dependsOn:** []
- **pipeline:** explore -> worker -> test-writer -> test-runner -> reviewer
- **securitySensitive:** false
- **needsExplore:** true
- **needsArchitectureReview:** false
- **parallelSafe:** false
- **Complexity:** Moderate
- **Files/components:** `app/services/commercial_plate_resolve.py`, `tests/test_invalid_width_resolve.py`
- **Что делаем:** гарантируем, что resolve invalid width находит целевую строку не только exact match, но и через fallback сопоставление по размерам (когда `item.line` - display name, а в тексте compact mark).
- **Ожидаемые артефакты:**
  - обновленная функция сопоставления строки resolve,
  - backend unit/integration тесты на fallback по размерам,
  - зелёный прогон целевого backend test-файла.
- **Acceptance criteria:**
  - кейс `line="Плиты ПБ ..."` + `input_text="ПБ ... qty"` проходит и корректно переписывает строку;
  - существующие сценарии exact match остаются зелёными;
  - нет регрессии в replace-поведении resolve.

### P2 - Frontend two-screen behavior для шага "Плиты"

- **Type:** `feat-fe`
- **Priority:** Critical
- **dependsOn:** [P1]
- **pipeline:** explore -> worker -> test-writer -> test-runner -> reviewer
- **securitySensitive:** false
- **needsExplore:** true
- **needsArchitectureReview:** false
- **parallelSafe:** false
- **Complexity:** Moderate
- **Files/components:** `frontend/src/features/commercial-offer/components/steps/PlateInputStep.tsx`, `frontend/src/features/commercial-offer/components/steps/__tests__/PlateInputStep.test.tsx`
- **Что делаем:** закрепляем two-screen flow внутри шага "Плиты":
  - экран 1 (`pendingBatchReview=true`) - фокус на OCR сверке;
  - экран 2 (`pendingBatchReview=false`) - карточки решений и предпросмотр.
- **Ожидаемые артефакты:**
  - vitest тесты на двухэкранное поведение,
  - подтверждение, что "Список верен" не блокируется unresolved wide на экране сверки,
  - подтверждение, что после confirm карточки рендерятся на втором экране.
- **Acceptance criteria:**
  - на batch review экране width-cards скрыты;
  - post-review экран показывает карточки решений и preview;
  - gate "Готово, далее" работает по текущим правилам unresolved items.

### P3 - Тесты подсветок batch review

- **Type:** `ui`
- **Priority:** High
- **dependsOn:** [P2]
- **pipeline:** explore -> worker -> test-writer -> test-runner -> reviewer
- **securitySensitive:** false
- **needsExplore:** true
- **needsArchitectureReview:** false
- **parallelSafe:** false
- **Complexity:** Simple
- **Files/components:** `frontend/src/features/commercial-offer/lib/plateLineHighlights.ts`, `frontend/src/features/commercial-offer/hooks/useBatchReviewHighlights.ts`, `frontend/src/features/commercial-offer/lib/__tests__/plateLineHighlights.test.ts`
- **Что делаем:** закрепляем визуальные правила:
  - wide-only highlights на batch review;
  - invalid highlights скрыты на batch review.
- **Ожидаемые артефакты:**
  - vitest тест на wide-only подсветки,
  - vitest тест на отсутствие invalid highlight в batch review,
  - стабильный expected tooltip/message для wide в режиме сверки.
- **Acceptance criteria:**
  - batch review возвращает только wide highlight типы;
  - invalid_width highlight не попадает в результирующий map;
  - post-review поведение highlight не ломается.

### P4 - Тестовый прогон и merge gate

- **Type:** `chore`
- **Priority:** Critical
- **dependsOn:** [P1, P2, P3]
- **pipeline:** test-runner -> reviewer
- **securitySensitive:** false
- **needsExplore:** false
- **needsArchitectureReview:** false
- **parallelSafe:** false
- **Complexity:** Simple
- **Files/components:** тестовые наборы backend/frontend
- **Что делаем:** обязательный контрольный прогон перед merge.
- **Ожидаемые артефакты:**
  - журнал прогона обязательных тест-команд,
  - отметка pass/fail по каждой команде,
  - финальный чек "готово к merge".
- **Acceptance criteria:**
  - все обязательные команды из раздела "Test Commands" завершаются успешно;
  - нет красных тестов в целевых наборах;
  - typecheck frontend зелёный.

### P5 - Smoke checklist UX по "скриншоту 3"

- **Type:** `docs`
- **Priority:** High
- **dependsOn:** [P2, P3]
- **pipeline:** reviewer
- **securitySensitive:** false
- **needsExplore:** false
- **needsArchitectureReview:** false
- **parallelSafe:** false
- **Complexity:** Simple
- **Files/components:** smoke notes в plan/PR checklist
- **Что делаем:** вручную подтверждаем целевой UX после batch review.
- **Ожидаемые артефакты:**
  - заполненный smoke checklist,
  - фиксация наблюдений по "скриншоту 3",
  - решение go/no-go на merge.
- **Acceptance criteria:**
  - каждый пункт checklist проверен вручную;
  - расхождения с целевым экраном документированы;
  - при критичном расхождении merge блокируется.

## Dependencies Graph

`P1 -> P2 -> P3 -> P4`  
`P2 -> P5`  
`P3 -> P5`

## Test Commands

```bash
# Backend: invalid width resolve fallback
/home/username/Code/plan_web/.venv/bin/python -m pytest tests/test_invalid_width_resolve.py -q

# Frontend: two-screen behavior + highlights
cd frontend && npm test -- --run PlateInputStep.test.tsx plateLineHighlights.test.ts

# Frontend static checks
cd frontend && npm run typecheck
```

## Manual Smoke Checklist (Screenshot 3 UX)

1. На экране batch review видно фото + список, но не видно width decision cards.
2. Wide строки подсвечены, invalid подсветка на batch review отсутствует.
3. Кнопка "Список верен" доступна при unresolved wide.
4. После confirm открывается второй экран шага "Плиты" с карточками и preview.
5. Карточка "Нестандартная ширина" визуально и по структуре совпадает с целевым "скриншот 3".
6. Выбор варианта + "Применить" обновляет состав КП в preview.
7. После успешного применения соответствующий invalid alert/card исчезает.
8. "Готово, далее" активируется только при закрытых unresolved wide/invalid/unpriced.

## Merge Acceptance Criteria

- P1-P5 завершены и отмечены выполненными.
- Backend unit/integration тесты fallback по размерам - PASS.
- Frontend vitest по двухэкранному поведению - PASS.
- Frontend vitest по подсветкам batch review - PASS.
- Smoke checklist по "скриншоту 3" выполнен без blocker-расхождений.
- Typecheck frontend - PASS.

## Notes for Multitask Execution

- В каждую worker-задачу инжектить контекст из `plan-web-context`.
- На этапе orchestration не пропускать `test-writer`/`test-runner` в задачах P1-P3.
- Перед merge обязателен user checkpoint с результатами тестов и smoke.
