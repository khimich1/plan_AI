# Spec: Стабилизация P0 по аудиту 2026-08-28

**Статус**: implemented (2026-08-28); визуальный контроль PDF — у пользователя
**Связано**: [ai_docs/develop/audits/2026-08-28-full-project-audit.md](../develop/audits/2026-08-28-full-project-audit.md)
**Следующая спека**: stabilizaciya-p1-commercial-2026-08-28.md (декомпозиция коммерческого контура)

## Objective

Закрыть находки аудита с максимальным отношением ценность/риск, без изменения
поведения системы:

| Находка | Что делаем |
|---------|-----------|
| Q3 | Удалить ~720 строк unreachable-кода в `viz_modules/layout_sequence/builder.py` |
| S2 | Обновить FastAPI/Starlette до **последней стабильной** связки; pip-audit чист |
| S4 | `npm audit fix` по high-уязвимостям frontend |
| S1 | Политика «общий архив КП» — осознанная; фиксируем ADR (код доступа НЕ меняем) |

S9 (CSRF multipart) **исключён из P0** по решению пользователя — целиком
(и nginx-лимит, и код-фикс) перенесён в спеку фазы 2.

Успех: находки Q3/S2/S4/S1 сняты или задокументированы, регрессия зелёная,
поведение API и генерации PDF не изменилось.

## Принятые решения (из Q&A)

1. **S2**: прыжок на последнюю стабильную версию (не минимальную). Регрессионная
   страховка — полный прогон pytest + vitest.
2. **PDF-проверка (Q3)**: pytest + программная генерация «Схемы дорожек» до/после
   удаления + визуальный контроль пользователем двух PDF.
3. **Git**: изменения остаются в рабочем дереве, коммитов агент не делает —
   ревью и коммит выполняет пользователь.
4. **S9**: исключён из P0 по решению пользователя (2026-08-28); nginx-лимит
   и код-фикс CSRF — в спеке фазы 2. На заметку для фазы 2: в nginx-конфигах
   репозитория нет `client_max_body_size` (дефолт 1 МБ несовместим с
   backend-капом 50 МБ — вероятен config drift с боевым конфигом).

## Tech Stack

- Backend: Python 3.12.3, FastAPI (пин в `requirements.txt`), pytest
- Frontend: React 19.2.7 + TypeScript, Vite 8, Vitest

## Commands

Backend (cwd=корень):

```bash
pytest                              # addopts уже содержит --ignore=tests/archived
pip-audit                           # требует сеть; ручной шаг
```

Frontend (cwd=frontend):

```bash
npm run test                        # vitest run
npm run typecheck                   # tsc --noEmit
npm run build                       # tsc -b && vite build
npm run audit:ci                    # npm audit --audit-level=high
```

Ручной smoke (Q3): визуальное сравнение PDF «Схема дорожек» до/после
(файлы генерирует агент, сравнивает пользователь).

## Project Structure (затрагиваемые пути)

```
viz_modules/layout_sequence/builder.py   → удаление мёртвого кода (Q3)
requirements.txt                          → пины fastapi/starlette (S2)
frontend/package.json, package-lock.json  → npm-фиксы (S4)
ai_docs/develop/architecture/offer-access-policy.md → новый ADR (S1)
app/security/offer_access.py              → docstring-ссылка на ADR (S1)
ai_docs/develop/audits/2026-08-28-full-project-audit.md → статусы находок
tests/                                    → только прогон, новых тестов не требуется
```

## Code Style

Существующие конвенции проекта: комментарии и сообщения об ошибках на русском;
без новых комментариев, пересказывающих код; conventional commits оставляем
пользователю.

## Testing Strategy

Новых тестов не добавляем (поведение не меняется). Safety net:

- pytest (весь набор) — особенно `test_layout_*`, `test_visualization.py`,
  `test_core_viz_import_boundary.py` после Q3
- vitest + tsc --noEmit + vite build после S4
- программный smoke PDF до/после Q3 (генерация без ошибок, сопоставимый размер)
- визуальный контроль PDF пользователем

## Boundaries

- Always: прогон pytest после каждой backend-задачи; vitest+typecheck после frontend-задачи
- Always: baseline PDF генерируется ДО удаления кода в builder.py
- Ask first: если pip-audit потребует обновить что-то кроме fastapi/starlette
- Ask first: если npm audit fix тянет major-версию с breaking changes
- Never: трогать живой код builder.py выше строки 271 (`return sequence`)
- Never: менять логику `offer_access.py` (S1 закрывается документом, не кодом)
- Never: коммитить, пушить, менять git-конфиг
- Never: отключать/удалять падающие тесты

## Success Criteria

- [x] `builder.py` ≤ ~271 строки; `grep "DEPRECATED / UNREACHABLE"` пуст; layout pytest 51 passed
- [ ] Baseline и итоговый PDF сгенерированы; пользователь подтвердил визуальную идентичность
- [x] pip-audit не содержит уязвимостей в starlette/fastapi (fastapi 0.141.1 / starlette 1.6.0)
- [x] `npm run audit:ci` exit 0; `npm run test` + `typecheck` + `build` зелёные.
      Исключение: uuid@8.3.2 через exceljs (moderate; фикс — breaking-даунгрейд exceljs)
- [x] `ai_docs/develop/architecture/offer-access-policy.md` создан (Status: accepted);
      в `offer_access.py` docstring — ссылка на ADR
- [x] Отчёт аудита помечен: Q3/S2/S4/S1 → resolved/documented

## Open Questions

- ~~Точные целевые версии FastAPI/Starlette~~ — fastapi 0.141.1, starlette 1.6.0, pydantic 2.13.4
- pip 24.0: 7 CVE, фикс 25.3+ — ask-first, не трогали
- 12 pre-existing падений pytest на HEAD (не регрессия S2) — вне скоупа P0
