# Spec: Архив — занести контрагента из 1С и привязать к КП

> **Источник идеи:** [`ai_docs/ideas/archive-counterparty-create-and-bind.md`](../ideas/archive-counterparty-create-and-bind.md)  
> **План:** [`ai_docs/develop/plans/2026-09-21-archive-counterparty-create-and-bind.md`](../develop/plans/2026-09-21-archive-counterparty-create-and-bind.md)  
> **Продолжает:** [`kp-counterparty-bind-on-production.md`](kp-counterparty-bind-on-production.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → IMPLEMENT  
> **Статус:** IMPLEMENT  
> **Дата:** 2026-09-21  
> **Связанные модули:** карточка архива, `archive.bind_counterparty`, справочник `counterparties`.  
> **Рядом, не это:** мастер КП, живой API 1С, импорт xlsx/GUID, обмен счетами, `POST /counterparties` из UI мастера.

---

## Assumptions I'm Making

1. **Менеджер уже завёл клиента в 1С.** В форму копируют реквизиты карточки, а не выдумывают код.
2. **Код 1С — якорь для будущего обмена.** GUID в MVP не вводим и не требуем. Счёт в 1С этим релизом не заработает.
3. **После create+bind** `customer_name` / ИНН / КПП **перезаписываются** из карточки справочника (как у текущего PATCH). PDF/XLSX не регенерируем.
4. Писать может тот же круг, что bind: `assert_offer_write_access` (автор КП / admin).
5. Схему SQLite **не трогаем.** `counterparties.guid_1c` остаётся nullable; `source=manual` уже есть.
6. **`POST /api/v1/counterparties` не закрываем** и из мастера по-прежнему не зовём. Создание из UI — только архив.
7. Предупреждение «такой ИНН уже есть» в архиве **не показываем**: не блокирует, отдельный конверт ответа не вводим.
8. Один combobox = имя. Отдельное поле «название» не дублируем.

→ Если что-то из этого неверно — править spec до кода.

---

## Decisions locked (ideation 2026-09-21)

| # | Тема | Решение |
|---|------|---------|
| D1 | Место | Только карточка архива, статус «в архиве». Мастер не менять |
| D2 | Поиск | Автокомплит как сейчас. Нашёл → PATCH `counterparty_id` |
| D3 | Нет в списке | Под combobox поля **код 1С** (обяз.), ИНН, КПП. Имя = текст combobox (обяз.) |
| D4 | Кнопка | Одна: выбран id → PATCH; иначе имя+код → POST create+bind |
| D5 | API создания | `POST /api/v1/commercial/archive/{kp_id}/counterparty` — атомарно insert + bind |
| D6 | Дубль `code_1c` | Не 409 наружу: если активный клиент — привязать его; иначе 400 как у bind |
| D7 | GUID | Не в форме, не в теле POST |
| D8 | Перепривязка | Пока «в архиве» — можно сменить (поиск или новое занесение). Отвязать нельзя |
| D9 | Файлы | Не регенерируем |
| D10 | Гейт производства | Без изменений: нужен активный `counterparty_id` |
| D11 | Имя в combobox | Префилл `customer_name`; можно поправить под юрлицо 1С до отправки |

---

## Objective

Менеджер в архиве заносит в наш справочник контрагента, которого уже создал в 1С, и сразу привязывает его к КП. Бейдж «нет 1С» снимается, «В производство» открывается. Это готовит якорь `code_1c` для будущего обмена, не сам обмен.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | менеджер | найти клиента в справочнике и занести его в КП | не создавать дубль, если выгрузка уже была |
| US-2 | менеджер | если поиска нет — вписать код 1С (и ИНН/КПП) к имени КП | карточка появилась в 1С, а в программе ещё нет |
| US-3 | менеджер | одним нажатием и создать, и привязать | не ловить «карточка есть, а КП всё ещё без 1С» |
| US-4 | система | при повторном коде 1С привязать существующего клиента | не плодить теневые строки |
| US-5 | система | не пускать в производство без активного клиента | гейт предыдущей спеки жив |

### Reframed success criteria

| Требование | Критерий |
|------------|----------|
| Поиск | Выбрали из списка → тот же PATCH, что сейчас; кнопка активна |
| Занесение | Имя + код 1С, без id → POST 200, строка в `counterparties` (`source=manual`, `is_client=1`), у КП заполнены `counterparty_id` и снапшот |
| Дубль кода | POST с чужим существующим кодом активного клиента → 200, **без** новой строки, КП привязан к найденному id |
| Не клиент / неактивен | POST с таким кодом → 400, КП без изменений |
| Не архив | POST/PATCH не «в архиве» → 400 |
| Без кода | POST без `code_1c` / без имени → 422, справочник и КП не трогаем |
| Мастер | По-прежнему нет «Добавить контрагента» / `createCounterparty` |
| Производство | После успешного POST кнопка «В производство» активна; `move_to_production` проходит `require_active_client` |

---

## Tech Stack

Тот же: FastAPI + Pydantic v2, SQLite `plita.db`, React 19 / TypeScript, vitest + pytest. Новых зависимостей нет.

## Commands

```bash
# backend
pytest tests/test_archive_counterparty_bind.py tests/test_archive_endpoints.py \
  tests/test_archive_service.py tests/test_counterparties_service.py \
  tests/test_counterparties_api.py -q

# frontend
cd frontend && npx vitest run \
  src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx \
  src/features/commercial-archive/api/archiveApi.test.ts \
  src/features/commercial-offer/components/CounterpartyAutocomplete.test.tsx
```

Точные имена новых тестов — в плане.

## Project Structure

```
app/schemas/archive.py                         # CreateAndBindCounterpartyRequest
app/api/v1/endpoints/archive.py                # POST /{kp_id}/counterparty
app/services/archive_service.py                # create_and_bind_counterparty
app/services/counterparties_service.py         # resolve-or-create по code_1c (без утечки дубля)
frontend/.../archiveApi.ts                     # createAndBindCounterparty
frontend/.../useArchiveQueries.ts              # мутация POST
frontend/.../OfferDetailsDrawer.tsx            # поля кода/ИНН/КПП, один submit
tests/test_archive_counterparty_bind.py        # create+bind, дубль, гейты
```

Мастер КП, `CounterpartyAutocomplete` как общий поиск, импорт xlsx — не в этом изменении (autocomplete только если понадобится проп «текст без выбора»).

## Code Style

Контракт POST — поля как у ручного создания, без GUID:

```python
class CreateAndBindCounterpartyRequest(BaseModel):
    name: str = Field(min_length=1, max_length=COUNTERPARTY_NAME_MAX_LENGTH)
    code_1c: str = Field(min_length=1, max_length=COUNTERPARTY_CODE_1C_MAX_LENGTH)
    inn: str | None = Field(default=None, max_length=COUNTERPARTY_INN_MAX_LENGTH)
    kpp: str | None = Field(default=None, max_length=COUNTERPARTY_KPP_MAX_LENGTH)
```

Сервис — одна транзакция на том же `db_path`, что архив и справочник:

```python
def create_and_bind_counterparty(self, kp_id: int, payload, *, user: dict) -> ArchiveOfferDetails:
    # статус «в архиве», write access — как bind_counterparty
    cps = CounterpartiesService(db_path=self.repository.db_path)
    row = cps.resolve_active_client_for_bind(
        name=payload.name, code_1c=payload.code_1c, inn=payload.inn, kpp=payload.kpp
    )
    # update_counterparty(...) снапшот из row, не из свободного имени КП
    return self.get_details(kp_id, user=user)
```

`resolve_active_client_for_bind`: есть `code_1c` → `require_active_client(existing.id)`; нет → `insert(..., source="manual", is_client=True)` и вернуть строку. Не писать свободное имя КП в справочник, если выбран id из поиска (тот путь — старый PATCH).

Сообщения по-русски, в том же тоне: «Контрагент не найден в справочнике», «Контрагент не отмечен как клиент в 1С», «Привязать контрагента можно только у КП в статусе «в архиве»».

Кнопка в drawer:

```ts
const canBindExisting = Boolean(bindSelected);
const canCreate =
  !bindSelected && name.trim().length > 0 && code1c.trim().length > 0;
// disabled={!canBindExisting && !canCreate}
```

## Testing Strategy

| Слой | Где | Что |
|------|-----|-----|
| Backend unit | `tests/test_archive_counterparty_bind.py` | POST создаёт+привязывает; дубль кода → bind без insert; supplier/inactive → 400; не архив → 400; пустые имя/код → 422 |
| Backend API | `tests/test_archive_endpoints.py` | POST зовёт сервис; PATCH без изменений |
| FE drawer | `OfferDetailsDrawer.test.tsx` | без выбора кнопка disabled, пока нет кода; с кодом → POST, не PATCH; с выбором из списка → PATCH как сейчас |
| FE api | `archiveApi.test.ts` | POST тела name/code_1c/inn/kpp |
| Регрессия | существующие bind/wizard тесты | мастер без create; гейт производства |

Не ходить в живую 1С. Не расширять импорт xlsx в этом MVP.

## Boundaries

- **Always:** атомарность create+bind; дубль кода не плодит строку; `require_active_client` на результате; мастер без UI создания; тесты на create, дубль, отказ.
- **Ask first:** обязательный ИНН; GUID в форме; закрыть `POST /counterparties` для manager; фильтр архива «нет 1С»; реген PDF; очередь занесения.
- **Never:** живой API 1С; фейковый `code_1c` с бэка; создание из шага «Клиент»; снимать гейт производства; менять схему SQLite; менять skip append/resume.

## Success Criteria

- [ ] КП «в архиве» без `counterparty_id`: имя в combobox + код 1С → POST → карточка в справочнике и привязка, бейдж «нет 1С» снят, «В производство» активно
- [ ] Тот же клиент уже в справочнике: выбор из списка → PATCH, новой строки нет
- [ ] POST с кодом существующего активного клиента → привязка к нему, count `counterparties` не растёт
- [ ] POST с кодом поставщика / неактивного → 400, КП без id
- [ ] В мастере по-прежнему нельзя завести контрагента
- [ ] Регрессии: текущий PATCH, `move_to_production` без id → 400, скидка/логистика архива

## Out of scope

- Живой поиск и создание в 1С
- GUID в UI и в теле занесения
- Доработка `scripts/import_counterparties.py` (колонка GUID) — отдельная задача обмена
- Фильтр/очередь «без 1С»
- Реген файлов после bind
- Смена контрагента у КП «в работе» / «выполнено»
- Закрытие прямого `POST /counterparties`

## Open Questions

Нет блокирующих. На ревью spec: если ИНН должен быть обязательным — сказать до плана; иначе оставляем как у текущего `POST /counterparties`.
