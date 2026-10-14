# Plan: Пример ФБС и мостовых свай из прайса

**Created:** 2026-09-14  
**Status:** ⏳ Waiting for spec/plan approval (SDD PLAN)  
**Idea:** [`ai_docs/ideas/kp-fbs-bridge-placeholder.md`](../../ideas/kp-fbs-bridge-placeholder.md)  
**Spec:** [`ai_docs/specs/2026-09-14-kp-fbs-bridge-placeholder.md`](../../specs/2026-09-14-kp-fbs-bridge-placeholder.md)

## Goal

Заменить copy-paste leftover `С120.35…` в плейсхолдерах шага 1 ФБС и мостовых свай на две реальные марки из текущего `pb.db` (или на пустую строку, если марок меньше двух). Менеджер и admin — один мастер. Парсер, OCR и расчёт не трогать.

## Architecture Decisions

- **D1.** Источник SKU — read-only инспекция `PRICE_DB_PATH` / `pb.db` **на IMPLEMENT**, не в рантайме. Строки замораживаются в `PRODUCT_TYPE_CONFIG`.
- **D2.** Таблицы: `fbs_prices`, `bridge_pile_prices`. Выбор независимый. Алгоритм — в спеке (ORDER BY mark, grade через `resolve_default_*`, qty 2 и 3, display-grade `B7.5` не `B7_5`).
- **D3.** &lt; 2 suitable marks → `""`. Нет фоллбека на сваи и нет pytest-марок `ФБС 9.3.6-Т` / `C8-35T1`, если их нет в этой БД.
- **D4.** Нет endpoint «дай пример», нет кнопки вставки, нет легенды, нет автокомплита, нет правок OCR.
- **D5.** Completeness-тест конфига больше не требует `placeholder.length > 0` у `fbs` / `bridge_piles`.

## Recorded SKUs (fill in Task 1)

Заполнить до правки конфига. Пока пусто — IMPLEMENT ещё не начинался.

| Type | DB path used | Distinct priced marks | Placeholder literal |
|------|----------------|----------------------|---------------------|
| `fbs` | `/home/username/Code/plan_web/pb.db` (`PRICE_DB_PATH`) | 14 | `ФБС 12.4.3-Т B25 2\nФБС 12.4.6-Т B25 3` |
| `bridge_piles` | `/home/username/Code/plan_web/pb.db` (`PRICE_DB_PATH`) | 114 | `C10-35B7 B25 2\nC10-35T1 B25 3` |

## Task List

### Phase 1: Inspect live price DB

- [ ] **PLACE-001: Inspect live `pb.db` and record 2+2 marks (or empty)** `(type: spike)` (⏳ Pending)

  **Description:** Open the same price DB the running stack uses (`core.project_paths.PRICE_DB_PATH`: env `PRICE_DB_PATH` \|\| `PB_DB_PATH` \|\| repo-root `pb.db`). For `fbs_prices` and separately `bridge_pile_prices`, list DISTINCT `mark` with `price > 0`, `ORDER BY mark`. If fewer than 2 marks, record `""`. Else take the first two marks; grade via `resolve_default_fbs_grade` / `resolve_default_bridge_pile_grade`; write display grade (`formatFbsGradeLabel` / `B25`/`B30`); qty 2 then 3. Use the stored `mark` column (bridge aliases already split — do not join with `;`). Paste the literals into the table above. Do not invent SKUs. Do not change product code in this task.

  **Acceptance criteria:**
  - [ ] Table «Recorded SKUs» filled for both types
  - [ ] Each non-empty line is `{mark} {grade} {qty}` from this DB
  - [ ] If count &lt; 2: literal is `""`, not a pile example

  **Verification:**
  - [ ] Re-run the inspect snippet from the spec; first two marks match the table
  - [ ] Optional sanity (not a product change): `get_fbs_price` / `get_bridge_pile_price` return a number for each chosen pair

  **Dependencies:** None

  **Files likely touched:**
  - `ai_docs/develop/plans/2026-09-14-kp-fbs-bridge-placeholder.md` (this recorded-SKU table only)

  **Estimated scope:** XS

### Phase 2: Config + pin tests

- [ ] **PLACE-002: Write placeholders into config and pin tests** `(type: feat-fe, dependsOn: PLACE-001)` (⏳ Pending)

  **Description:** Set `PRODUCT_TYPE_CONFIG.fbs.labels.placeholder` and `…bridge_piles…` to the literals from PLACE-001. Remove leftover comments («copy-paste leftover», «plan 2026-09-08 §5»). Update JSDoc on `placeholder` so empty is allowed for these two types. In `productTypeConfig.test.ts`: pin the new exact strings (or `""`); keep plates/piles/steps/marches pins unchanged; relax `placeholder.length > 0` so it does not fail on empty fbs/bridge. Do not touch `SimpleProductInputStep` (already binds `labels.placeholder`). Do not add a backend example endpoint.

  **Acceptance criteria:**
  - [ ] fbs and bridge placeholders ≠ `С120.35-12 B25 5\nС120.35-13и 3`
  - [ ] Pins match config literals from PLACE-001
  - [ ] Other four product placeholders unchanged
  - [ ] Completeness test allows `""` only for fbs/bridge_piles

  **Verification:**
  - [ ] `cd frontend && npx vitest run src/features/commercial-offer/lib/productTypeConfig.test.ts`
  - [ ] `cd frontend && npm run typecheck`

  **Dependencies:** PLACE-001

  **Files likely touched:**
  - `frontend/src/features/commercial-offer/lib/productTypeConfig.ts`
  - `frontend/src/features/commercial-offer/lib/productTypeConfig.test.ts`

  **Estimated scope:** S

### Checkpoint: after PLACE-001–002

- [ ] Vitest pin file green
- [ ] Typecheck clean
- [ ] No parser / OCR / pricing files in the diff

### Phase 3: Manual verify

- [ ] **PLACE-003: Manager and admin wizard step 1 smoke** `(type: chore, dependsOn: PLACE-002)` (⏳ Pending)

  **Description:** On the live local stack (`./run+logs.sh`), same route `/new` for both roles. Empty field on step 1 must show the new placeholder (or none). Paste as-is → «Обработать текст» → priced rows, no «не найдено в прайсе» on those example lines. If PLACE-001 recorded `""`, confirm the field has no `С120.35` hint.

  **Acceptance criteria:**
  - [ ] Manager + ФБС: paste → priced (or empty placeholder if `""`)
  - [ ] Manager + мостовые сваи: paste → priced (or empty)
  - [ ] Admin: same two types, same copy, same outcome
  - [ ] Placeholder visible only on an empty field (native HTML; draft text is out of scope)

  **Verification:**
  - [ ] Manual: `/new` as manager, then as admin (spec Commands § smoke 1–4)

  **Dependencies:** PLACE-002

  **Files likely touched:** none (verify only)

  **Estimated scope:** XS

### Checkpoint: Complete

- [ ] Spec success criteria S1–S7
- [ ] Ready for review; no commit unless asked

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Local `pb.db` has &lt; 2 FBS or bridge marks | Low | Empty placeholder is the locked behavior |
| Chosen `mark` form does not parse (rare alias spelling) | Med | Use stored `mark`; smoke PLACE-003; do **not** change the parser — pick another priced mark in ORDER BY order if the first fails parse **only if** it still follows «first two by mark». If first two fail parse, stop and ask — do not invent SKUs |
| Completeness test still requires length &gt; 0 | High | PLACE-002 explicitly relaxes it for fbs/bridge |
| Implementer copies pytest `ФБС 9.3.6-Т` / `C8-35T1` | High | Forbidden unless those exact marks exist in **this** DB as the first two |
| Scope creep: live API / insert button | Med | D4; not in task list |

## Open Questions

None.

## Implementation Notes

Execute only after human approval. Inject `plan-web-context`. Do not implement in the spec/plan session. Do not commit unless asked. Do not add `GET /example` or any commercial endpoint.
