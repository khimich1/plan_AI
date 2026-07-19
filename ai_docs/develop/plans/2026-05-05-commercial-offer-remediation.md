# Plan: Commercial Offer module remediation

**Created:** 2026-05-05  
**Orchestration:** `orch-2026-05-05-12-00-commercial-remediation`  
**Source audit:** `ai_docs/develop/audits/2026-05-05-commercial-offer-scope-audit.md`  
**Goal:** Close P0 security findings (S-H01–S-H05), then highest-impact P1 architecture and frontend reliability work within a single 10-task cycle.  
**Total tasks:** 10  
**Priority order:** P0 security → P1 (API encapsulation, contract, FE state, header decoupling, phased refactor, tests)

---

## Tasks Overview (execution order)

| ID | Title | Priority | Depends on |
|----|-------|----------|------------|
| KP-001 | DraftStore path traversal hardening | P0 | — |
| KP-002 | Draft IDOR: ownership model and enforcement | P0 | KP-001\* |
| KP-003 | Upload limits, rate limits, and magic-byte validation | P0 | KP-002\*\* |
| KP-004 | Sanitize commercial API error responses | P0 | — |
| KP-005 | Public workflow API for generated files (replace `_resolve_generated_file`) | P1 | KP-004 |
| KP-006 | Formal wizard/orchestration response contract (A1) | P1 | KP-005 |
| KP-007 | Server-truth money + React Query v5 + POST /calculate in wizard | P1 | KP-006 |
| KP-008 | Decouple AppHeader from commercial wizard store | P1 | KP-007 |
| KP-009 | CommercialWorkflowService Phase 1 extraction (calculation slice) | P1 | KP-006 |
| KP-010 | Regression tests: security + commercial web flow | P1 | KP-001–KP-004 (min.) |

\*KP-002 can start after KP-001 only where paths touch the same code paths; if parallelized, complete KP-001 before merge.  
\*\*KP-003 should enforce limits only after authenticated ownership checks exist; sequence KP-002 before KP-003 in implementation.

---

## Dependencies graph

```
KP-001 ──┐
         ├── KP-002 ── KP-003
KP-004 ──┴── KP-005 ── KP-006 ─┬── KP-007 ── KP-008
                               └── KP-009
KP-001..004 ── KP-010 (after critical security merged)
```

---

## Task details

### KP-001 — DraftStore path traversal hardening (P0)

- **Audit:** S-H01  
- **Complexity:** Moderate  
- **Files:** `app/services/commercial_service.py` (DraftStore path joins); cross-check `frontend/.../draftStorage.ts` if filenames round-trip to the server.  
- **Actions:** Reject `..`, `/`, `\`, absolute segments in any user-influenced filename component; resolve under base and assert `relative_to(base)`; prefer deterministic names `{draft_id}_{field}.{ext}` where possible.  
- **Acceptance:** Unit/integration tests prove `../` and absolute paths cannot escape draft base; existing draft CRUD still works for valid names.

### KP-002 — Draft IDOR: ownership model and enforcement (P0)

- **Audit:** S-H02  
- **Complexity:** Moderate–Complex  
- **Files:** `app/schemas/commercial.py`, `app/services/commercial_service.py`, `app/api/v1/endpoints/commercial.py`; persistence layer where drafts are loaded (repository / file JSON).  
- **Actions:** Persist `user_id` (or equivalent tenant key) on draft create; add `verify_draft_ownership` dependency used by all draft-scoped routes (`GET/PATCH/POST ... /draft/{draft_id}`, uploads, calculate, downloads). Return 403/404 per project norm for cross-user access.  
- **Acceptance:** Automated test: user A cannot read or mutate user B’s `draft_id`; owner flows unchanged.

### KP-003 — Upload limits, rate limits, and magic-byte validation (P0)

- **Audit:** S-H03, S-H04  
- **Complexity:** Moderate  
- **Files:** `app/api/v1/endpoints/commercial.py`; optional `app/core/` config or existing settings for limits; optional rate-limit middleware (`slowapi`/Redis) if project already uses it.  
- **Actions:** Max body size per upload; per-user rate limits for OCR/upload endpoints; validate content via magic bytes (JPEG/PNG/PDF per audit), not `Content-Type` alone; safe storage path consistent with KP-001.  
- **Acceptance:** Tests: oversized body → 413; wrong signature → 400; exceeding rate → 429; valid file passes.

### KP-004 — Sanitize commercial API error responses (P0)

- **Audit:** S-H05  
- **Complexity:** Simple–Moderate  
- **Files:** `app/api/v1/endpoints/commercial.py`; optional shared exception helpers in `app/core/`.  
- **Actions:** No `detail=str(exc)` for unexpected errors; log full exception server-side; map known validation errors to stable client messages; generic 500 body.  
- **Acceptance:** Tests/assertions that internal messages do not appear in HTTP JSON for synthetic failures.

### KP-005 — Public workflow API for generated files (P1)

- **Audit:** A5  
- **Complexity:** Simple  
- **Files:** `app/services/commercial_workflow_service.py`, `app/api/v1/endpoints/commercial.py`.  
- **Actions:** Introduce public `get_or_generate_file(...)` (or equivalent); endpoint uses only public API; keep underscore helper private or inline.  
- **Acceptance:** No calls to `_resolve_generated_file` from router; behavior unchanged for authorized users.

### KP-006 — Formal wizard/orchestration response contract (P1)

- **Audit:** A1  
- **Complexity:** Complex  
- **Files:** `app/schemas/commercial.py`, `app/services/commercial_workflow_service.py`; `frontend/src/features/commercial-offer/types/commercialOffer.ts`; step flow (`wizardStepOrder.ts`, wizard container).  
- **Actions:** Pydantic response model: `current_step`, `can_proceed_to`, `next_required_action`, `validation_errors`; server is authoritative for step advancement; align FE `WizardStepId` with BE.  
- **Acceptance:** FE can gate “Next” using response fields; contract documented in code (types + OpenAPI).

### KP-007 — Server-truth money + React Query v5 + POST /calculate in wizard (P1)

- **Audit:** A3, Q1, Q2, Q3  
- **Complexity:** Complex  
- **Files:** `frontend/src/features/commercial-offer/` (hooks, `CommercialOfferWizard.tsx`, stores, mutations); `app/services/commercial_workflow_service.py` as needed for response shape.  
- **Actions:** Remove duplicate client total calculations where feasible; display `draft.totals` from server; replace invalid `useQuery` `onSuccess` with `useEffect` on `data` or mutation-first flow; call calculate when advancing steps / persisting per KP-006; mutations use stable patterns (`invalidateQueries` + derive state from query).  
- **Acceptance:** Manual or E2E: step transitions trigger calculate; UI shows server totals after sync; no reliance on deprecated RQ v5 callbacks.

### KP-008 — Decouple AppHeader from commercial wizard store (P1)

- **Audit:** A4  
- **Complexity:** Moderate  
- **Files:** `frontend/src/app/layout/AppHeader.tsx`; page/container for commercial offer route under `frontend/src/pages/commercial-offer-create/` or feature entry.  
- **Actions:** Header accepts callbacks/context for navigation/close/reset; wizard-specific wiring lives in page-level container.  
- **Acceptance:** AppHeader has no direct imports from commercial wizard store; other pages unaffected.

### KP-009 — CommercialWorkflowService Phase 1 extraction (P1)

- **Audit:** A2 (phased)  
- **Complexity:** Complex (scoped)  
- **Files:** `app/services/commercial_workflow_service.py`, new module e.g. `app/services/commercial_calculation_service.py` (name per project conventions), `app/api/v1/endpoints/commercial.py` (DI wiring).  
- **Actions:** Extract calculation/totals validation slice only in this cycle; leave file generation / wide-plate orchestration in workflow service for a future cycle.  
- **Acceptance:** Clear boundaries; tests still pass; no behavioral regression on calculate endpoint.

### KP-010 — Regression tests: security + commercial web flow (P1)

- **Audit:** (verification for S-H01–S-H05 + module health)  
- **Complexity:** Moderate  
- **Files:** `tests/test_commercial_web_flow.py`, fixtures as needed.  
- **Actions:** Cover path traversal, IDOR, upload validation/limits (where testable without heavy infra), error sanitization; extend existing flow tests for wizard contract if feasible.  
- **Acceptance:** `pytest` green; critical security cases cannot regress silently.

---

## Progress (orchestrator)

- ⏳ KP-001: DraftStore path traversal hardening  
- ⏳ KP-002: Draft IDOR ownership  
- ⏳ KP-003: Upload limits + magic bytes + rate limits  
- ⏳ KP-004: Error sanitization  
- ⏳ KP-005: Public file API  
- ⏳ KP-006: Wizard contract  
- ⏳ KP-007: Server truth + React Query + calculate  
- ⏳ KP-008: AppHeader decoupling  
- ⏳ KP-009: Workflow Phase 1 extraction  
- ⏳ KP-010: Regression tests  

---

## Architecture decisions (this cycle)

- **Security first:** No production deploy of partial fixes without KP-001–KP-004 addressed.  
- **Server authority:** KP-006 + KP-007 together establish server-driven wizard and pricing truth.  
- **A2 deferral:** Full god-module split is multi-sprint; KP-009 delivers one extracted slice only.  
- **Second cycle (out of scope here):** Remaining `CommercialWorkflowService` splits, medium items (A6–A8), optional AV scanning (S-H04 note).

---

## Verification

- Run `pytest tests/test_commercial_web_flow.py` (and full suite if time).  
- Manual smoke: create draft, upload sample image/PDF, calculate, download artifact, cross-user denial.
