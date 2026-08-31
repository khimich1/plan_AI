# Project Audit Report

**Date**: 2026-08-03
**Scope**: full project (`app/`, `core/`, `frontend/src/` — critical ~20–30%)
**Audited by**: senior-reviewer + security-auditor + reviewer

---

## Executive Summary

**Overall Health Score**: 2.0/10

| Severity | Architecture | Security | Code Quality | Total |
|----------|-------------|----------|--------------|-------|
| Critical | 2           | 0        | 0            | **2** |
| High     | 6           | 3        | 7            | **16** |
| Medium   | 7           | 9        | 11           | **27** |
| Low      | 3           | 5        | 4            | **12** |

**Health Score calculation**:

- Start: **10**
- Critical: 2 × −2 = **−4** (cap −6)
- High: 16 × −0.5 capped at **−3**
- Medium: 27 × −0.1 capped at **−1**
- Low: ignored
- **Score = 10 − 4 − 3 − 1 = 2.0/10**

**Recommendation**: Address 2 critical architecture issues before next release; also prioritize High security findings S1–S3 (especially S3 inventory integrity).

---

## Critical Issues (fix immediately)

### [A1] ShipmentService is a god module mixing persistence, domain logic, and API mapping

| Field | Value |
|-------|-------|
| **Category** | Architecture |
| **Location** | `app/services/shipment_service.py` (~1521 lines, 40+ methods, ~96 direct DB touchpoints) |
| **Impact** | Single module owns CRUD, propose/confirm, packing, complete/cancel with SGP side-effects, XLSX export, KP/pile search, and schema mapping. Violates SRP; impossible to test subsystems in isolation; any logistics change risks wide regressions |
| **Fix** | Split into `ShipmentRepository`, `ShipmentProposeService`, `ShipmentCompletionService`, `ShipmentExportService`; thin orchestrator; map Pydantic only in endpoints |

### [A2] Stateful storage architecture is single-instance-only with no hard enforcement

| Field | Value |
|-------|-------|
| **Category** | Architecture |
| **Location** | `core/kp_db_common.py` (SQLite+WAL), `app/services/draft_store.py` (filesystem drafts), `core/config/settings.py` (`APP_STORAGE_LAYOUT=single_instance`, warn-only at ~450–458), in-process rate limits (`app/security/login_rate_limit.py`) |
| **Impact** | SQLite + filesystem drafts + in-process counters work only on a single instance. ADR exists but startup only warns. Multi-worker or multi-replica deployment causes data races, lost drafts, and bypassed rate limits |
| **Fix** | Enforce `UVICORN_WORKERS=1` / replica count at startup in production; or migrate drafts + counters to shared store before multi-instance |

---

## High Priority Issues (fix soon)

### Architecture

#### [A3] CommercialWorkflowService god module + Service Locator

| Field | Value |
|-------|-------|
| **Category** | Architecture |
| **Location** | `app/services/commercial_workflow_service.py` (~1039 lines); constructor instantiates 10+ services inline without DI |
| **Impact** | Hidden coupling, untestable without full dependency graph, violates DIP |
| **Fix** | Decompose by use-case; constructor injection via `app/dependencies/services.py` |

#### [A4] Application services return Pydantic HTTP schemas

| Field | Value |
|-------|-------|
| **Category** | Architecture |
| **Location** | `shipment_service`, `sgp_service`, `carrier_service`, `archive_service` import/return `app/schemas/*` |
| **Impact** | Transport layer (HTTP schemas) mixed with application/domain layer; API contract changes ripple into business logic |
| **Fix** | Return domain models from services; map to Pydantic schemas only in endpoints |

#### [A5] Logistics/SGP SQL lives in services, not repositories

| Field | Value |
|-------|-------|
| **Category** | Architecture |
| **Location** | `ShipmentService`, `SgpService`, `CarrierService` use `_connect` + raw SQL; only partial helpers in `core/kp_db_shipments.py` |
| **Impact** | SQL scattered across services; no single persistence boundary; harder to test and evolve schema |
| **Fix** | Extract repository layer; consolidate SQL in `core/kp_db_*` modules |

#### [A6] Dependency injection inconsistent

| Field | Value |
|-------|-------|
| **Category** | Architecture |
| **Location** | `app/dependencies/services.py`: only auth injects repo; logistics factories construct `KpRepository().db_path`; `CommercialWorkflow`/`Offers` self-construct |
| **Impact** | Inconsistent wiring; some services testable, others not; hidden singletons |
| **Fix** | Uniform factory pattern; inject all dependencies via FastAPI `Depends` |

#### [A7] Lazy imports mask cross-service coupling

| Field | Value |
|-------|-------|
| **Category** | Architecture |
| **Location** | `archive_service` runtime-imports `SgpService`, `KpReadinessService`, `kp_db_shipments`; similar in `plan_storage` |
| **Impact** | Circular dependency risk hidden at import time; coupling invisible in static analysis |
| **Fix** | Explicit dependency injection; break cycles via interfaces or event boundaries |

#### [A8] Legacy implicit mutable global state for plate order/optimization

| Field | Value |
|-------|-------|
| **Category** | Architecture |
| **Location** | `core/plate_runtime_state.py`, `core/config_and_data.py`, `core/domain/plate_order.py`; HTTP mitigated by middleware but bot/scripts may bypass |
| **Impact** | Shared mutable state across requests in bot/scripts; race conditions and stale data |
| **Fix** | Pass explicit context; eliminate module-level mutable state |

### Security

#### [S1] In-process rate limits bypassable with multiple workers

| Field | Value |
|-------|-------|
| **Category** | Security |
| **Location** | `login_rate_limit.py:49–86`, `commercial_upload_validation.py:23–45` |
| **Impact** | Brute-force and upload-abuse limits ineffective when running multiple workers or replicas |
| **Fix** | Redis counters or enforce single-worker deployment |

#### [S2] Commercial OCR sends documents to external LLM

| Field | Value |
|-------|-------|
| **Category** | Security |
| **Location** | `core/ocr/recognition.py`, `providers/openai.py`, `commercial.py:69–81`; `OCR_EXTERNAL_ENABLED` |
| **Impact** | Customer commercial documents egress to third-party LLM; data residency and confidentiality risk |
| **Fix** | Keep disabled by default; audit trail; prefer on-prem processing |

#### [S3] Shipment items not validated against shipment order KPs

| Field | Value |
|-------|-------|
| **Category** | Security |
| **Location** | `shipment_service.py` `put_items`/`_prepare_item` (~589–713), `complete` (~727–804) |
| **Impact** | Logistics role can attach any `completed_plate_id` then complete — cross-customer inventory write-off |
| **Fix** | Require `plate.kp_id ∈ shipment order KPs` in `_prepare_item` and `complete` |

### Code Quality

#### [Q1] Duplicated SGP plate-matching SQL across unlink/relink/reserve/free

| Field | Value |
|-------|-------|
| **Category** | Code Quality |
| **Location** | `sgp_service.py` |
| **Impact** | Four near-identical SQL blocks; fix in one path may miss others |
| **Fix** | Extract shared `_match_plates` helper or repository method |

#### [Q2] Duplicated availability validation `_prepare_item` vs `_preflight_availability`

| Field | Value |
|-------|-------|
| **Category** | Code Quality |
| **Location** | `shipment_service.py` |
| **Impact** | Divergent validation logic between code paths; subtle bugs when rules change |
| **Fix** | Single shared availability check function |

#### [Q3] N+1 DB access building archive list rows

| Field | Value |
|-------|-------|
| **Category** | Code Quality |
| **Location** | `archive_service._to_list_item` |
| **Impact** | List endpoint scales O(n) queries; slow archive page under load |
| **Fix** | Batch-fetch related data; join or prefetch in repository |

#### [Q4] SgpService mutation methods near-duplicates (~170 lines each)

| Field | Value |
|-------|-------|
| **Category** | Code Quality |
| **Location** | `sgp_service.py` — `unlink`/`relink`/`reserve_on_conn` |
| **Impact** | ~170 lines duplicated per method; maintenance burden and drift risk |
| **Fix** | Extract common transaction skeleton and plate-matching logic |

#### [Q5] ShipmentItemsSection 644-line stateful component with fragile sync

| Field | Value |
|-------|-------|
| **Category** | Code Quality |
| **Location** | `frontend/src/features/logistics/components/ShipmentItemsSection.tsx` |
| **Impact** | Complex local/server state sync; `eslint-disable exhaustive-deps` masks stale-closure bugs |
| **Fix** | Split into sub-components; derive state; fix dependency array |

#### [Q6] Thin OffersService unit tests

| Field | Value |
|-------|-------|
| **Category** | Code Quality |
| **Location** | `tests/test_offers_service.py` — only 2 PDF/XLSX date tests |
| **Impact** | Core commercial workflow untested at unit level |
| **Fix** | Add tests for move_to_production, validation, error paths |

#### [Q7] Frontend gaps: header save/complete/cancel, CarrierAutocomplete, invalidateRelated tests

| Field | Value |
|-------|-------|
| **Category** | Code Quality |
| **Location** | `ShipmentDrawer`, `CarrierAutocomplete`, `useLogisticsQueries.ts` |
| **Impact** | Critical user flows and cache invalidation untested; regressions likely |
| **Fix** | Add component and hook tests for listed flows |

---

## Medium Priority Issues (plan for next sprint)

### Architecture

- **[A9]** `core/kp_db.py` god facade re-exporting persistence surface — consolidate or document boundary; reduce re-export surface.
- **[A10]** Parallel planning orchestration layers — `app/planning/`, `production_planning_service.py`, `core/production/planning.py` — pick canonical layer; deprecate duplicates.
- **[A11]** Frontend god component — `ShipmentItemsSection.tsx` (~617–644 lines) — split into item row, bulk actions, validation sub-components.
- **[A12]** Over-broad React Query invalidation — `useLogisticsQueries.ts` `invalidateRelated` invalidates logistics+archive+production+sgp — scope invalidation to affected keys.
- **[A13]** Thin re-export services blur app/core boundary — `kp_persistence_service`, `plate_completion_service`, `rest_matching_service` — inline or rename to clarify passthrough.
- **[A14]** SGP warehouse nested under production API — `production.py` vs `/logistics` — align route ownership with domain boundary.
- **[A15]** Archive list enrichment crosses logistics/SGP domains — `archive_service._to_list_item` — move enrichment to dedicated assembler or batch query layer.

### Security

- **[S4]** CSP Report-Only with `unsafe-inline` — `security_headers.py:11–37` — tighten CSP; move to enforce mode when ready.
- **[S5]** No rate limiting on authenticated mutating business APIs — logistics/production/commercial/archive/offers — add per-user/per-endpoint limits.
- **[S6]** Sensitive data at rest unencrypted — SQLite + `drafts_dir` plaintext — encrypt at rest or restrict filesystem access.
- **[S7]** Long session lifetime without idle timeout — `session.py` / settings default 12h — add idle timeout and sliding expiration.
- **[S8]** CSRF cookie readable by JS — `csrf.py` httponly False (needed for double-submit but XSS amplifies) — minimize XSS surface; consider alternative CSRF pattern.
- **[S9]** Dynamic SQL column names without service allowlist — `shipment_service.py:282–286` (Pydantic mitigates at API) — add explicit allowlist in service layer.
- **[S10]** Logistics KP search exposes cross-manager customer data — intentional for logistics? Confirm business rule and document or restrict.
- **[S11]** Shipment creation does not validate KP status — `_assert_kp_exists` only — validate KP is in shippable state.
- **[S12]** No Python dependency scanning in CI — `frontend-audit.yml` npm only — add `pip-audit` or Dependabot for Python.

### Code Quality

- **[Q8]** Magic number 0.005 dimension tolerance not centralized — extract to shared constant in domain/packing config.
- **[Q9]** Hardcoded `'в производстве'` instead of `PlateStatus.IN_PRODUCTION.value` — use enum value.
- **[Q10]** Repeated transaction boilerplate in ShipmentService — extract `_with_transaction` helper.
- **[Q11]** `patch()` accepts `actor` but never persists it — incomplete audit trail — persist actor on mutation or remove parameter.
- **[Q12]** Dual propose implementations (legacy FIFO vs v2 packing) — deprecate legacy path or feature-flag clearly.
- **[Q13]** Duplicated progress enrichment in archive mappers — extract shared enrichment function.
- **[Q14]** OffersService/ArchiveService duplicate `move_to_production` orchestration — single shared use-case service.
- **[Q15]** OffersService untyped `ValueError` string codes — use typed exception hierarchy or error enum.
- **[Q16]** XLSX export embedded in ShipmentService — move to `ShipmentExportService` (see A1).
- **[Q17]** `reserve_on_conn` lacks direct unit tests — add isolated tests with mocked connection.
- **[Q18]** Duplicated `generate_pdf`/`generate_xlsx` in OffersService — extract shared document generation helper.

---

## Low Priority / Suggestions

### Architecture

- **[A16]** Dual PlateOrder model hierarchy — `core/domain` vs `app/domain` — consolidate or document mapping.
- **[A17]** `plan_manager.py` manipulates `sys.path` at import time — fix import structure; remove runtime path hack.
- **[A18]** Positive: `core/shipment_packing/` is well-bounded — maintain as template for future domain modules.

### Security

- **[S13]** CORS `allow_headers=["*"]` — `main.py:61–67` — restrict to required headers.
- **[S14]** `session_version` exposed in auth responses — evaluate necessity; remove if not client-used.
- **[S15]** Missing Permissions-Policy header — add restrictive Permissions-Policy.
- **[S16]** Production role can list all managers — confirm business rule; restrict if not required.
- **[S17]** Legacy web login blocked without CSRF (dead path) — remove dead code path.

### Code Quality

- **[Q19]** Module-level mutable counter for React draft keys — `draftItems.ts` — use `useRef` or UUID per instance.
- **[Q20]** Duplicated dimension formatting logistics vs SgpWarehouseView — extract shared formatter.
- **[Q21]** `type: ignore[arg-type]` on product_type mapping — `shipment_service.py:1068` — fix typing at source.
- **[Q22]** Empty untracked scratch file `_tmp_old.py` — delete from repo.

---

## Priority Matrix

| ID | Issue | Severity | Effort | Priority |
|----|-------|----------|--------|----------|
| A1 | ShipmentService god module | Critical | High | P0 |
| A2 | Single-instance storage not enforced | Critical | High | P0 |
| S3 | Shipment items not validated against order KPs | High | Medium | P0 |
| S1 | In-process rate limits bypassable | High | Medium | P1 |
| A3 | CommercialWorkflowService god module | High | High | P1 |
| A4 | Services return Pydantic HTTP schemas | High | Medium | P1 |
| A5 | SQL in services, not repositories | High | High | P1 |
| Q1 | Duplicated SGP plate-matching SQL | High | Medium | P1 |
| Q2 | Duplicated availability validation | High | Low | P1 |
| Q3 | N+1 DB access in archive list | High | Medium | P1 |
| S2 | OCR sends documents to external LLM | High | Low | P2 |
| A6 | Dependency injection inconsistent | High | Medium | P2 |
| A7 | Lazy imports mask coupling | High | Medium | P2 |
| A8 | Legacy mutable global plate state | High | Medium | P2 |
| Q4 | SgpService mutation near-duplicates | High | Medium | P2 |
| Q5 | ShipmentItemsSection god component | High | Medium | P2 |
| Q6 | Thin OffersService unit tests | High | Medium | P2 |
| Q7 | Frontend test gaps | High | Medium | P2 |
| S4 | CSP Report-Only with unsafe-inline | Medium | Medium | P3 |
| S5 | No rate limit on mutating APIs | Medium | Medium | P3 |
| S11 | Shipment creation KP status not validated | Medium | Low | P3 |
| A11 | ShipmentItemsSection frontend god component | Medium | Medium | P3 |
| A12 | Over-broad React Query invalidation | Medium | Low | P3 |
| Q14 | Duplicate move_to_production orchestration | Medium | Medium | P3 |
| A9 | kp_db.py god facade | Medium | Medium | P4 |
| A10 | Parallel planning layers | Medium | High | P4 |
| A13 | Thin re-export services | Medium | Low | P4 |
| A14 | SGP under production API | Medium | Medium | P4 |
| A15 | Archive enrichment crosses domains | Medium | Medium | P4 |
| S6 | Data at rest unencrypted | Medium | High | P4 |
| S7 | Long session without idle timeout | Medium | Low | P4 |
| S8 | CSRF cookie readable by JS | Medium | Low | P4 |
| S9 | Dynamic SQL column names | Medium | Low | P4 |
| S10 | Logistics KP search cross-manager | Medium | Low | P4 |
| S12 | No Python dependency scanning in CI | Medium | Low | P4 |
| Q8–Q18 | Remaining code quality medium items | Medium | Low–Med | P4 |
| A16–A18 | Architecture low / positive | Low | Low | P5 |
| S13–S17 | Security low items | Low | Low | P5 |
| Q19–Q22 | Code quality low items | Low | Low | P5 |

---

## Next Steps

1. **Immediate** (before next commit): critical fixes — **A1** (begin decomposition), **A2** (startup guard for single-instance)
2. **This sprint**: high fixes especially **S3** (inventory integrity), **S1** (rate limits), **A3–A5** (service/repository boundaries)
3. **Next sprint**: medium — security hardening (S4–S12), frontend decomposition (A11, Q5), test coverage (Q6, Q7)
4. **Backlog**: low — cleanup, headers, dead code, formatting dedup

Use `/refactor [file]` for structural issues.
Use `/implement [fix]` for feature-level security fixes.

---

## Related Documentation

- Previous audit: [2026-08-02-full-project-audit.md](./2026-08-02-full-project-audit.md)
- Audit comparison: [2026-08-02-audit-comparison.md](./2026-08-02-audit-comparison.md)
- Deployment ADR: [deployment-single-instance.md](../architecture/deployment-single-instance.md)
- Stabilization specs: [stabilizaciya-p0-p1-audit-2026-08-02.md](../../specs/stabilizaciya-p0-p1-audit-2026-08-02.md)
