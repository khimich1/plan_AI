# Commercial Offer Module Audit

**Date:** 2026-05-05  
**Scope:** Commercial offer slice — `frontend/src/app/layout/AppHeader.tsx`, `frontend/src/pages/commercial-offer-create/`, `frontend/src/features/commercial-offer/`, `app/api/v1/endpoints/commercial.py`, `app/services/commercial_service.py`, `app/services/commercial_workflow_service.py`, `app/schemas/commercial.py`  
**Audited by:** senior-reviewer + security-auditor + reviewer

---

## Executive Summary

**Health Score: 4.0 / 10** (Critical assessment required)

The commercial offer module exhibits significant architectural, security, and code quality issues that must be addressed before production. Primary concerns include orchestration contract drift between client and server, identity and access control vulnerabilities (IDOR, path traversal), unbounded upload exposure, and monolithic design patterns.

### Severity Summary

| Category | Critical | High | Medium | Low | Total |
|----------|----------|------|--------|-----|-------|
| Architecture | 1 | 4 | 8 | 3 | 16 |
| Security | — | 3 | 2 | 4 | 9 |
| Code Quality | — | 3 | 9 | 7 | 19 |
| **Total** | **1** | **10** | **19** | **13** | **43** |

**Recommendation:**  
Address critical A1 immediately; remediate all high-priority items before sprint merge. Medium-priority refactoring (architecture, DI injection, type safety) should be scheduled within 2 weeks. Consider architectural review of wizard state management before feature expansion.

---

## Critical Issues

### A1 – Orchestration Contract Drift (Critical)

**Category:** Architecture  
**Location:** `frontend/src/pages/commercial-offer-create/`, `frontend/src/features/commercial-offer/`, `app/services/commercial_workflow_service.py`  
**Severity:** Critical  

**Issue:**  
Two competing orchestrators manage wizard state:
- **Client side:** UI steps indexed by `WizardStepId` enum; step flow driven by React state and query mutations.
- **Server side:** `CommercialWorkflowService.calculate_draft()` and `current_step` metadata manage business logic progression.

**Impact:**  
- UI may skip POST /calculate when navigating steps; server metadata (`current_step`, `generated_file_step`) becomes stale.
- Draft state diverges from server intent; financial calculations may not execute when expected.
- Client-side financial totals in `draft.totals` may conflict with server-recalculated values.
- Contract not formally defined; field names (`calculate_draft` on server, no explicit state machine on client) are implicit.

**Fix Suggestion:**
1. Define a formal state machine contract (request/response schema) specifying:
   - Valid state transitions (`DRAFT_CREATED` → `PRICE_ESTIMATED` → `WIDE_PLATES_SET` → ...)
   - Server-side `current_step` as source of truth; POST /calculate always required before client transitions.
2. Add Pydantic schema `WizardStateResponse` with `current_step`, `can_proceed_to: list[WizardStepId]`, `next_required_action`.
3. Client mutation: Before each UI step transition, call `POST /calculate` with current form data; use response `can_proceed_to` to gate UI buttons.
4. Synchronize `WizardStepId` enum between FE and BE (`schemas/commercial.py`).

---

## High Priority Issues

### A2 – CommercialWorkflowService God-Module

**Category:** Architecture  
**Location:** `app/services/commercial_workflow_service.py`  
**Severity:** High  

**Issue:**  
`CommercialWorkflowService` owns calculation, file generation, draft state resolution, and wide-plate orchestration—violating single responsibility principle.

**Impact:**  
- Module is difficult to test in isolation; dependencies on `FileGenerationService`, `ProductionPlanService`, price providers span multiple concerns.
- New business logic (e.g., alternative pricing strategies, new draft stages) adds to the same god-module.
- Circular dependency risk with `commercial.py` endpoint that calls `workflow._resolve_generated_file()` (private).

**Fix Suggestion:**
1. Extract sub-services:
   - `DraftCalculationService`: compute totals, validate order data.
   - `WizardOrchestrationService`: manage state transitions, determine `current_step`, enforce rules.
   - `OfferGenerationService`: coordinate file generation and template rendering.
2. Use factory pattern in `commercial.py` to compose services via DI.
3. Define clear interfaces (protocols) for each sub-service.

---

### A3 – Duplicate Money Logic (Client vs Server)

**Category:** Architecture  
**Location:** `frontend/src/features/commercial-offer/`, `app/services/commercial_workflow_service.py`  
**Severity:** High  

**Issue:**  
Client computes totals (subtotal, tax, total) independently; server also calculates via `calculate_draft()`. `draft.totals` is populated by POST /calculate, but client may display different values before sync.

**Impact:**  
- User sees inconsistent pricing if server recalculation differs from client estimate.
- Maintenance burden: money logic changes require updates in both FE and BE.
- Risk of desync on slow networks or if client caching is stale.

**Fix Suggestion:**
1. Server is source of truth: POST /calculate returns `CalculationResult` with all totals.
2. Client mutation uses `onSuccess` to update Redux/Zustand with server-calculated `draft.totals`.
3. Remove client-side total computation; display server `draft.totals` only.
4. Validate on submit that client form data matches server totals (ETag or checksum).

---

### A4 – AppHeader Tightly Coupled to Wizard Store

**Category:** Architecture  
**Location:** `frontend/src/app/layout/AppHeader.tsx`  
**Severity:** High  

**Issue:**  
`AppHeader` directly selects and mutates wizard store state (`setStep`, `resetWizard`, `close`). State tied to a single use-case; changing wizard UX requires AppHeader refactor.

**Impact:**  
- Reusability limited: AppHeader cannot be used in other flows without conditional logic.
- Store leak: UI chrome (header) has coupling to business domain (wizard).
- Risk of state corruption if close/reset logic differs in other pages.

**Fix Suggestion:**
1. Inject step-agnostic callbacks: `onStepChange(step)`, `onReset()`, `onClose()`.
2. Wrap AppHeader in container component that supplies these callbacks from context.
3. Move wizard-specific logic (e.g., reset confirmation) to page level.
4. Use composition: `<AppLayout onClose={handleClose} />` instead of direct store binding.

---

### A5 – commercial.py Calls Private Workflow Method

**Category:** Architecture  
**Location:** `app/api/v1/endpoints/commercial.py`, `app/services/commercial_workflow_service.py`  
**Severity:** High  

**Issue:**  
`commercial.py` calls `workflow._resolve_generated_file()` (underscore-prefixed private method) directly, instead of using public interface.

**Impact:**  
- Violates encapsulation; endpoint coupled to private implementation details.
- Refactoring workflow internals risks breaking endpoint without clear dependency graph.
- No contract specification for what `_resolve_generated_file()` is supposed to do.

**Fix Suggestion:**
1. Create public method `CommercialWorkflowService.get_or_generate_file(draft_id: UUID) -> FileModel`.
2. Endpoint calls `workflow.get_or_generate_file()`.
3. Mark internal helper `_resolve_generated_file()` as private in docstring or move to module level.

---

### S-H01 – Path Traversal in DraftStore (High)

**Category:** Security  
**Location:** `app/services/commercial_service.py` (DraftStore path join)  
**Severity:** High  

**Issue:**  
If draft files are stored in `DRAFT_DIR / {user_id} / {draft_id} / {filename}`, an attacker could supply `filename` with `../` sequences to escape the draft directory.

**Impact:**  
- Arbitrary file read/write within user's directory tree.
- Potential disclosure of other user's draft files if directory traversal crosses ownership boundary.

**Fix Suggestion:**
1. Sanitize filename: reject any path containing `..`, `/`, `\`, or absolute paths.
   ```python
   import pathlib
   def safe_join(base: Path, *parts: str) -> Path:
       result = base
       for part in parts:
           if ".." in part or part.startswith(("/", "\\")):
               raise ValueError(f"Invalid path component: {part}")
           result = result / part
       return result.resolve().relative_to(base.resolve())  # Verify within base
   ```
2. Or use a deterministic filename: `{draft_id}_{field_name}.{ext}` with no user input in path.

---

### S-H02 – IDOR: No Per-User Draft Ownership Check (High)

**Category:** Security  
**Location:** `app/api/v1/endpoints/commercial.py`  
**Severity:** High  

**Issue:**  
Endpoints like `GET /draft/{draft_id}`, `POST /draft/{draft_id}/calculate` do not verify the draft belongs to the requesting user. Role-based file access (`files_folder_access_by_role`) controls access but does not tie drafts to users.

**Impact:**  
- User A can read/modify drafts created by User B by guessing or enumerating `draft_id`.
- Financial offers leaked, modified, or exfiltrated.

**Fix Suggestion:**
1. Add `user_id` to `Draft` schema; set during creation.
2. Middleware or dependency: `verify_draft_ownership(draft_id, current_user)`.
3. All draft endpoints check ownership before proceeding:
   ```python
   @router.get("/draft/{draft_id}")
   async def get_draft(draft_id: UUID, current_user: User = Depends(get_current_user)):
       draft = await repository.get_draft(draft_id)
       if draft.user_id != current_user.id:
           raise HTTPException(403, "Access denied")
       return draft
   ```

---

### S-H03 – Unbounded File Uploads, No Rate Limits (High)

**Category:** Security  
**Location:** `app/api/v1/endpoints/commercial.py` (file upload handlers)  
**Severity:** High  

**Issue:**  
File upload endpoints (e.g., OCR image upload, invoice scan) do not enforce:
- Upload size limits.
- Rate limiting per user.
- OCR cost/quota controls (abuse potential).

**Impact:**  
- Disk exhaustion: attacker uploads 10 GB files, service runs out of space.
- OCR cost abuse: attacker triggers 1000s of OCR requests, inflating API costs.
- DoS vector.

**Fix Suggestion:**
1. Add upload size limit (e.g., 50 MB per file):
   ```python
   @router.post("/draft/{draft_id}/upload-invoice")
   async def upload_invoice(
       draft_id: UUID,
       file: UploadFile = File(...),
       current_user: User = Depends(get_current_user)
   ):
       MAX_SIZE = 50 * 1024 * 1024  # 50 MB
       content = await file.read()
       if len(content) > MAX_SIZE:
           raise HTTPException(413, "File too large")
   ```
2. Implement rate limiting (e.g., using `slowapi` or Redis):
   - Max 10 OCR uploads per user per hour.
   - Max 500 MB per user per day.
3. Track OCR costs; alert or reject if user exceeds budget.

---

### Q1 – React Query v5 useQuery onSuccess Invalid (High)

**Category:** Code Quality  
**Location:** `frontend/src/features/commercial-offer/`  
**Severity:** High  

**Issue:**  
`useQuery` with `onSuccess` callback is deprecated in React Query v5. The callback may not fire reliably, or its closure captures stale state.

**Impact:**  
- Draft state may not update after calculation; UI shows stale data.
- User perceives failed operations or inconsistent state.

**Fix Suggestion:**
1. Replace `onSuccess` with `useEffect` on `data` change:
   ```typescript
   const { data: result } = useQuery({
     queryKey: ["draft", draftId],
     queryFn: async () => await api.calculateDraft(draftId),
   });
   
   useEffect(() => {
     if (result) {
       dispatch(setDraft(result));
     }
   }, [result, dispatch]);
   ```
2. Or use `useMutation` with `onSuccess` for mutations (more appropriate):
   ```typescript
   const calculateMutation = useMutation({
     mutationFn: async () => await api.calculateDraft(draftId),
     onSuccess: (result) => {
       dispatch(setDraft(result));
     },
   });
   ```

---

### Q2 – calculateMutation Exported but Unused; No POST /calculate from Wizard (High)

**Category:** Code Quality  
**Location:** `frontend/src/features/commercial-offer/`  
**Severity:** High  

**Issue:**  
`calculateMutation` is exported from a hook/store module but never called from the wizard UI flow. Wizard steps proceed without recalculating server state; financial calculations may be skipped.

**Impact:**  
- Totals and business rules not enforced server-side; client data dominates.
- Violates A1 contract (server-side orchestration).

**Fix Suggestion:**
1. Call mutation when user moves to next step or saves form:
   ```typescript
   const handleNextStep = async () => {
     try {
       const result = await calculateMutation.mutateAsync();
       if (result.current_step !== nextStep) {
         // Server says you can't proceed yet
         showError(`Cannot proceed to ${nextStep}: ${result.reason}`);
         return;
       }
       setStep(nextStep);
     } catch (e) {
       showError("Calculation failed");
     }
   };
   ```
2. Or integrate into wizard state machine (see A1 fix).

---

### Q3 – Mutation onSuccess Closure Stale State (High)

**Category:** Code Quality  
**Location:** `frontend/src/features/commercial-offer/`  
**Severity:** High  

**Issue:**  
If `useMutation` is created in a hook without re-running on dependency change, `onSuccess` callback captures stale variables (e.g., form data, user context).

**Impact:**  
- Mutation succeeds but updates use outdated state; form UI shows incorrect values.
- Race condition if user changes form while mutation is pending.

**Fix Suggestion:**
1. Move `useMutation` to mutation-only hooks; avoid capturing form state in closure.
2. Use `onSuccess` only to invalidate queries:
   ```typescript
   const calculateMutation = useMutation({
     mutationFn: async (data) => await api.calculateDraft(draftId, data),
     onSuccess: () => {
       queryClient.invalidateQueries({ queryKey: ["draft", draftId] });
     },
   });
   ```
3. Handle state updates in `useEffect` (Q1 pattern) after query re-fetch.

---

## Medium Priority Issues

### A6 – /parse and /generate-preview Return Untyped dict (Medium)

**Category:** Architecture  
**Location:** `app/api/v1/endpoints/commercial.py`  
**Severity:** Medium  

**Issue:**  
Endpoints `POST /parse-invoice` and `POST /generate-preview` return `dict` instead of typed Pydantic schemas.

**Impact:**  
- No schema documentation in OpenAPI; FE must guess response structure.
- Type safety lost; refactoring endpoint internals can silently break client.

**Fix Suggestion:**
```python
class InvoiceParseResult(BaseModel):
    extracted_items: list[OrderItem]
    total_amount: float
    vendor: str
    confidence: float

@router.post("/parse-invoice", response_model=InvoiceParseResult)
async def parse_invoice(file: UploadFile, ...):
    # ...
    return InvoiceParseResult(...)
```

---

### A7 – No FastAPI Dependency Injection; Service Per Request (Medium)

**Category:** Architecture  
**Location:** `app/api/v1/endpoints/commercial.py`  
**Severity:** Medium  

**Issue:**  
Services (`CommercialService`, `CommercialWorkflowService`) instantiated inline in endpoints or passed as function arguments instead of using FastAPI `Depends()`.

**Impact:**  
- Harder to test: mock services cannot be injected into route handlers.
- Boilerplate: each endpoint repeats service instantiation.
- Coupling: endpoint imports and instantiates specific service classes.

**Fix Suggestion:**
```python
from fastapi import Depends

async def get_commercial_service() -> CommercialService:
    return CommercialService(db=get_db())

@router.post("/draft/{draft_id}/calculate")
async def calculate_draft(
    draft_id: UUID,
    commercial_service: CommercialService = Depends(get_commercial_service),
):
    return await commercial_service.calculate(draft_id)
```

---

### A8 – CommercialService Uses Private FileGenerationService._legacy_order_context (Medium)

**Category:** Architecture  
**Location:** `app/services/commercial_service.py`  
**Severity:** Medium  

**Issue:**  
`CommercialService` accesses private method `FileGenerationService._legacy_order_context()` instead of calling public API.

**Impact:**  
- Encapsulation violation; refactoring `FileGenerationService` risks breaking `CommercialService`.
- No version contract; unclear if `_legacy_order_context` is temporary or permanent.

**Fix Suggestion:**
1. Create public method `FileGenerationService.get_order_context(draft: Draft) -> OrderContext`.
2. Call `generation_service.get_order_context(draft)` instead of `._legacy_order_context()`.

---

### A9 – generateFiles onSuccess Uses Stale Closure State (Medium)

**Category:** Architecture  
**Location:** `frontend/src/pages/commercial-offer-create/`  
**Severity:** Medium  

**Issue:**  
If `generateFiles()` mutation has `onSuccess` callback, the callback may use stale `draft` or `formData` variables.

**Impact:**  
- File generation completes but uses old form data; generated file is incorrect.
- User downloads outdated offer.

**Fix Suggestion:**
- See Q1 and Q3 patterns; use `useEffect` on query data change instead of `onSuccess` closure.

---

### A10 – FE order_data Typed as Record<string, unknown> (Medium)

**Category:** Architecture  
**Location:** `frontend/src/features/commercial-offer/`  
**Severity:** Medium  

**Issue:**  
Frontend `order_data` or form state typed as `Record<string, unknown>` instead of a specific schema matching backend `DraftOrder` or `OrderData`.

**Impact:**  
- No type checking; typos in field names not caught at compile time.
- Mismatch between client and server field names (e.g., `plateLines` vs `plate_lines`).

**Fix Suggestion:**
1. Define TypeScript interface:
   ```typescript
   interface DraftOrderData {
     customer_name: string;
     order_date: string;
     plate_lines: PlateLineData[];
     // ... fields matching backend schema
   }
   ```
2. Use in form state: `const [orderData, setOrderData] = useState<DraftOrderData>({...})`.

---

### A11 – POST /from-form Endpoint Exists but Not in commercialOfferApi (Medium)

**Category:** Architecture  
**Location:** `app/api/v1/endpoints/commercial.py`, `frontend/src/features/commercial-offer/`  
**Severity:** Medium  

**Issue:**  
Backend endpoint `POST /from-form` exists but FE API client (`commercialOfferApi`) does not have a method calling it. UI may be using a different endpoint or bypassing the API layer.

**Impact:**  
- API discrepancy; hard to track what client actually calls.
- Maintenance risk: if `/from-form` logic changes, FE may not notice.

**Fix Suggestion:**
1. Add FE method:
   ```typescript
   export async function createDraftFromForm(data: FormData) {
     return api.post("/draft/from-form", data);
   }
   ```
2. Use consistently in wizard.
3. Document in OpenAPI schema which endpoint is primary for form-based draft creation.

---

### A12 – Step Naming Mismatch FE vs BE (Medium)

**Category:** Architecture  
**Location:** `frontend/src/features/commercial-offer/`, `app/services/commercial_workflow_service.py`  
**Severity:** Medium  

**Issue:**  
Frontend `WizardStepId` enum (e.g., `PRODUCTS`, `WIDE_PLATES`, `REVIEW`) does not match backend `current_step` values (e.g., `draft_created`, `price_estimated`, `wide_plates_set`).

**Impact:**  
- See A1: orchestration contract drift; client cannot interpret server `current_step`.

**Fix Suggestion:**
1. Define shared enum in backend: `class WizardStep(str, Enum)`.
2. Export from `schemas/commercial.py` for client import or sync via OpenAPI generator.
3. Backend returns `current_step: WizardStep` (not free-form string).

---

### A13 – hydrate-draft Keeps Old currentStep vs Server Metadata (Medium)

**Category:** Architecture  
**Location:** `frontend/src/features/commercial-offer/`, `app/services/commercial_workflow_service.py`  
**Severity:** Medium  

**Issue:**  
If draft is reloaded (hydration), client re-initializes `wizard.currentStep` from saved Redux state instead of server `draft.current_step` metadata.

**Impact:**  
- Wizard UI shows old step; server is ahead or behind. User confusion.

**Fix Suggestion:**
1. Hydration: `setStep(draft.current_step)` (not local store state).
2. On load: fetch latest draft via `GET /draft/{id}` to sync server state.
3. Client step is not authoritative; always defer to server on re-entry.

---

### S-M01 – Weak Upload Validation (Medium)

**Category:** Security  
**Location:** `app/api/v1/endpoints/commercial.py`  
**Severity:** Medium  

**Issue:**  
File upload endpoints validate only `content-type` header; attacker can set `Content-Type: image/jpeg` and upload executable or malicious content.

**Impact:**  
- Malware uploaded disguised as image.
- If uploaded file is later processed (e.g., OCR), vulnerability chains.

**Fix Suggestion:**
1. Validate file signature (magic bytes):
   ```python
   ALLOWED_SIGNATURES = {
       b'\xff\xd8\xff': 'jpeg',  # JPEG
       b'\x89\x50\x4e\x47': 'png',  # PNG
       b'%PDF': 'pdf',  # PDF
   }
   
   content = await file.read()
   file_type = None
   for sig, ftype in ALLOWED_SIGNATURES.items():
       if content.startswith(sig):
           file_type = ftype
           break
   if not file_type:
       raise HTTPException(400, "Invalid file type")
   ```
2. Scan with antivirus (e.g., ClamAV) before processing.
3. Store in restricted directory; serve with `Content-Disposition: attachment` to prevent browser execution.

---

### S-M02 – Error Detail Leakage (Medium)

**Category:** Security  
**Location:** `app/api/v1/endpoints/commercial.py`  
**Severity:** Medium  

**Issue:**  
Exception handlers return `detail=str(exc)`, exposing internal error messages (database URLs, paths, stack traces) to client.

**Impact:**  
- Information disclosure: attacker learns infrastructure details, file paths, library versions.
- Aids reconnaissance for further attacks.

**Fix Suggestion:**
```python
@router.post("/draft/{draft_id}/calculate")
async def calculate_draft(...):
    try:
        return await commercial_service.calculate(draft_id, data)
    except ValueError as e:
        logger.error(f"Validation failed: {e}")
        raise HTTPException(400, "Invalid input")
    except Exception as e:
        logger.exception(f"Unexpected error: {e}")
        raise HTTPException(500, "Internal error")  # Generic message
```

---

### Q4 – Monolithic Wizard Component (Medium)

**Category:** Code Quality  
**Location:** `frontend/src/pages/commercial-offer-create/`  
**Severity:** Medium  

**Issue:**  
Wizard page combines step rendering, form validation, API calls, and state management in one file/component.

**Impact:**  
- Hard to test individual steps.
- Reuse limited; wizard step UX cannot be shared with other flows.
- Maintenance burden; changes to one step risk breaking others.

**Fix Suggestion:**
1. Extract step components: `WizardProductsStep`, `WizardPlatesStep`, etc.
2. Each step receives: `formData`, `onSubmit`, `onError` props.
3. Wizard page orchestrates: renders active step, handles navigation.
4. Shared logic (API calls, validation) in custom hooks.

---

### Q5–Q12 – Additional Code Quality Issues (Medium)

**Category:** Code Quality  
**Severity:** Medium  

- **Q5**: Repeated `try/catch` pattern in `CommercialWorkflowService` methods; extract common error handler.
- **Q6**: Duplicate `resetWizard` logic in AppHeader and page component; centralize in store.
- **Q7**: `_build_order_data()` method >100 lines; break into smaller helpers (name, items, totals, metadata).
- **Q8**: `resolve_wide_plates()` method >80 lines; extract plate validation and normalization.
- **Q9**: `preview_response` and `price_rows` typed as `dict` or `Any`; use `PreviewResult` and `PriceRow` schemas.
- **Q10**: `Draft.current_step` default value mismatch: FE expects `DRAFT_CREATED`, BE defaults to `None`.
- **Q11**: DRY violation: `/calculate`, `/preview`, `/generate` all call similar setup/validation boilerplate; extract.
- **Q12**: `CalculationResultStep` has dual `totals` (client + server); unify to server `draft.totals`.

---

## Low Priority Issues

### A14 – calculateMutation Exported but Unused (Low)

**Category:** Architecture  
**Location:** `frontend/src/features/commercial-offer/`  
**Severity:** Low  

**Issue:**  
`calculateMutation` hook is exported from module but never imported elsewhere (dead export).

**Impact:**  
- Code clutter; misleads future developers.

**Fix Suggestion:**
- Remove export; or document why it exists (e.g., for future use in another component).

---

### A15 – Path Constants Duplicate Router Definition (Low)

**Category:** Architecture  
**Location:** `frontend/src/features/commercial-offer/`, `app/api/v1/endpoints/commercial.py`  
**Severity:** Low  

**Issue:**  
API endpoint paths (e.g., `/api/v1/draft/{id}/calculate`) are hardcoded in multiple places instead of a single constant file.

**Impact:**  
- If endpoint path changes, must update multiple files.

**Fix Suggestion:**
```typescript
// src/features/commercial-offer/api.constants.ts
export const COMMERCIAL_API = {
  DRAFT_CALCULATE: "/api/v1/draft/{id}/calculate",
  DRAFT_GENERATE: "/api/v1/draft/{id}/generate",
  // ...
};
```

---

### A16 – WidePlateReviewStep useEffect Side Effects (Low)

**Category:** Architecture  
**Location:** `frontend/src/features/commercial-offer/WidePlateReviewStep.tsx`  
**Severity:** Low  

**Issue:**  
`useEffect` hook has multiple side effects (fetch, subscribe, cleanup) combined; hard to trace dependencies.

**Impact:**  
- Potential memory leaks if cleanup is missed.
- Stale closures if dependencies are wrong.

**Fix Suggestion:**
```typescript
// Separate concerns
useEffect(() => {
  // Fetch wide plates
}, [draftId]);

useEffect(() => {
  // Subscribe to changes
  const unsubscribe = store.subscribe(handleStateChange);
  return unsubscribe;
}, [store]);
```

---

### S-L01–S-L04 – Low-Severity Security Issues (Low)

**Category:** Security  
**Severity:** Low  

- **S-L01**: Client-side totals not validated against server; users could theoretically manipulate form before POST (though totals recalculated server-side). Impact low if server is authoritative.
- **S-L02**: Absolute URLs in error messages or logs; could disclose environment if exposed. Use relative URLs or redact environment-specific parts.
- **S-L03**: SessionStorage used for wizard state; XSS vulnerability if sensitive data stored. Use `httpOnly` cookies for sensitive state or auth tokens.
- **S-L04**: Download file path constructed from user input; validate that served file is within expected directory (already covered in S-H01, but re-iterate for downloads).

---

### Q13–Q18 – Low-Severity Code Quality Issues (Low)

**Category:** Code Quality  
**Severity:** Low  

- **Q13**: Dead code: `lastPlateMode` variable in wizard state; remove.
- **Q14**: Alias without use: `_normalize_wide_plate_lines = normalize_wide_plate_lines`; clean up.
- **Q15**: `Record<string, any>` types in place of specific interfaces; prefer strict typing.
- **Q16**: Test coverage gap: commercial offer endpoints <60% covered; add integration tests for calculate, generate workflows.
- **Q17**: VAT hardcoded as 18%; move to config or settings.
- **Q18**: `resolveWidePlatesMutation` typed as `useMutation<Result, Error, WidePlateData>`, but called with `lineId` instead of full plate data; type vs. usage mismatch.

---

## Priority Matrix

| ID | Issue | Severity | Category | Estimated Effort | Priority | Next Action |
|---|---|---|---|---|---|---|
| A1 | Orchestration Contract Drift | Critical | Arch | 21 days | **P0** | Immediate: define state machine contract |
| S-H01 | Path Traversal in DraftStore | High | Security | 2 days | **P0** | Immediate: sanitize file paths |
| S-H02 | IDOR: No Draft Ownership | High | Security | 3 days | **P0** | Immediate: add user_id check |
| S-H03 | Unbounded Uploads | High | Security | 3 days | **P0** | Immediate: add size/rate limits |
| A2 | CommercialWorkflowService God-Module | High | Arch | 8 days | **P1** | Sprint 1: extract sub-services |
| A3 | Duplicate Money Logic | High | Arch | 5 days | **P1** | Sprint 1: unify to server |
| A4 | AppHeader Coupling | High | Arch | 2 days | **P1** | Sprint 1: inject callbacks |
| A5 | Private Method Call | High | Arch | 1 day | **P1** | Sprint 1: expose public API |
| Q1 | React Query onSuccess | High | QA | 1 day | **P1** | Sprint 1: migrate to useEffect |
| Q2 | No POST /calculate from Wizard | High | QA | 2 days | **P1** | Sprint 1: integrate mutation |
| Q3 | onSuccess Stale Closure | High | QA | 2 days | **P1** | Sprint 1: refactor hooks |
| A6 | Untyped dict Response | Medium | Arch | 2 days | **P2** | Sprint 2: add Pydantic schemas |
| A7 | No DI in Endpoints | Medium | Arch | 3 days | **P2** | Sprint 2: migrate to Depends() |
| A8 | Private Method Usage | Medium | Arch | 1 day | **P2** | Sprint 2: expose public API |
| A9 | generateFiles Stale State | Medium | Arch | 1 day | **P2** | Sprint 2: use useEffect pattern |
| A10 | FE Record Unknown Type | Medium | Arch | 2 days | **P2** | Sprint 2: add TypeScript interface |
| A11 | Missing API Method | Medium | Arch | 1 day | **P2** | Sprint 2: add to client |
| A12 | Step Naming Mismatch | Medium | Arch | 2 days | **P2** | Sprint 2: align enums |
| A13 | hydrate-draft Old Step | Medium | Arch | 1 day | **P2** | Sprint 2: fetch server state |
| S-M01 | Weak Upload Validation | Medium | Security | 2 days | **P2** | Sprint 2: validate signatures |
| S-M02 | Error Leakage | Medium | Security | 1 day | **P2** | Sprint 2: generic errors |
| Q4 | Monolithic Wizard | Medium | QA | 5 days | **P2** | Sprint 2: extract step components |
| Q5–Q11 | Code Patterns & Duplication | Medium | QA | 8 days | **P2** | Sprint 2–3: refactor services & components |
| A14 | Dead Export | Low | Arch | <1 day | **P3** | Backlog: remove |
| A15 | Path Duplicate Constants | Low | Arch | 1 day | **P3** | Backlog: centralize |
| A16 | useEffect Side Effects | Low | Arch | 1 day | **P3** | Backlog: separate concerns |
| S-L01–L04 | Security Warnings | Low | Security | 1 day | **P3** | Backlog: document or mitigate |
| Q13–Q18 | Code Cleanup | Low | QA | 2 days | **P3** | Backlog: refactor |

---

## Next Steps

### Immediate (This Week) – Critical & High-Security

1. **A1 – Orchestration Contract Drift**
   - Define `WizardStateResponse` schema (backend).
   - Sync `WizardStepId` enum between FE and BE.
   - Client: gate step transitions on server `can_proceed_to` response.
   
2. **S-H01 – Path Traversal**
   - Implement `safe_join()` path sanitization.
   - Add test: attempt `../` in filename; verify rejection.
   
3. **S-H02 – IDOR**
   - Add `user_id` to `Draft` model.
   - Add dependency: `verify_draft_ownership(draft_id, current_user)`.
   - Apply to all draft endpoints.
   
4. **S-H03 – Unbounded Uploads**
   - Add 50 MB file size limit.
   - Implement rate limit: 10 OCR per user per hour, 500 MB per user per day.

### Sprint 1 (Next 2 Weeks) – Architectural High Priority

5. **A2 – God-Module Refactoring**
   - Extract `DraftCalculationService`, `WizardOrchestrationService`, `OfferGenerationService`.
   - Update `commercial.py` to use DI.
   
6. **A3 – Duplicate Money Logic**
   - Remove client-side total computation; display server `draft.totals`.
   - Add sync check in useEffect.
   
7. **A4 – AppHeader Decoupling**
   - Inject callbacks: `onStepChange`, `onReset`, `onClose`.
   - Wrap in container component.
   
8. **Q1, Q2, Q3 – React Query & Mutation Fixes**
   - Migrate `onSuccess` to `useEffect` pattern.
   - Integrate `calculateMutation` into wizard flow.
   - Ensure server state is authoritative.

### Sprint 2 (Weeks 3–4) – Medium-Priority Refactoring

9. **A6, A7, A8, A10–A13 – Schemas, DI, Types, Step Alignment**
   - Add Pydantic response schemas.
   - Full Depends() migration.
   - TypeScript interfaces for order data.
   - Align step enums.
   
10. **Security Q2 Items**
    - Validate file signatures (S-M01).
    - Generic error responses (S-M02).
    
11. **Q4–Q11 – Code Organization**
    - Extract wizard step components.
    - Consolidate try/catch patterns.
    - Refactor long methods (_build_order_data, resolve_wide_plates).

### Backlog – Low Priority

12. Remove dead code (A14, Q13–Q14).
13. Centralize path constants (A15).
14. Refactor useEffect side effects (A16).
15. Address security notes (S-L01–L04).

---

## Recommendation Summary

**Overall Health: 4.0/10 – Action Required**

The commercial offer module has critical flaws in orchestration, security, and architecture that must be remedied before production use. The path traversal and IDOR vulnerabilities are immediate security risks; the orchestration contract drift undermines business logic integrity.

**Recommended Path:**
1. **This week:** Close security holes (path traversal, IDOR, upload limits).
2. **Next 2 weeks:** Fix orchestration, DI, and state management (A1–A5, Q1–Q3).
3. **Weeks 3–4:** Refactor code organization (A6–A13, Q4–Q11).
4. **Ongoing:** Maintain type safety and API contracts.

A re-audit after Sprint 1 is recommended to validate remediation effectiveness.
