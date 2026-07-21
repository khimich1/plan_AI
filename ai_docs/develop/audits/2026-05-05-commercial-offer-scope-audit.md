# Commercial Offer Scope Audit – Consolidated Report

**Date:** 2026-05-05  
**Scope:** `AppHeader.tsx`, `frontend/src/pages/commercial-offer-create/`, `frontend/src/features/commercial-offer/`, `app/api/v1/endpoints/commercial.py`, `app/services/commercial_service.py`, `app/services/commercial_workflow_service.py`, `app/schemas/commercial.py`, `draftStorage` (referenced by security audit)  
**Audited by:** senior-reviewer + security-auditor + reviewer (subagents)

---

## Executive Summary

**Health Score: 2.0 / 10** (Critical assessment required)

The commercial offer module exhibits severe architectural, security, and code quality defects requiring immediate remediation. Critical vulnerabilities in path traversal (S-H01), identity/access control (S-H02), and file upload enforcement (S-H03) present production security risks. Architectural issues stem from orchestration contract drift (A1), monolithic service design (A2), and missing dependency injection (A5). The module is not production-ready.

### Severity Aggregation

| Category | Critical | High | Medium | Low | Total |
|----------|----------|------|--------|-----|-------|
| Architecture | 1 | 4 | 6 | 3 | 14 |
| Security | 1 | 5 | 5 | 4 | 15 |
| Code Quality | 0 | 2 | 4 | 4 | 10 |
| **Total** | **2** | **11** | **15** | **11** | **39** |

### Health Score Calculation

- **Base:** 10.0
- **Critical issues:** 2 × (−2.0) = −4.0
- **High issues:** 11 issues; cap at −3.0
- **Medium issues:** 15 issues; cap at −1.0
- **Final:** 10.0 − 4.0 − 3.0 − 1.0 = **2.0 / 10**

---

## Critical Issues

### A1 – Orchestration Contract Drift (Critical – Architecture)

**Location:** `frontend/src/pages/commercial-offer-create/`, `frontend/src/features/commercial-offer/`, `app/services/commercial_workflow_service.py`

**Issue:**  
Two competing orchestrators manage wizard state without formal contract:
- **Client side:** UI steps indexed by `WizardStepId` enum; step flow driven by React state and query mutations.
- **Server side:** `CommercialWorkflowService.calculate_draft()` and `current_step` metadata manage business logic progression.

**Impact:**
- UI may skip POST /calculate when navigating steps; server metadata becomes stale.
- Draft state diverges from server intent; financial calculations may not execute when expected.
- Client-side financial totals in `draft.totals` may conflict with server-recalculated values.
- Contract not formally defined; field names are implicit.

**Fix Suggestion:**
1. Define formal state machine contract (Pydantic schema):
   ```python
   class WizardStateResponse(BaseModel):
       current_step: WizardStepId
       can_proceed_to: list[WizardStepId]
       next_required_action: str
       validation_errors: list[str] = []
   ```
2. Server-side `current_step` as source of truth; POST /calculate always required before client transitions.
3. Client mutation gates UI buttons on response `can_proceed_to`.
4. Synchronize `WizardStepId` enum between FE and BE.

---

### S-H01 – Path Traversal in DraftStore (Critical – Security)

**Location:** `app/services/commercial_service.py` (DraftStore path join), referenced by `draftStorage.ts`

**Issue:**  
If draft files stored as `DRAFT_DIR / {user_id} / {draft_id} / {filename}`, attacker supplies `filename` with `../` sequences to escape directory.

**Impact:**
- Arbitrary file read/write within user's directory tree.
- Potential disclosure of other users' draft files if traversal crosses ownership boundary.

**Fix Suggestion:**
1. Sanitize filename; reject any path containing `..`, `/`, `\`, or absolute paths:
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
2. Or use deterministic filename: `{draft_id}_{field_name}.{ext}` with no user input in path.

---

## High Priority Issues

### A2 – CommercialWorkflowService God-Module (High – Architecture)

**Location:** `app/services/commercial_workflow_service.py` (~940 lines)

**Issue:**  
`CommercialWorkflowService` owns calculation, file generation, draft state resolution, and wide-plate orchestration—violating single responsibility principle.

**Impact:**
- Module difficult to test in isolation; dependencies span multiple concerns.
- New business logic adds to same god-module.
- Circular dependency risk with `commercial.py` endpoint calling private `_resolve_generated_file()`.

**Fix Suggestion:**
1. Extract sub-services:
   - `DraftCalculationService`: compute totals, validate order data.
   - `WizardOrchestrationService`: manage state transitions, determine `current_step`.
   - `OfferGenerationService`: coordinate file generation and template rendering.
2. Use factory pattern in `commercial.py` to compose services via DI.
3. Define clear interfaces (protocols) for each sub-service.

---

### A3 – Duplicate Money Logic (Client vs Server) (High – Architecture)

**Location:** `frontend/src/features/commercial-offer/`, `app/services/commercial_workflow_service.py`

**Issue:**  
Client computes totals (subtotal, tax, total) independently; server also calculates via `calculate_draft()`. `draft.totals` populated by POST /calculate, but client may display different values before sync.

**Impact:**
- User sees inconsistent pricing if server recalculation differs from client estimate.
- Maintenance burden: money logic changes require updates in both FE and BE.
- Risk of desync on slow networks or stale caching.

**Fix Suggestion:**
1. Server is source of truth: POST /calculate returns `CalculationResult` with all totals.
2. Client mutation uses updated state management (not stale `onSuccess`) to update Redux/Zustand with server-calculated `draft.totals`.
3. Remove client-side total computation; display server `draft.totals` only.
4. Validate on submit that client form data matches server totals (ETag or checksum).

---

### A4 – AppHeader Tightly Coupled to Wizard Store (High – Architecture)

**Location:** `frontend/src/app/layout/AppHeader.tsx`

**Issue:**  
`AppHeader` directly selects and mutates wizard store state (`setStep`, `resetWizard`, `close`). State tied to single use-case.

**Impact:**
- Reusability limited: AppHeader cannot be used in other flows without conditional logic.
- Store leak: UI chrome (header) coupled to business domain (wizard).
- Risk of state corruption if close/reset logic differs in other pages.

**Fix Suggestion:**
1. Inject step-agnostic callbacks: `onStepChange(step)`, `onReset()`, `onClose()`.
2. Wrap AppHeader in container component that supplies callbacks from context.
3. Move wizard-specific logic (e.g., reset confirmation) to page level.
4. Use composition: `<AppLayout onClose={handleClose} />` instead of direct store binding.

---

### A5 – API Layer Calls Private Workflow Method (High – Architecture)

**Location:** `app/api/v1/endpoints/commercial.py` (~279–284), `app/services/commercial_workflow_service.py`

**Issue:**  
`commercial.py` calls `workflow._resolve_generated_file()` (underscore-prefixed private method) directly instead of using public interface.

**Impact:**
- Violates encapsulation; endpoint coupled to private implementation details.
- Refactoring workflow internals risks breaking endpoint without clear dependency graph.
- No contract specification for what `_resolve_generated_file()` does.

**Fix Suggestion:**
1. Create public method `CommercialWorkflowService.get_or_generate_file(draft_id: UUID) -> FileModel`.
2. Endpoint calls `workflow.get_or_generate_file()`.
3. Mark internal helper `_resolve_generated_file()` as private in docstring or move to module level.

---

### S-H02 – IDOR: No Per-User Draft Ownership Check (High – Security)

**Location:** `app/api/v1/endpoints/commercial.py`

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

### S-H03 – Unbounded File Uploads, No Rate Limits (High – Security)

**Location:** `app/api/v1/endpoints/commercial.py` (file upload handlers)

**Issue:**  
File upload endpoints (OCR image upload, invoice scan) do not enforce upload size limits, rate limiting per user, or OCR cost/quota controls.

**Impact:**
- Disk exhaustion: attacker uploads 10 GB files, service runs out of space.
- OCR cost abuse: attacker triggers 1000s of OCR requests, inflating API costs.
- DoS vector.

**Fix Suggestion:**
1. Add upload size limit (50 MB per file):
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
2. Implement rate limiting (e.g., `slowapi` or Redis):
   - Max 10 OCR uploads per user per hour.
   - Max 500 MB per user per day.
3. Track OCR costs; alert or reject if user exceeds budget.

---

### S-H04 – Weak Upload Validation (High – Security)

**Location:** `app/api/v1/endpoints/commercial.py`

**Issue:**  
File upload endpoints validate only `Content-Type` header; attacker sets `Content-Type: image/jpeg` and uploads executable or malicious content.

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

### S-H05 – Error Detail Leakage (High – Security)

**Location:** `app/api/v1/endpoints/commercial.py`

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

### Q1 – React Query v5 useQuery onSuccess Invalid (High – Code Quality)

**Location:** `frontend/src/features/commercial-offer/`

**Issue:**  
`useQuery` with `onSuccess` callback is deprecated in React Query v5. Callback may not fire reliably, or closure captures stale state.

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
2. Or use `useMutation` with `onSuccess` for mutations (more appropriate).

---

### Q2 – No POST /calculate from Wizard (High – Code Quality)

**Location:** `frontend/src/features/commercial-offer/`

**Issue:**  
`calculateMutation` hook is exported but never called from wizard UI flow. Wizard steps proceed without recalculating server state.

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
         showError(`Cannot proceed: ${result.reason}`);
         return;
       }
       setStep(nextStep);
     } catch (e) {
       showError("Calculation failed");
     }
   };
   ```

---

### Q3 – Mutation onSuccess Closure Stale State (High – Code Quality)

**Location:** `frontend/src/features/commercial-offer/`

**Issue:**  
If `useMutation` created in hook without re-running on dependency change, `onSuccess` callback captures stale variables (form data, user context).

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
3. Handle state updates in `useEffect` after query re-fetch.

---

## Medium Priority Issues

### A6 – Wizard Orchestrator God-Component (Medium – Architecture)

**Location:** `frontend/src/features/commercial-offer/CommercialOfferWizard.tsx`

**Issue:**  
Wizard component combines step rendering, form validation, API calls, and state management in single file.

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

### A7 – Loose API Payload Types (Medium – Architecture)

**Location:** `frontend/src/features/commercial-offer/types/commercialOffer.ts`

**Issue:**  
Frontend `order_data` or form state typed as `Record<string, unknown>` instead of specific schema matching backend `DraftOrder`.

**Impact:**
- No type checking; typos in field names not caught at compile time.
- Mismatch between client and server field names (e.g., `plateLines` vs `plate_lines`).

**Fix Suggestion:**
```typescript
interface DraftOrderData {
  customer_name: string;
  order_date: string;
  plate_lines: PlateLineData[];
  // ... fields matching backend schema
}
```
Use consistently in form state.

---

### A8 – Heavy Client Persistence Full State to sessionStorage (Medium – Architecture)

**Location:** `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`, `draftStorage.ts`

**Issue:**  
All wizard state persisted to sessionStorage; no distinction between critical (identity, auth) and ephemeral (UI state).

**Impact:**
- Session storage bloat; performance degradation.
- XSS vulnerability if malicious script accesses sessionStorage (stored state accessible).
- Risk of stale state on browser refresh.

**Fix Suggestion:**
1. Persist only essential state: `current_step`, `draft_id`, `user_changes` (not full Redux tree).
2. Use `httpOnly` cookies for auth tokens; avoid sensitive data in sessionStorage.
3. On hydration, fetch latest draft via `GET /draft/{id}` to sync server state.

---

### A9 – Leaky Abstraction Legacy Context in commercial_service.py (Medium – Architecture)

**Location:** `app/services/commercial_service.py`

**Issue:**  
`CommercialService` accesses private method `FileGenerationService._legacy_order_context()` instead of calling public API.

**Impact:**
- Encapsulation violation; refactoring `FileGenerationService` risks breaking `CommercialService`.
- No version contract; unclear if `_legacy_order_context` is temporary or permanent.

**Fix Suggestion:**
1. Create public method `FileGenerationService.get_order_context(draft: Draft) -> OrderContext`.
2. Call `generation_service.get_order_context(draft)` instead of `._legacy_order_context()`.

---

### A10 – Duplicate VAT/Totals Logic vs Backend (Medium – Architecture)

**Location:** `frontend/src/features/commercial-offer/steps/CalculationResultStep.tsx`

**Issue:**  
Client component independently computes VAT and totals; server also calculates in `CommercialWorkflowService`.

**Impact:**
- Two sources of truth for financial calculations.
- If server logic updates (e.g., VAT rate change), client UI becomes stale.
- Maintenance burden.

**Fix Suggestion:**
- Remove client-side calculation; display server `draft.totals` only.
- Server is authoritative; client is presentation layer.

---

### A11 – Triplicated Step Ordering (Medium – Architecture)

**Location:** `frontend/src/features/commercial-offer/lib/wizardStepOrder.ts`, component `WizardProgress.tsx`, `app/schemas/commercial.py`

**Issue:**  
Step definitions repeated in three places: FE `wizardStepOrder.ts`, `WizardProgress.tsx`, BE `commercial.py`.

**Impact:**
- Changes to step order require updates in three places; easy to miss one and break consistency.
- Single source of truth missing.

**Fix Suggestion:**
1. Define canonical `WizardStep` enum in backend `schemas/commercial.py`.
2. Export to OpenAPI/TypeScript generator; FE imports from generated types.
3. Both FE and BE reference same enum.

---

### S-M01 – Weak Upload Validation (Medium – Security)

**Location:** `app/api/v1/endpoints/commercial.py`

**Issue:**  
File upload endpoints validate only `Content-Type` header.

**Impact:**
- Malware uploaded disguised as image.
- If uploaded file is later processed, vulnerability chains.

**Fix Suggestion:**  
(Covered in S-H04 above; duplicate for completeness)

---

### S-M02 – Error Detail Leakage (Medium – Security)

**Location:** `app/api/v1/endpoints/commercial.py`

**Issue:**  
Exception handlers expose internal error messages.

**Impact:**
- Information disclosure; attacker learns infrastructure details.

**Fix Suggestion:**  
(Covered in S-H05 above; duplicate for completeness)

---

### S-M03 – Missing Validation Schema for /parse and /generate (Medium – Security)

**Location:** `app/api/v1/endpoints/commercial.py`

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
    return InvoiceParseResult(...)
```

---

### S-M04 – No Dependency Injection; Service Per Request (Medium – Security)

**Location:** `app/api/v1/endpoints/commercial.py`

**Issue:**  
Services (`CommercialService`, `CommercialWorkflowService`) instantiated inline in endpoints or passed as function arguments instead of using FastAPI `Depends()`.

**Impact:**
- Harder to test: mock services cannot be injected into route handlers.
- Boilerplate: each endpoint repeats service instantiation.
- Coupling: endpoint imports and instantiates specific service classes.

**Fix Suggestion:**
```python
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

### Q4 – Long Methods in CommercialWorkflowService (Medium – Code Quality)

**Location:** `app/services/commercial_workflow_service.py`

**Issue:**
- `_build_order_data()` >100 lines; combines name assembly, item collection, totals, metadata.
- `resolve_wide_plates()` >80 lines; mixes validation, normalization, state management.

**Impact:**
- Hard to test individual logic paths.
- Difficult to understand and maintain.
- Risk of bugs when refactoring.

**Fix Suggestion:**
1. Extract helpers:
   ```python
   def _extract_customer_data(order_data) -> CustomerInfo: ...
   def _calculate_totals(items) -> TotalsSummary: ...
   def _normalize_wide_plates(raw_plates) -> list[WidePlate]: ...
   ```
2. Each helper tested independently.

---

### Q5 – Duplicate try/catch Pattern (Medium – Code Quality)

**Location:** `app/services/commercial_workflow_service.py`

**Issue:**  
Multiple methods repeat similar error handling (try/catch + logging + raise).

**Impact:**
- Code duplication; difficult to maintain error handling consistently.

**Fix Suggestion:**
```python
def handle_service_error(exc: Exception, operation: str) -> None:
    logger.error(f"{operation} failed: {exc}")
    if isinstance(exc, ValueError):
        raise HTTPException(400, f"Invalid {operation} data")
    raise HTTPException(500, "Internal error")
```

---

## Low Priority Issues

### A12 – Validation Split Zod vs Imperative (Low – Architecture)

**Location:** `frontend/src/features/commercial-offer/steps/PlateInputStep.tsx`

**Issue:**  
Some validation done via Zod schemas, other validation imperative in component handlers.

**Impact:**
- Inconsistent validation approach; hard to understand what rules apply.
- Potential gap if imperative validation is missed.

---

### A13 – Thin Page Wrapper (Low – Architecture)

**Location:** `frontend/src/pages/commercial-offer-create/`

**Issue:**  
Page component is minimal wrapper; real logic in nested features.

**Impact:**
- Navigation and page-level concerns not cleanly separated.

---

### A14 – Mutation Merge Closure Staleness (Low – Architecture)

**Location:** `frontend/src/features/commercial-offer/useCommercialOfferWizard.ts`

**Issue:**  
If custom hook creates mutation without dependency array, closure captures stale form state.

**Impact:**
- Subtle bug if form updates between mutation creation and invocation.

---

### S-L01 – Client-Side Totals Not Validated (Low – Security)

**Location:** `frontend/src/features/commercial-offer/`

**Issue:**  
Client-side totals not validated against server; users could theoretically manipulate form before POST.

**Impact:**
- Low if server is authoritative; mitigated if server recalculates.

---

### S-L02 – Absolute URLs in Error Messages (Low – Security)

**Location:** `app/services/`

**Issue:**  
Absolute URLs in error messages or logs could disclose environment.

**Impact:**
- Minor information disclosure if logs exposed.

---

### S-L03 – SessionStorage for Wizard State (Low – Security)

**Location:** `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`

**Issue:**  
Wizard state persisted in sessionStorage; XSS vulnerability if sensitive data stored.

**Impact:**
- Low if no sensitive data (passwords, tokens) stored; session-specific state is acceptable.

---

### S-L04 – Download File Path Validation (Low – Security)

**Location:** `app/api/v1/endpoints/commercial.py` (file download handlers)

**Issue:**  
Download file path constructed from user input; should validate that served file is within expected directory.

**Impact:**
- Mitigated by S-H01 path traversal fix.

---

### Q6 – Dead Code: lastPlateMode Variable (Low – Code Quality)

**Location:** `frontend/src/features/commercial-offer/store/wizardDraftStore.tsx`

**Issue:**  
Unused state variable.

**Impact:**
- Code clutter.

---

### Q7 – calculateMutation Exported but Unused (Low – Code Quality)

**Location:** `frontend/src/features/commercial-offer/`

**Issue:**  
`calculateMutation` hook exported from module but never imported elsewhere.

**Impact:**
- Misleads future developers; suggests mutation is used elsewhere.

---

### Q8 – Type Mismatch in resolveWidePlatesMutation (Low – Code Quality)

**Location:** `frontend/src/features/commercial-offer/`

**Issue:**  
`resolveWidePlatesMutation` typed as `useMutation<Result, Error, WidePlateData>`, but called with `lineId` instead of full plate data.

**Impact:**
- Type confusion; mutation may receive unexpected argument type.

---

### Q9 – VAT Hardcoded (Low – Code Quality)

**Location:** `app/services/commercial_workflow_service.py`

**Issue:**  
VAT rate hardcoded as 18%; should be moved to config.

**Impact:**
- Difficult to update VAT rate without code change.

---

## Priority Matrix

| ID | Issue | Severity | Category | Estimated Effort | Priority | Action |
|---|---|---|---|---|---|---|
| A1 | Orchestration Contract Drift | Critical | Arch | 21 days | **P0** | Define state machine contract |
| S-H01 | Path Traversal in DraftStore | Critical | Security | 2 days | **P0** | Sanitize file paths |
| S-H02 | IDOR: No Draft Ownership | High | Security | 3 days | **P0** | Add user_id check |
| S-H03 | Unbounded Uploads | High | Security | 3 days | **P0** | Add size/rate limits |
| S-H04 | Weak Upload Validation | High | Security | 2 days | **P0** | Validate file signatures |
| S-H05 | Error Detail Leakage | High | Security | 1 day | **P0** | Generic error responses |
| A2 | God-Module CommercialWorkflowService | High | Arch | 8 days | **P1** | Extract sub-services |
| A3 | Duplicate Money Logic | High | Arch | 5 days | **P1** | Unify to server |
| A4 | AppHeader Coupling | High | Arch | 2 days | **P1** | Inject callbacks |
| A5 | Private Method Call | High | Arch | 1 day | **P1** | Expose public API |
| Q1 | React Query onSuccess | High | QA | 1 day | **P1** | Migrate to useEffect |
| Q2 | No POST /calculate from Wizard | High | QA | 2 days | **P1** | Integrate mutation |
| Q3 | onSuccess Stale Closure | High | QA | 2 days | **P1** | Refactor hooks |
| A6 | Wizard God-Component | Medium | Arch | 5 days | **P2** | Extract step components |
| A7 | Loose API Payload Types | Medium | Arch | 2 days | **P2** | Add TypeScript interface |
| A8 | Heavy sessionStorage Persistence | Medium | Arch | 2 days | **P2** | Minimize persisted state |
| A9 | Leaky Abstraction (legacy context) | Medium | Arch | 1 day | **P2** | Expose public API |
| A10 | Duplicate VAT/Totals Logic | Medium | Arch | 2 days | **P2** | Remove client calculation |
| A11 | Triplicated Step Ordering | Medium | Arch | 2 days | **P2** | Canonicalize in backend |
| S-M03 | Missing Validation Schema | Medium | Security | 2 days | **P2** | Add Pydantic schemas |
| S-M04 | No DI in Endpoints | Medium | Security | 3 days | **P2** | Migrate to Depends() |
| Q4 | Long Methods | Medium | QA | 3 days | **P2** | Extract helpers |
| Q5 | Duplicate try/catch | Medium | QA | 1 day | **P2** | Centralize error handling |
| A12 | Validation Split (Zod vs imperative) | Low | Arch | 1 day | **P3** | Unify validation approach |
| A13 | Thin Page Wrapper | Low | Arch | <1 day | **P3** | Add page-level logic |
| A14 | Mutation Closure Staleness | Low | Arch | 1 day | **P3** | Verify dependencies |
| S-L01–L04 | Low-Severity Security | Low | Security | 1 day | **P3** | Document or mitigate |
| Q6–Q9 | Code Cleanup | Low | QA | 1 day | **P3** | Remove dead code, fix types |

---

## Next Steps

### Immediate (This Week) – Critical & High-Security

1. **S-H01 – Path Traversal**
   - Implement `safe_join()` path sanitization in `commercial_service.py`.
   - Add test: attempt `../` in filename; verify rejection.

2. **S-H02 – IDOR**
   - Add `user_id` to `Draft` model.
   - Add dependency: `verify_draft_ownership(draft_id, current_user)`.
   - Apply to all draft endpoints.

3. **S-H03, S-H04 – Upload Enforcement**
   - Add 50 MB file size limit.
   - Implement rate limiting (10 OCR per user per hour, 500 MB per user per day).
   - Validate file signatures (magic bytes).

4. **S-H05 – Error Leakage**
   - Wrap endpoint exception handlers; return generic error messages to client.
   - Log detailed errors server-side only.

### Sprint 1 (Next 2 Weeks) – Architectural High Priority

5. **A1 – Orchestration Contract Drift**
   - Define `WizardStateResponse` schema.
   - Sync `WizardStepId` enum between FE and BE.
   - Client gates step transitions on server `can_proceed_to` response.

6. **A2 – God-Module Refactoring**
   - Extract `DraftCalculationService`, `WizardOrchestrationService`, `OfferGenerationService`.
   - Update `commercial.py` to use DI.

7. **A3 – Duplicate Money Logic**
   - Remove client-side total computation; display server `draft.totals`.
   - Add sync check in `useEffect`.

8. **A4, A5 – Decoupling & Encapsulation**
   - Inject callbacks into AppHeader.
   - Expose public `get_or_generate_file()` method in workflow service.

9. **Q1, Q2, Q3 – React Query & Mutation Fixes**
   - Migrate `onSuccess` to `useEffect` pattern.
   - Integrate `calculateMutation` into wizard flow.
   - Ensure server state is authoritative.

### Sprint 2 (Weeks 3–4) – Medium-Priority Refactoring

10. **A6–A11 – Code Organization & Type Safety**
    - Extract wizard step components (A6).
    - Add TypeScript interfaces for order data (A7).
    - Minimize sessionStorage persistence (A8).
    - Expose public API for legacy context (A9).
    - Remove client VAT/totals (A10).
    - Canonicalize step ordering in backend (A11).

11. **S-M03, S-M04 – Schemas & Dependency Injection**
    - Add Pydantic response schemas for `/parse` and `/generate`.
    - Full `Depends()` migration for services.

12. **Q4, Q5 – Code Quality**
    - Extract long methods into helpers.
    - Centralize error handling.

### Backlog – Low Priority

13. Unify validation approach (A12).
14. Enhance page-level logic (A13).
15. Fix mutation closure dependencies (A14).
16. Address low-security concerns (S-L01–L04).
17. Remove dead code and fix type mismatches (Q6–Q9).

---

## Remediation Workflow

**To remediate issues, use one of the following commands:**

- **`/refactor`** – Orchestrate refactoring workflow for architectural and code quality improvements (A2, A3, A4, A5, A6–A11, Q4, Q5).
- **`/implement`** – Create or update code for specific features (security patches, schema additions, DI setup).
- **`/review`** – Code review before committing to verify fixes.

**Example workflow for critical security fixes:**
1. User confirms readiness: "Please fix S-H01, S-H02, S-H03, S-H04, S-H05"
2. Run `/implement` → fixes security issues + tests.
3. Run `/review` → verifies quality.
4. Commit and verify.

**Example workflow for architectural refactoring:**
1. User prioritizes Sprint 1 items: "Fix A1, A2, A3, A4, A5"
2. Run `/refactor` → plan + implement + test + verify.
3. Commit and document architectural decisions.

**Phase 5 (Post-Remediation) Only on User Confirm:**
After all P0 and P1 items fixed and verified, run comprehensive re-audit to validate health score improvement.

---

## Recommendation Summary

**Overall Health: 2.0 / 10 – Action Required Immediately**

The commercial offer module has **critical security vulnerabilities** (path traversal, IDOR, unbounded uploads), **architectural contract drift** (orchestration), and **god-module design** that must be remedied before any production deployment.

**Recommended remediation sequence:**

1. **This week:** Close security holes (S-H01–S-H05). Estimated effort: 1 week.
2. **Next 2 weeks (Sprint 1):** Fix orchestration (A1), god-module (A2–A5), React Query (Q1–Q3). Estimated effort: 2 weeks.
3. **Weeks 3–4 (Sprint 2):** Refactor code organization (A6–A11), add schemas (S-M03), DI (S-M04), long methods (Q4–Q5). Estimated effort: 2 weeks.
4. **Ongoing:** Maintain type safety and API contracts.

**Total estimated effort for P0 + P1 + P2:** ~5 weeks.

After Sprint 1 completion, a re-audit is **recommended** to validate that critical and high-priority security issues are resolved and architectural contract is established.

---

**Audit completed:** 2026-05-05  
**Next scheduled audit:** After Sprint 1 remediation (target: 2026-05-19)
