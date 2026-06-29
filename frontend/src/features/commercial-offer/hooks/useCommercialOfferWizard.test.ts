import { act, cleanup, renderHook, waitFor } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { createElement, type PropsWithChildren, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import { WizardDraftProvider } from "@/features/commercial-offer/store/wizardDraftStore";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import { useCommercialOfferWizard } from "./useCommercialOfferWizard";

vi.mock("@/features/commercial-offer/store/draftStorage", () => ({
  draftStorage: {
    load: () => null,
    save: vi.fn(),
    clear: vi.fn(),
  },
}));

vi.mock("@/features/commercial-offer/api/commercialOfferApi", () => ({
  commercialOfferApi: {
    getManagers: vi.fn(),
    getDraft: vi.fn(),
    getBreakdown: vi.fn(),
    createDraft: vi.fn(),
    updateDraftPlates: vi.fn(),
    applyAiPlates: vi.fn(),
    resolveWidePlates: vi.fn(),
    updateDraftMeta: vi.fn(),
    calculateDraft: vi.fn(),
    generateFiles: vi.fn(),
    generateSchemaFiles: vi.fn(),
    saveDraft: vi.fn(),
  },
}));

const baseWizardState = {
  current_step: "plates" as const,
  can_proceed_to: [],
  next_required_action: "none" as const,
  validation_errors: [],
};

function makeDraft(overrides: Partial<CommercialDraftDetails> = {}): CommercialDraftDetails {
  const { metadata: metaOverrides, wizard_state: wizardOverrides, ...rest } = overrides;
  return {
    draft_id: "draft-test-1",
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data: [],
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    ...rest,
    metadata: {
      source_type: "text",
      original_text: "",
      ocr_text: "",
      input_text: "",
      accumulated_text: "",
      manager_id: 42,
      manager_name: "Менеджер",
      manager_phone: "",
      manager_email: "",
      client_name: "ООО Тест",
      discount_percent: 0,
      conditions_mode: "standard",
      delivery_conditions: "",
      payment_conditions: "",
      warnings: [],
      unparsed_lines: [],
      normalized_text: "",
      normalized_lines: [],
      wide_plate_lines: [],
      diagnostics: [],
      price_rows_count: 0,
      breakdown_tables_count: 0,
      total_sum: 0,
      plate_batches: [],
      wide_plates_resolved: true,
      last_source_filename: "",
      current_step: "plates",
      current_save_mode: null,
      execution_terms: "",
      logistics_cost: 0,
      ...metaOverrides,
    },
    wizard_state: { ...baseWizardState, ...wizardOverrides },
  };
}

function createWrapper(queryClient: QueryClient) {
  return function Wrapper({ children }: PropsWithChildren) {
    return createElement(
      QueryClientProvider,
      { client: queryClient },
      createElement(WizardDraftProvider, null, children as ReactNode),
    );
  };
}

describe("useCommercialOfferWizard", () => {
  let queryClient: QueryClient;

  beforeEach(() => {
    queryClient = new QueryClient({
      defaultOptions: {
        queries: { retry: false },
        mutations: { retry: false },
      },
    });
    vi.mocked(commercialOfferApi.getManagers).mockResolvedValue([]);
    vi.mocked(commercialOfferApi.getDraft).mockResolvedValue(makeDraft());
    vi.mocked(commercialOfferApi.getBreakdown).mockResolvedValue({
      draft_id: "draft-test-1",
      items: [],
    });
  });

  afterEach(() => {
    cleanup();
    queryClient.clear();
    vi.clearAllMocks();
  });

  it("starts on plates step", () => {
    const { result } = renderHook(() => useCommercialOfferWizard(), {
      wrapper: createWrapper(queryClient),
    });

    expect(result.current.state.currentStep).toBe("plates");
    expect(result.current.state.draftId).toBeNull();
  });

  it("advances wizard steps via dispatch", () => {
    const { result } = renderHook(() => useCommercialOfferWizard(), {
      wrapper: createWrapper(queryClient),
    });

    const steps = ["wide-plates", "manager", "client", "result"] as const;

    for (const step of steps) {
      act(() => {
        result.current.dispatch({ type: "set-step", step });
      });
      expect(result.current.state.currentStep).toBe(step);
    }
  });

  it("hydrates draft from draftQuery without jumping past plates", async () => {
    const draft = makeDraft({
      wizard_state: { ...baseWizardState, current_step: "result" },
    });
    vi.mocked(commercialOfferApi.getDraft).mockResolvedValue(draft);

    const { result } = renderHook(() => useCommercialOfferWizard(), {
      wrapper: createWrapper(queryClient),
    });

    act(() => {
      result.current.dispatch({ type: "set-draft-id", draftId: draft.draft_id });
    });

    await waitFor(() => {
      expect(result.current.draftQuery.isSuccess).toBe(true);
    });

    expect(result.current.state.currentStep).toBe("plates");
    expect(result.current.state.draftId).toBe(draft.draft_id);
    expect(result.current.currentDraft?.draft_id).toBe(draft.draft_id);
  });

  it("fetches breakdown only on result step when draft exists", async () => {
    const draft = makeDraft();
    vi.mocked(commercialOfferApi.getDraft).mockResolvedValue(draft);
    vi.mocked(commercialOfferApi.getBreakdown).mockResolvedValue({
      draft_id: "draft-test-1",
      items: [],
    });

    const { result } = renderHook(() => useCommercialOfferWizard(), {
      wrapper: createWrapper(queryClient),
    });

    act(() => {
      result.current.dispatch({ type: "set-draft-id", draftId: draft.draft_id });
      result.current.dispatch({ type: "set-step", step: "manager" });
    });

    await waitFor(() => {
      expect(result.current.draftQuery.isSuccess).toBe(true);
    });
    expect(commercialOfferApi.getBreakdown).not.toHaveBeenCalled();

    act(() => {
      result.current.dispatch({ type: "set-step", step: "result" });
    });

    await waitFor(() => {
      expect(commercialOfferApi.getBreakdown).toHaveBeenCalledWith(draft.draft_id);
      expect(result.current.breakdownQuery.isSuccess).toBe(true);
    });
  });

  it("createDraftMutation hydrates store and sets draft id", async () => {
    const createdDraft = makeDraft({ draft_id: "draft-created-99" });
    vi.mocked(commercialOfferApi.createDraft).mockResolvedValue(createdDraft);

    const { result } = renderHook(() => useCommercialOfferWizard(), {
      wrapper: createWrapper(queryClient),
    });

    await act(async () => {
      await result.current.createDraftMutation.mutateAsync({
        text: "Плита 1.2x3.0 — 10 шт",
        image: null,
      });
    });

    expect(result.current.state.draftId).toBe("draft-created-99");
    expect(result.current.state.clientName).toBe("ООО Тест");
    expect(result.current.currentDraft?.draft_id).toBe("draft-created-99");
  });
});
