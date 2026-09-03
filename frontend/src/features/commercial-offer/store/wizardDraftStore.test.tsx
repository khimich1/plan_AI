import { act, cleanup, render, screen } from "@testing-library/react";
import type { MutableRefObject } from "react";
import { afterEach, describe, expect, it, vi, beforeEach } from "vitest";
import type {
  CommercialDraftDetails,
  CommercialWizardState,
  WizardStepId,
  WizardStoreState,
} from "@/features/commercial-offer/types/commercialOffer";
import {
  getWizardStepOrder,
  shouldSkipClientStep,
} from "@/features/commercial-offer/lib/wizardStepOrder";
import { WizardDraftProvider, useWizardDraftStore } from "@/features/commercial-offer/store/wizardDraftStore";

vi.mock("@/features/commercial-offer/store/draftStorage", () => ({
  draftStorage: {
    load: () => null,
    save: vi.fn(),
    clear: vi.fn(),
  },
}));

const baseWizardState: CommercialWizardState = {
  current_step: "plates",
  can_proceed_to: [],
  next_required_action: "none",
  validation_errors: [],
};

const baseMetadata = (): CommercialDraftDetails["metadata"] => ({
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
});

function makeDraft(overrides: Partial<CommercialDraftDetails>): CommercialDraftDetails {
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
    metadata: { ...baseMetadata(), ...metaOverrides },
    wizard_state: { ...baseWizardState, ...wizardOverrides },
  };
}

type WizardDraftDispatch = ReturnType<typeof useWizardDraftStore>["dispatch"];

/** MNA-502 — append-cycle fields on store (not implemented yet). */
type AppendCycleStoreFields = {
  isPickingProductType?: boolean;
};

type StoreSnapshot = WizardStoreState & AppendCycleStoreFields;

/** Dispatch append actions before reducer union grows (RED). */
type AppendCycleAction =
  | { type: "start-append-cycle" }
  | { type: "cancel-append-pick" }
  | { type: "set-product-type"; productType: WizardStoreState["productType"] }
  | { type: "set-step"; step: WizardStepId }
  | { type: "hydrate-draft"; payload: CommercialDraftDetails; refreshBatchText?: boolean }
  | { type: "set-source"; text: string; imageName: string | null }
  | { type: "set-batch-review-text"; text: string }
  | { type: "reset" };

const dispatchAppend = (
  dispatch: WizardDraftDispatch | null,
  action: AppendCycleAction,
): void => {
  dispatch?.(action as Parameters<WizardDraftDispatch>[0]);
};

const asAppendState = (state: WizardStoreState | null | undefined): StoreSnapshot | null =>
  (state as StoreSnapshot | null | undefined) ?? null;

/** Доступ к dispatch и state без user-event: выставляем ref в первом рендере. */
function Harness({
  actionRef,
  stateRef,
}: {
  actionRef: MutableRefObject<WizardDraftDispatch | null>;
  stateRef?: MutableRefObject<ReturnType<typeof useWizardDraftStore>["state"] | null>;
}) {
  const { dispatch, state } = useWizardDraftStore();
  actionRef.current = dispatch;
  if (stateRef) {
    stateRef.current = state;
  }
  return <span data-testid="current-step">{state.currentStep}</span>;
}

describe("WizardDraftProvider hydrate-draft step merge", () => {
  let dispatchRef: MutableRefObject<WizardDraftDispatch | null>;

  beforeEach(() => {
    dispatchRef = { current: null };
  });

  afterEach(() => {
    cleanup();
  });

  it("не переходит с plates на client после hydrate-draft (только через «Обработать»)", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} />
      </WizardDraftProvider>,
    );

    expect(screen.getByTestId("current-step")).toHaveTextContent("plates");

    await act(async () => {
      dispatchRef.current?.({
        type: "hydrate-draft",
        payload: makeDraft({
          wizard_state: { ...baseWizardState, current_step: "client" },
        }),
      });
    });

    expect(screen.getByTestId("current-step")).toHaveTextContent("plates");
  });

  it("маппит legacy wide-plates в plates при hydrate-draft", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} />
      </WizardDraftProvider>,
    );

    await act(async () => {
      dispatchRef.current?.({
        type: "hydrate-draft",
        payload: makeDraft({
          wizard_state: { ...baseWizardState, current_step: "wide-plates" as unknown as WizardStepId },
        }),
      });
    });

    expect(screen.getByTestId("current-step")).toHaveTextContent("plates");
  });

  it("не поднимает шаг с plates до result при hydrate-draft (остаётся на plates)", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} />
      </WizardDraftProvider>,
    );

    expect(screen.getByTestId("current-step")).toHaveTextContent("plates");

    await act(async () => {
      dispatchRef.current?.({
        type: "hydrate-draft",
        payload: makeDraft({
          wizard_state: { ...baseWizardState, current_step: "result" },
        }),
      });
    });

    expect(screen.getByTestId("current-step")).toHaveTextContent("plates");
  });

  it("для обоих валидных шагов берёт максимум (local client, legacy server manager → client)", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} />
      </WizardDraftProvider>,
    );

    await act(async () => {
      dispatchRef.current?.({ type: "set-step", step: "client" });
    });
    expect(screen.getByTestId("current-step")).toHaveTextContent("client");

    await act(async () => {
      dispatchRef.current?.({
        type: "hydrate-draft",
        payload: makeDraft({
          wizard_state: { ...baseWizardState, current_step: "manager" as unknown as WizardStepId },
        }),
      });
    });

    expect(screen.getByTestId("current-step")).toHaveTextContent("client");
  });

  it("игнорирует неизвестный серверный шаг и оставляет локальный", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} />
      </WizardDraftProvider>,
    );

    await act(async () => {
      dispatchRef.current?.({ type: "set-step", step: "client" });
    });

    await act(async () => {
      dispatchRef.current?.({
        type: "hydrate-draft",
        payload: makeDraft({
          wizard_state: {
            ...baseWizardState,
            current_step: "not-a-real-step" as WizardStepId,
          },
        }),
      });
    });

    expect(screen.getByTestId("current-step")).toHaveTextContent("client");
  });

  it("sync-after-wide-plates обновляет batchReviewText и сбрасывает widePlateActions", async () => {
    const stateRef: MutableRefObject<ReturnType<typeof useWizardDraftStore>["state"] | null> = { current: null };
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} stateRef={stateRef} />
      </WizardDraftProvider>,
    );

    await act(async () => {
      dispatchRef.current?.({
        type: "start-batch-review",
        payload: makeDraft({
          metadata: {
            plate_batches: [
              {
                source_type: "image",
                original_text: "",
                normalized_text: "ПБ 63-15-8 3\nПБ 58-15-8 10",
                ocr_text: "",
                filename: "photo.jpg",
              },
            ],
            normalized_text: "ПБ 63-15-8 3\nПБ 58-15-8 10",
            wide_plate_lines: [{ id: "w1", line: "ПБ 63-15-8 3", qty: 3 }],
            wide_plates_resolved: false,
          },
        }),
      });
      dispatchRef.current?.({
        type: "set-wide-action",
        lineId: "w1",
        action: "exclude",
        replacementText: "",
      });
    });

    expect(stateRef.current?.batchReviewText).toContain("ПБ 63-15-8 3");
    expect(stateRef.current?.widePlateActions.w1?.action).toBe("exclude");

    await act(async () => {
      dispatchRef.current?.({
        type: "sync-after-wide-plates",
        payload: makeDraft({
          metadata: {
            plate_batches: [
              {
                source_type: "image",
                original_text: "",
                normalized_text: "ПБ 58-15-8 10",
                ocr_text: "",
                filename: "photo.jpg",
              },
            ],
            normalized_text: "ПБ 58-15-8 10",
            wide_plate_lines: [],
            wide_plates_resolved: true,
          },
          order_data: [{ name: "ПБ 58-15-8 10", qty: 10 }],
        }),
      });
    });

    expect(stateRef.current?.batchReviewText).toBe("ПБ 58-15-8 10");
    expect(stateRef.current?.normalizedText).toBe("ПБ 58-15-8 10");
    expect(stateRef.current?.widePlateActions).toEqual({});
  });
});

/**
 * MNA-502 RED — Wizard orchestration: append loop + sticky header + skip client.
 * Expected store contract (worker):
 * - `start-append-cycle`: keep draftId + sticky client/discount/manager/conditions;
 *   clear ephemeral input; set `isPickingProductType`.
 * - `set-product-type` after pick: input step for new type; clear picking; sticky retained.
 * - `reset` («Создать новое КП»): full wipe including picking flag + sticky header.
 * Wizard (separate) must use shouldSkipClientStep / getWizardStepOrder({ skipClient })
 * and wire CalculationResultStep append/undo/delete → API + hydrate.
 */
describe("WizardDraftProvider append cycle sticky header (MNA-502)", () => {
  let dispatchRef: MutableRefObject<WizardDraftDispatch | null>;
  let stateRef: MutableRefObject<ReturnType<typeof useWizardDraftStore>["state"] | null>;

  const seedStickyResultState = async () => {
    await act(async () => {
      dispatchAppend(dispatchRef.current, {
        type: "hydrate-draft",
        payload: makeDraft({
          draft_id: "draft-append-sticky",
          wizard_state: { ...baseWizardState, current_step: "result" },
          metadata: {
            product_type: "plates",
            client_name: "ООО Стикки",
            discount_percent: 7.5,
            manager_id: 99,
            conditions_mode: "custom",
            delivery_conditions: "Самовывоз",
            payment_conditions: "100% предоплата",
            append_batches: [{ batch_id: "b1", product_type: "plates", line_ids: ["ln1"] }],
            resume_kp_id: null,
            current_step: "result",
          },
        }),
      });
      dispatchAppend(dispatchRef.current, { type: "set-step", step: "result" });
      dispatchAppend(dispatchRef.current, {
        type: "set-source",
        text: "остаток ввода с прошлого цикла",
        imageName: "old.png",
      });
      dispatchAppend(dispatchRef.current, {
        type: "set-batch-review-text",
        text: "batch review leftover",
      });
    });
  };

  beforeEach(() => {
    dispatchRef = { current: null };
    stateRef = { current: null };
  });

  afterEach(() => {
    cleanup();
  });

  it("start-append-cycle opens picker while retaining draftId and sticky header (client/discount)", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} stateRef={stateRef} />
      </WizardDraftProvider>,
    );

    await seedStickyResultState();

    await act(async () => {
      dispatchAppend(dispatchRef.current, { type: "start-append-cycle" });
    });

    const snap = asAppendState(stateRef.current);
    expect(snap?.draftId).toBe("draft-append-sticky");
    expect(snap?.clientName).toBe("ООО Стикки");
    expect(snap?.discountPercent).toBe(7.5);
    expect(snap?.managerId).toBe(99);
    expect(snap?.conditionsMode).toBe("custom");
    expect(snap?.deliveryConditions).toBe("Самовывоз");
    expect(snap?.paymentConditions).toBe("100% предоплата");
    expect(snap?.isPickingProductType).toBe(true);
    expect(snap?.sourceText).toBe("");
    expect(snap?.selectedImageName).toBeNull();
    expect(snap?.batchReviewText).toBe("");
    expect(snap?.pendingBatchReview).toBe(false);
  });

  it("append → pick product type → input step; client sticky and discount retained", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} stateRef={stateRef} />
      </WizardDraftProvider>,
    );

    await seedStickyResultState();

    await act(async () => {
      dispatchAppend(dispatchRef.current, { type: "start-append-cycle" });
      dispatchAppend(dispatchRef.current, { type: "set-product-type", productType: "piles" });
    });

    const snap = asAppendState(stateRef.current);
    expect(snap?.isPickingProductType).toBe(false);
    expect(snap?.productType).toBe("piles");
    expect(snap?.currentStep).toBe("piles");
    expect(snap?.draftId).toBe("draft-append-sticky");
    expect(snap?.clientName).toBe("ООО Стикки");
    expect(snap?.discountPercent).toBe(7.5);
    expect(screen.getByTestId("current-step")).toHaveTextContent("piles");
  });

  it("sticky header after append implies skip-client step order (picker → input → result)", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} stateRef={stateRef} />
      </WizardDraftProvider>,
    );

    await seedStickyResultState();

    await act(async () => {
      dispatchAppend(dispatchRef.current, { type: "start-append-cycle" });
    });
    expect(asAppendState(stateRef.current)?.isPickingProductType).toBe(true);

    await act(async () => {
      dispatchAppend(dispatchRef.current, { type: "set-product-type", productType: "piles" });
    });

    const snap = asAppendState(stateRef.current);
    expect(snap?.isPickingProductType).toBe(false);
    expect(snap?.currentStep).toBe("piles");

    const skipClient = shouldSkipClientStep({
      clientName: snap?.clientName,
      appendBatches: snap?.lastDraft?.metadata.append_batches,
      resumeKpId: snap?.lastDraft?.metadata.resume_kp_id ?? null,
    });

    expect(skipClient).toBe(true);
    expect(getWizardStepOrder("piles", { skipClient })).toEqual(["piles", "result"]);
    expect(getWizardStepOrder("piles", { skipClient })).not.toContain("client");
  });

  it("«Создать новое КП» reset fully clears sticky header, draft, and picking flag", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} stateRef={stateRef} />
      </WizardDraftProvider>,
    );

    await seedStickyResultState();

    await act(async () => {
      dispatchAppend(dispatchRef.current, { type: "start-append-cycle" });
      dispatchAppend(dispatchRef.current, { type: "reset" });
    });

    const snap = asAppendState(stateRef.current);
    expect(snap?.draftId).toBeNull();
    expect(snap?.lastDraft).toBeNull();
    expect(snap?.clientName).toBe("");
    expect(snap?.discountPercent).toBe(0);
    expect(snap?.managerId).toBeNull();
    expect(snap?.conditionsMode).toBe("standard");
    expect(snap?.deliveryConditions).toBe("");
    expect(snap?.paymentConditions).toBe("");
    expect(snap?.isPickingProductType).toBe(false);
    expect(snap?.productType).toBe("plates");
    expect(snap?.currentStep).toBe("plates");
    expect(snap?.sourceText).toBe("");
    expect(screen.getByTestId("current-step")).toHaveTextContent("plates");
  });

  it("cancel-append-pick clears picking and returns to result without wiping draft", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} stateRef={stateRef} />
      </WizardDraftProvider>,
    );

    await seedStickyResultState();

    await act(async () => {
      dispatchAppend(dispatchRef.current, { type: "start-append-cycle" });
    });
    expect(asAppendState(stateRef.current)?.isPickingProductType).toBe(true);

    await act(async () => {
      dispatchAppend(dispatchRef.current, { type: "cancel-append-pick" });
    });

    const snap = asAppendState(stateRef.current);
    expect(snap?.isPickingProductType).toBe(false);
    expect(snap?.currentStep).toBe("result");
    expect(snap?.draftId).toBe("draft-append-sticky");
    expect(snap?.clientName).toBe("ООО Стикки");
    expect(snap?.discountPercent).toBe(7.5);
    expect(screen.getByTestId("current-step")).toHaveTextContent("result");
  });

  it("set-step alone does not clear isPickingProductType", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} stateRef={stateRef} />
      </WizardDraftProvider>,
    );

    await seedStickyResultState();

    await act(async () => {
      dispatchAppend(dispatchRef.current, { type: "start-append-cycle" });
      dispatchAppend(dispatchRef.current, { type: "set-step", step: "result" });
    });

    expect(asAppendState(stateRef.current)?.isPickingProductType).toBe(true);
    expect(asAppendState(stateRef.current)?.currentStep).toBe("result");
  });

  it("hydrate after undo/delete-style refresh keeps sticky discount and client (draft refresh contract)", async () => {
    render(
      <WizardDraftProvider>
        <Harness actionRef={dispatchRef} stateRef={stateRef} />
      </WizardDraftProvider>,
    );

    await seedStickyResultState();

    await act(async () => {
      dispatchAppend(dispatchRef.current, {
        type: "hydrate-draft",
        payload: makeDraft({
          draft_id: "draft-append-sticky",
          wizard_state: { ...baseWizardState, current_step: "result" },
          metadata: {
            product_type: "plates",
            client_name: "ООО Стикки",
            discount_percent: 7.5,
            manager_id: 99,
            append_batches: [],
            current_step: "result",
          },
          order_data: [],
        }),
      });
    });

    const snap = asAppendState(stateRef.current);
    expect(snap?.draftId).toBe("draft-append-sticky");
    expect(snap?.clientName).toBe("ООО Стикки");
    expect(snap?.discountPercent).toBe(7.5);
    expect(snap?.lastDraft?.metadata.append_batches).toEqual([]);
  });
});
