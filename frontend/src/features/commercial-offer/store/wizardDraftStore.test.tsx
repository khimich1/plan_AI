import { act, cleanup, render, screen } from "@testing-library/react";
import type { MutableRefObject } from "react";
import { afterEach, describe, expect, it, vi, beforeEach } from "vitest";
import type { CommercialDraftDetails, CommercialWizardState, WizardStepId } from "@/features/commercial-offer/types/commercialOffer";
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
