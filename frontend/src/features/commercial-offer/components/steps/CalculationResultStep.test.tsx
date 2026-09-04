import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import type { ComponentProps } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { CalculationResultStep } from "@/features/commercial-offer/components/steps/CalculationResultStep";
import type {
  CommercialDraftDetails,
  CommercialDraftMetadata,
} from "@/features/commercial-offer/types/commercialOffer";

/**
 * MNA-501 RED tests — Result step: «Тип» column, append CTA, undo/delete, trip cost gate.
 * UI + new step callbacks are intentionally not implemented yet.
 */

type ResultStepAppendHandlers = {
  onAddOtherNomenclature?: () => void;
  onUndoLastBatch?: () => Promise<void> | void;
  onDeleteLine?: (lineId: string) => Promise<void> | void;
};

const baseWizardState = {
  current_step: "result" as const,
  can_proceed_to: [] as string[],
  next_required_action: "none" as const,
  validation_errors: [] as string[],
};

const baseMetadata = (): CommercialDraftMetadata => ({
  source_type: "text",
  original_text: "",
  ocr_text: "",
  input_text: "",
  accumulated_text: "",
  manager_id: 10,
  manager_name: "Иванов",
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
  dobor_pairs: [],
  diagnostics: [],
  price_rows_count: 0,
  breakdown_tables_count: 0,
  total_sum: 0,
  plate_batches: [],
  product_type: "plates",
  wide_plates_resolved: true,
  last_source_filename: "",
  current_step: "result",
  current_save_mode: null,
  execution_terms: "",
  logistics_cost: 15000,
  append_batches: [],
  resume_kp_id: null,
});

function makeDraft(overrides: Partial<CommercialDraftDetails> = {}): CommercialDraftDetails {
  const { metadata: metaOverrides, wizard_state: wizardOverrides, ...rest } = overrides;
  return {
    draft_id: "draft-result-1",
    order: {},
    optimization: { total_plates: 1, total_cost: 1000 },
    order_data: [
      {
        line_id: "ln1",
        product_type: "plates",
        name: "ПБ 60-12-8п",
        mark: "ПБ 60-12-8п",
        qty: 2,
        unit_price: 10000,
        weight: 1500,
      },
    ],
    files: [],
    saved_offer: null,
    totals: {
      total_qty: 2,
      subtotal: 20000,
      vat_amount: 4000,
      total_with_vat: 24000,
    },
    offer_identity: { offer_number: "КП-1", offer_date: "12.08.2026", file_stem: "kp-1" },
    ...rest,
    metadata: {
      ...baseMetadata(),
      ...metaOverrides,
    },
    wizard_state: {
      ...baseWizardState,
      ...wizardOverrides,
    },
  };
}

function renderResultStep(
  draft: CommercialDraftDetails,
  handlers: ResultStepAppendHandlers = {},
  stepFlags: {
    isPileDraft?: boolean;
    isStepDraft?: boolean;
    isMarchDraft?: boolean;
    isBridgePileDraft?: boolean;
    isFbsDraft?: boolean;
    isSimpleKpDraft?: boolean;
    breakdownTables?: ComponentProps<typeof CalculationResultStep>["breakdownTables"];
    isBreakdownLoading?: boolean;
  } = {},
) {
  const onAddOtherNomenclature = handlers.onAddOtherNomenclature ?? vi.fn();
  const onUndoLastBatch = handlers.onUndoLastBatch ?? vi.fn();
  const onDeleteLine = handlers.onDeleteLine ?? vi.fn();
  const { breakdownTables = [], isBreakdownLoading = false, ...flags } = stepFlags;

  // Cast: MNA-501 will add these props to CalculationResultStep.
  const props = {
    draft,
    breakdownTables,
    isBreakdownLoading,
    errorMessage: null,
    isGeneratingFiles: false,
    isGeneratingSchema: false,
    isSaving: false,
    lastSaveResult: null,
    onBack: vi.fn(),
    onCreateNew: vi.fn(),
    onGenerateFiles: vi.fn(),
    onGenerateSchema: vi.fn(),
    onSave: vi.fn(async () => undefined),
    isUpdatingDiscount: false,
    onDiscountSubmit: vi.fn(async () => undefined),
    onLogisticsCostSubmit: vi.fn(async () => undefined),
    onAddOtherNomenclature,
    onUndoLastBatch,
    onDeleteLine,
    ...flags,
  };

  render(<CalculationResultStep {...(props as ComponentProps<typeof CalculationResultStep>)} />);

  return { onAddOtherNomenclature, onUndoLastBatch, onDeleteLine };
}

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

describe("CalculationResultStep back navigation", () => {
  it("hides Назад when onBack is not provided (no return to input without append)", () => {
    const props = {
      draft: makeDraft(),
      breakdownTables: [],
      isBreakdownLoading: false,
      errorMessage: null,
      isGeneratingFiles: false,
      isGeneratingSchema: false,
      isSaving: false,
      lastSaveResult: null,
      onCreateNew: vi.fn(),
      onGenerateFiles: vi.fn(),
      onGenerateSchema: vi.fn(),
      onSave: vi.fn(async () => undefined),
      isUpdatingDiscount: false,
      onDiscountSubmit: vi.fn(async () => undefined),
      onLogisticsCostSubmit: vi.fn(async () => undefined),
    };
    render(<CalculationResultStep {...(props as ComponentProps<typeof CalculationResultStep>)} />);
    expect(screen.queryByRole("button", { name: "Назад" })).not.toBeInTheDocument();
  });

  it("shows Назад when onBack is provided", () => {
    renderResultStep(makeDraft());
    expect(screen.getByRole("button", { name: "Назад" })).toBeInTheDocument();
  });
});

describe("CalculationResultStep MNA-501 — Тип column", () => {
  it("hides the Тип column for mono one-shot (single type, no append cycles)", () => {
    renderResultStep(
      makeDraft({
        metadata: {
          ...baseMetadata(),
          product_type: "plates",
          append_batches: [{ batch_id: "b1", product_type: "plates", line_ids: ["ln1"] }],
        },
      }),
    );

    expect(screen.queryByRole("columnheader", { name: "Тип" })).not.toBeInTheDocument();
  });

  it("shows the Тип column when order has more than one product type", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_p",
            product_type: "plates",
            name: "ПБ 60-12-8п",
            qty: 1,
            unit_price: 10000,
            weight: 800,
          },
          {
            line_id: "ln_s",
            product_type: "piles",
            name: "С80.30-8",
            mark: "С80.30-8",
            concrete_grade: "B25",
            qty: 2,
            unit_price: 5000,
          },
        ],
        metadata: {
          ...baseMetadata(),
          product_type: "piles",
          append_batches: [
            { batch_id: "b1", product_type: "plates", line_ids: ["ln_p"] },
            { batch_id: "b2", product_type: "piles", line_ids: ["ln_s"] },
          ],
        },
      }),
    );

    expect(screen.getByRole("columnheader", { name: "Тип" })).toBeInTheDocument();
  });

  it("shows the Тип column when any append cycle exists beyond the first (append_batches > 1)", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln1",
            product_type: "plates",
            append_batch_id: "b1",
            name: "ПБ 60-12-8п",
            qty: 1,
            unit_price: 10000,
            weight: 800,
          },
          {
            line_id: "ln2",
            product_type: "plates",
            append_batch_id: "b2",
            name: "ПБ 72-12-8п",
            qty: 1,
            unit_price: 12000,
            weight: 900,
          },
        ],
        metadata: {
          ...baseMetadata(),
          product_type: "plates",
          append_batches: [
            { batch_id: "b1", product_type: "plates", line_ids: ["ln1"] },
            { batch_id: "b2", product_type: "plates", line_ids: ["ln2"] },
          ],
        },
      }),
    );

    expect(screen.getByRole("columnheader", { name: "Тип" })).toBeInTheDocument();
  });

  it("renders human-readable type labels in the Тип column for mixed lines", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_p",
            product_type: "plates",
            name: "ПБ 60-12-8п",
            qty: 1,
            unit_price: 10000,
            weight: 800,
          },
          {
            line_id: "ln_s",
            product_type: "piles",
            name: "С80.30-8",
            mark: "С80.30-8",
            concrete_grade: "B25",
            qty: 2,
            unit_price: 5000,
          },
        ],
        metadata: {
          ...baseMetadata(),
          append_batches: [
            { batch_id: "b1", product_type: "plates", line_ids: ["ln_p"] },
            { batch_id: "b2", product_type: "piles", line_ids: ["ln_s"] },
          ],
        },
      }),
    );

    const positionsCard = screen.getByRole("heading", { name: "Позиции" }).closest("section");
    expect(positionsCard).not.toBeNull();
    expect(within(positionsCard as HTMLElement).getByText("Плиты")).toBeInTheDocument();
    expect(within(positionsCard as HTMLElement).getByText("Сваи")).toBeInTheDocument();
  });
});

describe("CalculationResultStep MNA-501 — append CTA", () => {
  it("renders CTA «Добавить другое наименование»", () => {
    renderResultStep(makeDraft());

    expect(
      screen.getByRole("button", { name: "Добавить другое наименование" }),
    ).toBeInTheDocument();
  });

  it("calls onAddOtherNomenclature when CTA is clicked", () => {
    const onAddOtherNomenclature = vi.fn();
    renderResultStep(makeDraft(), { onAddOtherNomenclature });

    fireEvent.click(screen.getByRole("button", { name: "Добавить другое наименование" }));

    expect(onAddOtherNomenclature).toHaveBeenCalledOnce();
  });
});

describe("CalculationResultStep MNA-501 — undo last batch / delete line", () => {
  it("shows «Отменить последний заход» when append_batches is non-empty", () => {
    renderResultStep(
      makeDraft({
        metadata: {
          ...baseMetadata(),
          append_batches: [{ batch_id: "b1", product_type: "plates", line_ids: ["ln1"] }],
        },
      }),
    );

    expect(screen.getByRole("button", { name: "Отменить последний заход" })).toBeInTheDocument();
  });

  it("hides «Отменить последний заход» when there are no sealed batches", () => {
    renderResultStep(
      makeDraft({
        metadata: {
          ...baseMetadata(),
          append_batches: [],
        },
      }),
    );

    expect(screen.queryByRole("button", { name: "Отменить последний заход" })).not.toBeInTheDocument();
  });

  it("calls onUndoLastBatch when undo is clicked", async () => {
    const onUndoLastBatch = vi.fn(async () => undefined);
    renderResultStep(
      makeDraft({
        metadata: {
          ...baseMetadata(),
          append_batches: [
            { batch_id: "b1", product_type: "plates", line_ids: ["ln1"] },
            { batch_id: "b2", product_type: "piles", line_ids: ["ln2"] },
          ],
        },
      }),
      { onUndoLastBatch },
    );

    fireEvent.click(screen.getByRole("button", { name: "Отменить последний заход" }));

    expect(onUndoLastBatch).toHaveBeenCalledOnce();
  });

  it("shows a delete control per line with line_id", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_a",
            product_type: "plates",
            name: "ПБ 60-12-8п",
            qty: 1,
            unit_price: 10000,
            weight: 800,
          },
          {
            line_id: "ln_b",
            product_type: "plates",
            name: "ПБ 72-12-8п",
            qty: 1,
            unit_price: 12000,
            weight: 900,
          },
        ],
      }),
    );

    expect(screen.getByRole("button", { name: "Удалить строку ln_a" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Удалить строку ln_b" })).toBeInTheDocument();
  });

  it("calls onDeleteLine with line_id when delete is clicked", () => {
    const onDeleteLine = vi.fn(async () => undefined);
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_del",
            product_type: "plates",
            name: "ПБ 60-12-8п",
            qty: 1,
            unit_price: 10000,
            weight: 800,
          },
        ],
      }),
      { onDeleteLine },
    );

    fireEvent.click(screen.getByRole("button", { name: "Удалить строку ln_del" }));

    expect(onDeleteLine).toHaveBeenCalledExactlyOnceWith("ln_del");
  });

  it("does not show a button labelled Удалить and shows an edit icon", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_icons",
            product_type: "plates",
            name: "ПБ 60-12-8п",
            qty: 1,
            unit_price: 10000,
            weight: 800,
          },
        ],
      }),
    );

    expect(screen.queryByRole("button", { name: "Удалить" })).not.toBeInTheDocument();
    expect(screen.queryByText("Удалить")).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Изменить строку ln_icons" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Удалить строку ln_icons" })).toBeInTheDocument();
  });
});

describe("CalculationResultStep MNA-501 — trip cost gate", () => {
  it("enables trip cost for pile-only offers (writes pile_logistics_cost)", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_s",
            product_type: "piles",
            name: "С80.30-8",
            mark: "С80.30-8",
            concrete_grade: "B25",
            qty: 4,
            unit_price: 5000,
          },
        ],
        metadata: {
          ...baseMetadata(),
          product_type: "piles",
          logistics_cost: 0,
        },
      }),
      {},
      { isPileDraft: true, isSimpleKpDraft: true },
    );

    const tripInput = screen.getByPlaceholderText("Стоимость одного рейса");
    expect(tripInput).not.toBeDisabled();
  });

    it("enables trip cost when there is at least one plate line", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_p",
            product_type: "plates",
            name: "ПБ 60-12-8п",
            qty: 1,
            unit_price: 10000,
            weight: 800,
          },
          {
            line_id: "ln_s",
            product_type: "piles",
            name: "С80.30-8",
            mark: "С80.30-8",
            concrete_grade: "B25",
            qty: 2,
            unit_price: 5000,
          },
        ],
        metadata: {
          ...baseMetadata(),
          append_batches: [
            { batch_id: "b1", product_type: "plates", line_ids: ["ln_p"] },
            { batch_id: "b2", product_type: "piles", line_ids: ["ln_s"] },
          ],
        },
      }),
    );

    const plateField = screen.getByText("Рейс плит").closest("label");
    expect(plateField).not.toBeNull();
    expect(within(plateField as HTMLElement).getByPlaceholderText("Стоимость одного рейса")).not.toBeDisabled();
  });

  it("shows two trip fields for mixed plates+piles", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_p",
            product_type: "plates",
            name: "ПБ 60-12-8п",
            qty: 1,
            unit_price: 10000,
            weight: 800,
          },
          {
            line_id: "ln_s",
            product_type: "piles",
            name: "С80.30-8",
            mark: "С80.30-8",
            concrete_grade: "B25",
            qty: 2,
            unit_price: 5000,
          },
        ],
        metadata: {
          ...baseMetadata(),
          append_batches: [
            { batch_id: "b1", product_type: "plates", line_ids: ["ln_p"] },
            { batch_id: "b2", product_type: "piles", line_ids: ["ln_s"] },
          ],
        },
      }),
    );

    expect(screen.getByText("Рейс плит")).toBeInTheDocument();
    expect(screen.getByText("Рейс свай")).toBeInTheDocument();
  });

  it("asks how many trucks for pending marks and shows trip total only", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_c18",
            product_type: "bridge_piles",
            name: "C18-40T8",
            mark: "C18-40T8",
            qty: 49,
            unit_price: 1000,
          },
        ],
        metadata: {
          ...baseMetadata(),
          product_type: "bridge_piles",
        },
        totals: {
          total_qty: 49,
          subtotal: 0,
          vat_amount: 0,
          total_with_vat: 0,
          pile_delivery_ready: false,
          pile_trips: 0,
          pile_trip_pending_marks: ["C18-40T8"],
        },
      }),
      {},
      { isBridgePileDraft: true, isSimpleKpDraft: true },
    );

    expect(
      screen.getByText(/Для C18-40T8 \(49 шт\.\) нет нормы загрузки в справочнике/),
    ).toBeInTheDocument();
    expect(screen.getByPlaceholderText("Машин, шт.")).toBeInTheDocument();
    expect(screen.queryByText(/полных/)).not.toBeInTheDocument();
    expect(screen.queryByText("Рейсов свай")).not.toBeInTheDocument();
  });
});

describe("CalculationResultStep unparsed UX", () => {
  it("does not add a dead unparsed-lines warning and still shows other warnings", () => {
    renderResultStep(
      makeDraft({
        metadata: {
          ...baseMetadata(),
          unparsed_lines: ["xyz-not-a-plate (пропущено: не совпал формат строки)"],
          warnings: [
            "Не удалось распознать строк: 1",
            "Строки формата «длина×ширина×толщина» (мм), например «3880x1200x220»: нагрузка принята 8п по умолчанию. Проверьте нагрузку перед отправкой КП.",
          ],
        },
      }),
    );

    expect(screen.queryByText(/Строки, не попавшие в расчёт/)).not.toBeInTheDocument();
    expect(screen.queryByText(/Не удалось распознать строк: 1/)).not.toBeInTheDocument();
    expect(screen.getByText(/нагрузка принята 8п по умолчанию/)).toBeInTheDocument();
  });
});

describe("CalculationResultStep breakdown availability after invalidate", () => {
  it("keeps plate breakdown link when draft count is 0 but query tables are present", () => {
    renderResultStep(
      makeDraft({
        metadata: {
          ...baseMetadata(),
          breakdown_tables_count: 0,
        },
        order_data: [
          {
            line_id: "ln1",
            product_type: "plates",
            name: "ПБ 60-12-8п",
            mark: "ПБ 60-12-8п",
            qty: 2,
            unit_price: 10000,
            weight: 1500,
          },
        ],
      }),
      {},
      {
        breakdownTables: [
          {
            name: "ПБ 60-12-8п",
            rows: [["Итого", "", "10000"]],
          },
        ],
      },
    );

    expect(screen.getByTitle("Показать детальную разбивку цены")).toBeInTheDocument();
  });

  it("does not show breakdown link when count is 0 and query tables are empty", () => {
    renderResultStep(
      makeDraft({
        metadata: {
          ...baseMetadata(),
          breakdown_tables_count: 0,
        },
      }),
    );

    expect(screen.queryByTitle("Показать детальную разбивку цены")).not.toBeInTheDocument();
    expect(screen.getByText("ПБ 60-12-8п")).toBeInTheDocument();
  });
});

describe("CalculationResultStep tracks estimate", () => {
  it("shows ~N дорожек for plate KP from length_m × qty", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln1",
            product_type: "plates",
            name: "ПБ 60-12-8п",
            qty: 20,
            unit_price: 10000,
            weight: 1500,
            length_m: 6.0,
          },
        ],
      }),
    );

    const estimate = screen.getByTestId("result-tracks-estimate");
    // 20 × 6.0 м = 120 м → ceil(120/101) = 2
    expect(estimate).toHaveTextContent("~2 дорожек");
  });

  it("counts only plate lines in a mixed offer", () => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_p",
            product_type: "plates",
            name: "ПБ 60-12-8п",
            qty: 20,
            unit_price: 10000,
            weight: 800,
            length_m: 6.0,
          },
          {
            line_id: "ln_s",
            product_type: "piles",
            name: "С80.30-8",
            mark: "С80.30-8",
            qty: 50,
            unit_price: 5000,
            length_m: 8.0,
          },
        ],
      }),
    );

    expect(screen.getByTestId("result-tracks-estimate")).toHaveTextContent("~2 дорожек");
  });

  it.each([
    {
      flag: "isPileDraft" as const,
      productType: "piles",
      name: "С80.30-8",
    },
    {
      flag: "isStepDraft" as const,
      productType: "steps",
      name: "ЛС-12",
    },
    {
      flag: "isFbsDraft" as const,
      productType: "fbs",
      name: "ФБС 24-3-6",
    },
  ])("hides tracks estimate for $productType", ({ flag, productType, name }) => {
    renderResultStep(
      makeDraft({
        order_data: [
          {
            line_id: "ln_simple",
            product_type: productType,
            name,
            mark: name,
            qty: 10,
            unit_price: 1000,
            length_m: 8.0,
          },
        ],
        metadata: {
          ...baseMetadata(),
          product_type: productType as CommercialDraftMetadata["product_type"],
        },
      }),
      {},
      { [flag]: true, isSimpleKpDraft: true },
    );

    expect(screen.queryByTestId("result-tracks-estimate")).not.toBeInTheDocument();
    expect(screen.queryByText(/дорожек/)).not.toBeInTheDocument();
  });
});
