import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { KpGradedPreviewPanel } from "@/features/commercial-offer/components/KpGradedPreviewPanel";
import type { CommercialDraftDetails, ProductType } from "@/features/commercial-offer/types/commercialOffer";

afterEach(() => {
  cleanup();
});

const basePileDraft = (orderData: CommercialDraftDetails["order_data"]): CommercialDraftDetails => ({
  draft_id: "d1",
  order: {},
  optimization: { total_plates: 0, total_cost: 0 },
  order_data: orderData,
  files: [],
  saved_offer: null,
  totals: {},
  offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
  wizard_state: {
    current_step: "piles",
    can_proceed_to: [],
    next_required_action: "none",
    validation_errors: [],
  },
  metadata: {
    source_type: "text",
    original_text: "",
    ocr_text: "",
    input_text: "",
    accumulated_text: "",
    manager_id: null,
    manager_name: "",
    manager_phone: "",
    manager_email: "",
    client_name: "",
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
    product_type: "piles",
    wide_plates_resolved: true,
    last_source_filename: "",
    current_step: "piles",
    current_save_mode: null,
    execution_terms: "",
    logistics_cost: 0,
    default_concrete_grade: "B25",
  },
});

describe("KpGradedPreviewPanel bulk grade with sealed rows", () => {
  it("hides bulk grade control when all visible rows are sealed", () => {
    render(
      <KpGradedPreviewPanel
        productType="piles"
        draft={basePileDraft([
          {
            line_id: "ln1",
            product_type: "piles",
            append_batch_id: "b1",
            mark: "С60.30",
            name: "Свая С60.30",
            qty: 2,
            unit_price: 1000,
            concrete_grade: "B25",
          },
        ])}
        normalizedText=""
        onApplyGradeToAll={vi.fn()}
      />,
    );

    expect(screen.queryByText(/Применить класс ко всем/)).not.toBeInTheDocument();
  });

  it("shows «ко всем новым» when there is at least one unsealed row", () => {
    render(
      <KpGradedPreviewPanel
        productType="piles"
        draft={basePileDraft([
          {
            line_id: "ln1",
            product_type: "piles",
            append_batch_id: "b1",
            mark: "С60.30",
            name: "Свая С60.30",
            qty: 2,
            unit_price: 1000,
            concrete_grade: "B25",
          },
          {
            line_id: "ln2",
            product_type: "piles",
            mark: "С80.30",
            name: "Свая С80.30",
            qty: 1,
            unit_price: 2000,
            concrete_grade: "B25",
          },
        ])}
        normalizedText=""
        onApplyGradeToAll={vi.fn()}
      />,
    );

    expect(screen.getByText("Применить класс ко всем новым:")).toBeInTheDocument();
  });
});

const makeStepDraft = (): CommercialDraftDetails =>
  ({
    draft_id: "draft-step-1",
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data: [
      {
        mark: "ЛС11",
        name: "ЛС11",
        qty: 10,
        unit_price: 12000,
      },
    ],
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    metadata: {
      product_type: "steps",
      source_type: "text",
      original_text: "",
      ocr_text: "",
      input_text: "",
      accumulated_text: "",
      manager_id: null,
      manager_name: "",
      manager_phone: "",
      manager_email: "",
      client_name: "",
      discount_percent: 0,
      conditions_mode: "standard",
      delivery_conditions: "",
      payment_conditions: "",
      warnings: [],
      unparsed_lines: [],
      normalized_text: "ЛС11 10",
      normalized_lines: [],
      wide_plate_lines: [],
      diagnostics: [],
      price_rows_count: 0,
      breakdown_tables_count: 0,
      total_sum: 0,
      plate_batches: [],
      step_batches: [],
      wide_plates_resolved: true,
      last_source_filename: "",
      current_step: "steps",
      current_save_mode: null,
      execution_terms: "",
      logistics_cost: 0,
    },
    wizard_state: {
      current_step: "steps",
      can_proceed_to: ["client"],
      next_required_action: "none",
      validation_errors: [],
    },
  }) as CommercialDraftDetails;

describe("KpGradedPreviewPanel steps (no grades)", () => {
  it("renders mark, qty, price, and sum columns without concrete grade", () => {
    render(<KpGradedPreviewPanel productType="steps" draft={makeStepDraft()} normalizedText="ЛС11 10" />);

    expect(screen.getByText("Марка")).toBeInTheDocument();
    expect(screen.getByText("Кол-во")).toBeInTheDocument();
    expect(screen.getByText("Цена")).toBeInTheDocument();
    expect(screen.getByText("Сумма")).toBeInTheDocument();
    expect(screen.queryByText("Класс")).not.toBeInTheDocument();
    expect(screen.getByText("ЛС11")).toBeInTheDocument();
  });

  it("does not show bulk grade control even with grade callbacks passed", () => {
    render(
      <KpGradedPreviewPanel
        productType="steps"
        draft={makeStepDraft()}
        normalizedText="ЛС11 10"
        onApplyGradeToAll={vi.fn()}
        onLineGradeChange={vi.fn()}
      />,
    );

    expect(screen.queryByText(/Применить класс ко всем/)).not.toBeInTheDocument();
    expect(screen.queryByRole("combobox")).not.toBeInTheDocument();
  });

  it("does not show unparsed composition list or count banner", () => {
    const draft = makeStepDraft();
    draft.metadata.unparsed_lines = ["плохо"];
    draft.metadata.warnings = ["Не удалось распознать строк: 1"];
    render(<KpGradedPreviewPanel productType="steps" draft={draft} normalizedText="ЛС11 10" />);

    expect(screen.queryByText("Не попали в состав")).not.toBeInTheDocument();
    expect(screen.queryByText(/Не удалось распознать строк: 1/)).not.toBeInTheDocument();
  });
});

const makeUnparsedDraft = (productType: ProductType, name: string): CommercialDraftDetails =>
  ({
    draft_id: `draft-${productType}`,
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data: [{ mark: name, name, qty: 1, unit_price: 1000, concrete_grade: "B25" }],
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    metadata: {
      product_type: productType,
      source_type: "text",
      original_text: "",
      ocr_text: "",
      input_text: "",
      accumulated_text: "",
      manager_id: null,
      manager_name: "",
      manager_phone: "",
      manager_email: "",
      client_name: "",
      discount_percent: 0,
      conditions_mode: "standard",
      delivery_conditions: "",
      payment_conditions: "",
      warnings: ["Не удалось распознать строк: 1"],
      unparsed_lines: ["плохо"],
      normalized_text: name,
      normalized_lines: [],
      wide_plate_lines: [],
      diagnostics: [],
      price_rows_count: 0,
      breakdown_tables_count: 0,
      total_sum: 0,
      plate_batches: [],
      wide_plates_resolved: true,
      last_source_filename: "",
      current_step: productType,
      current_save_mode: null,
      execution_terms: "",
      logistics_cost: 0,
    },
    wizard_state: {
      current_step: productType,
      can_proceed_to: ["client"],
      next_required_action: "none",
      validation_errors: [],
    },
  }) as CommercialDraftDetails;

describe("KpGradedPreviewPanel unparsed UX", () => {
  it.each([
    ["piles", "С120.35-12"],
    ["fbs", "ФБС 9.3.6-Т"],
    ["marches", "1ЛМ 27-11-14-4"],
    ["bridge_piles", "С7-35Т5"],
  ] as const)("%s hides unparsed list and count banner", (productType, name) => {
    render(
      <KpGradedPreviewPanel
        productType={productType}
        draft={makeUnparsedDraft(productType, name)}
        normalizedText={name}
      />,
    );
    expect(screen.queryByText("Не попали в состав")).not.toBeInTheDocument();
    expect(screen.queryByText(/Не удалось распознать строк: 1/)).not.toBeInTheDocument();
  });
});

describe("KpGradedPreviewPanel per-row available grades", () => {
  it("limits the row grade select to row.available_grades when provided (fbs)", () => {
    const draft = basePileDraft([
      {
        line_id: "ln1",
        product_type: "fbs",
        mark: "ФБС 9.3.6-Т",
        name: "ФБС 9.3.6-Т",
        qty: 2,
        unit_price: 1000,
        concrete_grade: "B20",
        available_grades: ["B20", "B22_5"],
      },
    ]);
    draft.metadata.product_type = "fbs";
    draft.wizard_state.current_step = "fbs";
    draft.metadata.current_step = "fbs";
    render(
      <KpGradedPreviewPanel
        productType="fbs"
        draft={draft}
        normalizedText=""
        onLineGradeChange={vi.fn()}
      />,
    );

    const select = screen.getByRole("combobox");
    const options = Array.from(select.querySelectorAll("option")).map((option) => option.getAttribute("value"));
    expect(options).toEqual(["B20", "B22_5"]);
  });

  it("falls back to the full grade catalog when the row has no available_grades (fbs)", () => {
    const draft = basePileDraft([
      {
        line_id: "ln1",
        product_type: "fbs",
        mark: "ФБС 9.3.6-Т",
        name: "ФБС 9.3.6-Т",
        qty: 2,
        unit_price: 1000,
        concrete_grade: "B25",
      },
    ]);
    draft.metadata.product_type = "fbs";
    draft.wizard_state.current_step = "fbs";
    draft.metadata.current_step = "fbs";
    render(
      <KpGradedPreviewPanel
        productType="fbs"
        draft={draft}
        normalizedText=""
        onLineGradeChange={vi.fn()}
      />,
    );

    const select = screen.getByRole("combobox");
    const options = Array.from(select.querySelectorAll("option")).map((option) => option.getAttribute("value"));
    expect(options).toEqual(["B7_5", "B20", "B22_5", "B25"]);
  });
});
