import { render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { KpPlatePreviewPanel } from "@/features/commercial-offer/components/KpPlatePreviewPanel";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makeDraft = (overrides: Partial<CommercialDraftDetails["metadata"]> = {}): CommercialDraftDetails =>
  ({
    draft_id: "draft-plate-preview",
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data: [{ name: "ПБ 78-12-8п", qty: 2, unit_price: 1000 }],
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    metadata: {
      product_type: "plates",
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
      warnings: [
        "Не удалось распознать строк: 1",
        "Строки формата «длина×ширина×толщина» (мм), например «3880x1200x220»: нагрузка принята 8п по умолчанию. Проверьте нагрузку перед отправкой КП.",
      ],
      unparsed_lines: ["xyz-not-a-plate (пропущено: не совпал формат строки)"],
      normalized_text: "ПБ 78-12-8п 2",
      normalized_lines: [],
      wide_plate_lines: [],
      dobor_pairs: [],
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
      ...overrides,
    },
    wizard_state: {
      current_step: "plates",
      can_proceed_to: ["client"],
      next_required_action: "none",
      validation_errors: [],
    },
  }) as CommercialDraftDetails;

describe("KpPlatePreviewPanel unparsed UX", () => {
  it("hides unparsed list and count banner but keeps LWH load warning", () => {
    render(<KpPlatePreviewPanel draft={makeDraft()} normalizedText="ПБ 78-12-8п 2" />);

    expect(screen.queryByText("Не попали в состав")).not.toBeInTheDocument();
    expect(screen.queryByText(/Не удалось распознать строк: 1/)).not.toBeInTheDocument();
    expect(screen.getByText(/нагрузка принята 8п по умолчанию/)).toBeInTheDocument();
  });
});

describe("KpPlatePreviewPanel row icons", () => {
  it("shows edit and delete icons when line_id is present and handlers are passed", () => {
    const handlers = {
      onSaveLine: vi.fn(),
      onDeleteLine: vi.fn(),
      undoToast: null,
    };
    const draft = makeDraft();
    draft.order_data = [{ line_id: "ln_p1", name: "ПБ 78-12-8п", qty: 2, unit_price: 1000 }];
    render(<KpPlatePreviewPanel draft={draft} normalizedText="ПБ 78-12-8п 2" lineRowHandlers={handlers} />);
    expect(screen.getByRole("button", { name: "Изменить строку ln_p1" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Удалить строку ln_p1" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Изменить строку ln_p1" })).not.toBeDisabled();
  });

  it("keeps wide-plate alert and does not disable the pencil", () => {
    const handlers = {
      onSaveLine: vi.fn(),
      onDeleteLine: vi.fn(),
      undoToast: null,
    };
    const draft = makeDraft({
      wide_plate_lines: [{ id: "wide-1", line: "ПБ 59-15-8п 2", qty: 2 }],
      wide_plates_resolved: false,
    });
    draft.order_data = [
      {
        line_id: "ln_wide",
        name: "ПБ 59-15-8п",
        qty: 2,
        unit_price: 15000,
        length_m: 5.9,
        width_m: 1.5,
      },
    ];
    render(<KpPlatePreviewPanel draft={draft} normalizedText="ПБ 59-15-8п 2" lineRowHandlers={handlers} />);
    expect(screen.getByText(/позиция шире стандартной/)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Изменить строку ln_wide" })).not.toBeDisabled();
  });
});
