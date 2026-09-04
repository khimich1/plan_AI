import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { KpPilePreviewPanel } from "@/features/commercial-offer/components/KpPilePreviewPanel";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

afterEach(() => {
  cleanup();
});

const baseDraft = (orderData: CommercialDraftDetails["order_data"]): CommercialDraftDetails => ({
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

describe("KpPilePreviewPanel bulk grade with sealed rows", () => {
  it("hides bulk grade control when all visible rows are sealed", () => {
    render(
      <KpPilePreviewPanel
        draft={baseDraft([
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
      <KpPilePreviewPanel
        draft={baseDraft([
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
