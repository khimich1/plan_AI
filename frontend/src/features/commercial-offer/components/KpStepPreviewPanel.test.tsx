import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { KpStepPreviewPanel } from "@/features/commercial-offer/components/KpStepPreviewPanel";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makeDraft = (): CommercialDraftDetails =>
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

describe("KpStepPreviewPanel", () => {
  it("renders mark, qty, price, and sum columns without concrete grade", () => {
    render(<KpStepPreviewPanel draft={makeDraft()} normalizedText="ЛС11 10" />);

    expect(screen.getByText("Марка")).toBeInTheDocument();
    expect(screen.getByText("Кол-во")).toBeInTheDocument();
    expect(screen.getByText("Цена")).toBeInTheDocument();
    expect(screen.getByText("Сумма")).toBeInTheDocument();
    expect(screen.queryByText("Класс")).not.toBeInTheDocument();
    expect(screen.getByText("ЛС11")).toBeInTheDocument();
  });
});
