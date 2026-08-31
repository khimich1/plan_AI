import { describe, expect, it } from "vitest";

import {
  buildFbsLinesFromOrderData,
  buildFbsPreviewRows,
} from "@/features/commercial-offer/lib/buildFbsPreviewRows";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const baseDraft = (): CommercialDraftDetails =>
  ({
    draft_id: "d1",
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data: [
      {
        mark: "ФБС 9.3.6-Т",
        name: "ФБС 9.3.6-Т",
        concrete_grade: "B25",
        available_grades: ["B7_5", "B20", "B22_5", "B25"],
        qty: 2,
        unit_price: 1788.33,
        product_kind: "fbs",
      },
    ],
    metadata: {
      product_type: "fbs",
      source_type: "text",
      original_text: "",
      ocr_text: "",
      input_text: "ФБС 9.3.6-Т 2",
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
      dobor_pairs: [],
      diagnostics: [],
      price_rows_count: 1,
      breakdown_tables_count: 0,
      total_sum: 0,
      plate_batches: [],
      default_concrete_grade: "B25",
      wide_plates_resolved: true,
      last_source_filename: "",
      current_step: "fbs",
      current_save_mode: null,
      execution_terms: "",
      logistics_cost: 0,
    },
    wizard_state: {
      current_step: "fbs",
      can_proceed_to: ["client"],
      next_required_action: "none",
      validation_errors: [],
    },
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
  }) as CommercialDraftDetails;

describe("buildFbsPreviewRows", () => {
  it("maps order_data rows", () => {
    const rows = buildFbsPreviewRows(baseDraft());
    expect(rows).toHaveLength(1);
    expect(rows[0].mark).toBe("ФБС 9.3.6-Т");
    expect(rows[0].product_kind).toBe("fbs");
  });

  it("buildFbsLinesFromOrderData joins mark grade qty", () => {
    const text = buildFbsLinesFromOrderData([
      {
        mark: "ФБС 9.3.6-Т",
        concrete_grade: "B25",
        qty: 2,
        unit_price: 100,
        name: "ФБС 9.3.6-Т",
        product_kind: "fbs",
      },
      {
        mark: "ФБС 12.4.6-Т",
        concrete_grade: "B7_5",
        qty: 3,
        unit_price: 200,
        name: "ФБС 12.4.6-Т",
        product_kind: "fbs",
      },
    ]);
    expect(text).toBe("ФБС 9.3.6-Т B25 2\nФБС 12.4.6-Т B7_5 3");
  });
});
