import { describe, expect, it } from "vitest";
import {
  buildMarchLinesFromOrderData,
  buildMarchPreviewRows,
} from "@/features/commercial-offer/lib/buildMarchPreviewRows";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makeMarchDraft = (order_data: Array<Record<string, unknown>>): CommercialDraftDetails =>
  ({
    draft_id: "draft-march-1",
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data,
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    metadata: {
      product_type: "marches",
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
      march_batches: [],
      default_concrete_grade: "B25",
      wide_plates_resolved: true,
      last_source_filename: "",
      current_step: "marches",
      current_save_mode: null,
      execution_terms: "",
      logistics_cost: 0,
    },
    wizard_state: {
      current_step: "marches",
      can_proceed_to: ["client"],
      next_required_action: "none",
      validation_errors: [],
    },
  }) as CommercialDraftDetails;

describe("buildMarchPreviewRows", () => {
  it("maps order_data to march preview rows", () => {
    const draft = makeMarchDraft([
      {
        mark: "ЛМ-1",
        name: "ЛМ-1",
        concrete_grade: "B25",
        qty: 4,
        unit_price: 12000,
      },
    ]);

    expect(buildMarchPreviewRows(draft)).toEqual([
      {
        mark: "ЛМ-1",
        name: "ЛМ-1",
        concrete_grade: "B25",
        qty: 4,
        unit_price: 12000,
        line_total: 48000,
        product_kind: "march",
      },
    ]);
  });

  it("builds normalized lines for grade re-ingest", () => {
    const rows = buildMarchPreviewRows(
      makeMarchDraft([
        { mark: "ЛМ-1", concrete_grade: "B30_granite", qty: 2, unit_price: 100 },
        { mark: "ЛМ-2", concrete_grade: "B25", qty: 3, unit_price: 200 },
      ]),
    );

    expect(buildMarchLinesFromOrderData(rows)).toBe("ЛМ-1 B30_granite 2\nЛМ-2 B25 3");
  });
});
