import { describe, expect, it } from "vitest";
import { buildStepLinesFromOrderData, buildStepPreviewRows } from "@/features/commercial-offer/lib/buildStepPreviewRows";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makeStepDraft = (order_data: Array<Record<string, unknown>>): CommercialDraftDetails =>
  ({
    draft_id: "draft-step-1",
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data,
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
      normalized_text: "",
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

describe("buildStepPreviewRows", () => {
  it("maps order_data to step preview rows without concrete grade", () => {
    const draft = makeStepDraft([
      {
        mark: "ЛС11",
        name: "ЛС11",
        qty: 10,
        unit_price: 12000,
      },
    ]);

    expect(buildStepPreviewRows(draft)).toMatchObject([
      {
        mark: "ЛС11",
        name: "ЛС11",
        qty: 10,
        unit_price: 12000,
        line_total: 120000,
        product_kind: "step",
      },
    ]);
  });

  it("builds normalized lines for step re-ingest", () => {
    const rows = buildStepPreviewRows(
      makeStepDraft([
        { mark: "ЛС11", qty: 10, unit_price: 100 },
        { mark: "ЛС14-1лев", qty: 5, unit_price: 200 },
      ]),
    );

    expect(buildStepLinesFromOrderData(rows)).toBe("ЛС11 10\nЛС14-1лев 5");
  });
});
