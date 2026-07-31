import { describe, expect, it } from "vitest";
import { buildPileLinesFromOrderData, buildPilePreviewRows } from "@/features/commercial-offer/lib/buildPilePreviewRows";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makePileDraft = (order_data: Array<Record<string, unknown>>): CommercialDraftDetails =>
  ({
    draft_id: "draft-pile-1",
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data,
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    metadata: {
      product_type: "piles",
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
      pile_batches: [],
      default_concrete_grade: "B25",
      wide_plates_resolved: true,
      last_source_filename: "",
      current_step: "piles",
      current_save_mode: null,
      execution_terms: "",
      logistics_cost: 0,
    },
    wizard_state: {
      current_step: "piles",
      can_proceed_to: ["client"],
      next_required_action: "none",
      validation_errors: [],
    },
  }) as CommercialDraftDetails;

describe("buildPilePreviewRows", () => {
  it("maps order_data to pile preview rows", () => {
    const draft = makePileDraft([
      {
        mark: "С120.35-12",
        name: "С120.35-12",
        concrete_grade: "B25",
        qty: 5,
        unit_price: 44634.03,
      },
    ]);

    expect(buildPilePreviewRows(draft)).toEqual([
      {
        mark: "С120.35-12",
        name: "С120.35-12",
        concrete_grade: "B25",
        qty: 5,
        unit_price: 44634.03,
        line_total: 223170.15,
        product_kind: "pile",
      },
    ]);
  });

  it("builds normalized lines for grade re-ingest", () => {
    const rows = buildPilePreviewRows(
      makePileDraft([
        { mark: "С120.35-12", concrete_grade: "B30_granite", qty: 2, unit_price: 100 },
        { mark: "С120.35-13и", concrete_grade: "B25", qty: 3, unit_price: 200 },
      ]),
    );

    expect(buildPileLinesFromOrderData(rows)).toBe("С120.35-12 B30_granite 2\nС120.35-13и B25 3");
  });
});
