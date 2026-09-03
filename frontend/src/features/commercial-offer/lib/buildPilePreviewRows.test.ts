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
        product_type: "piles",
      },
    ]);

    expect(buildPilePreviewRows(draft)).toMatchObject([
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

  it("hides sealed prior plates during piles append cycle", () => {
    const draft = makePileDraft([
      {
        line_id: "ln-plate",
        product_type: "plates",
        append_batch_id: "b1",
        name: "Плиты ПБ 45-12-8п",
        qty: 14,
        unit_price: 1000,
      },
      {
        line_id: "ln-pile",
        product_type: "piles",
        mark: "С120.35-12",
        concrete_grade: "B25",
        qty: 5,
        unit_price: 44634.03,
      },
    ]);
    draft.metadata.append_batches = [{ batch_id: "b1", product_type: "plates", line_ids: ["ln-plate"] }];

    expect(buildPilePreviewRows(draft)).toMatchObject([
      {
        mark: "С120.35-12",
        name: "С120.35-12",
        concrete_grade: "B25",
        qty: 5,
        unit_price: 44634.03,
        line_total: 223170.15,
        product_kind: "pile",
        lineId: "ln-pile",
      },
    ]);
  });

  it("shows empty preview when only sealed prior (other-type) lines exist", () => {
    const draft = makePileDraft([
      {
        line_id: "ln-plate",
        product_type: "plates",
        append_batch_id: "b1",
        name: "Плиты ПБ 45-12-8п",
        qty: 14,
      },
    ]);
    expect(buildPilePreviewRows(draft)).toEqual([]);
  });

  it("keeps sealed same-type piles visible during append", () => {
    const draft = makePileDraft([
      {
        line_id: "ln-sealed",
        product_type: "piles",
        append_batch_id: "b1",
        mark: "С80.30-8",
        concrete_grade: "B25",
        qty: 2,
        unit_price: 1000,
      },
      {
        line_id: "ln-new",
        product_type: "piles",
        mark: "С120.35-12",
        concrete_grade: "B25",
        qty: 5,
        unit_price: 2000,
      },
    ]);
    draft.metadata.append_batches = [{ batch_id: "b1", product_type: "piles", line_ids: ["ln-sealed"] }];

    const rows = buildPilePreviewRows(draft);
    expect(rows).toHaveLength(2);
    expect(rows[0]).toMatchObject({ lineId: "ln-sealed", sealed: true, mark: "С80.30-8" });
    expect(rows[1]).toMatchObject({ lineId: "ln-new", sealed: false, mark: "С120.35-12" });
    expect(buildPileLinesFromOrderData(rows)).toBe("С120.35-12 B25 5");
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
