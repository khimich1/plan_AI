import { describe, expect, it } from "vitest";
import {
  getCurrentCycleOrderData,
  getProductTypeOrderData,
  isSealedOrderLine,
} from "@/features/commercial-offer/lib/currentCycleOrderData";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makeDraft = (order_data: Array<Record<string, unknown>>, product_type = "piles"): CommercialDraftDetails =>
  ({
    draft_id: "d1",
    order: {},
    optimization: {},
    order_data,
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    metadata: {
      product_type,
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
      append_batches: [{ batch_id: "b1", product_type: "plates", line_ids: ["ln-plate"] }],
    },
    wizard_state: {
      current_step: "piles",
      can_proceed_to: ["result"],
      next_required_action: "none",
      validation_errors: [],
    },
  }) as CommercialDraftDetails;

describe("currentCycleOrderData", () => {
  it("detects sealed lines by append_batch_id", () => {
    expect(isSealedOrderLine({ append_batch_id: "b1" })).toBe(true);
    expect(isSealedOrderLine({ append_batch_id: "" })).toBe(false);
    expect(isSealedOrderLine({})).toBe(false);
  });

  it("hides sealed prior nomenclature during a new piles cycle", () => {
    const draft = makeDraft([
      {
        line_id: "ln-plate",
        product_type: "plates",
        append_batch_id: "b1",
        name: "Плиты ПБ 45-12-8п",
        qty: 10,
      },
      {
        line_id: "ln-pile",
        product_type: "piles",
        mark: "С120.35-12",
        concrete_grade: "B25",
        qty: 5,
      },
    ]);

    const cycle = getCurrentCycleOrderData(draft, "piles");
    expect(cycle).toHaveLength(1);
    expect(cycle[0]?.line_id).toBe("ln-pile");
  });

  it("returns empty when append cycle has no new lines yet", () => {
    const draft = makeDraft([
      {
        line_id: "ln-plate",
        product_type: "plates",
        append_batch_id: "b1",
        name: "Плиты ПБ 45-12-8п",
        qty: 10,
      },
    ]);

    expect(getCurrentCycleOrderData(draft, "piles")).toEqual([]);
  });

  it("keeps unsealed mono lines for first cycle", () => {
    const draft = makeDraft(
      [{ line_id: "ln1", product_type: "piles", mark: "С120.35-12", qty: 2 }],
      "piles",
    );
    draft.metadata.append_batches = [];

    expect(getCurrentCycleOrderData(draft, "piles")).toHaveLength(1);
  });
});

describe("getProductTypeOrderData", () => {
  it("includes sealed same-type lines for preview (append same type)", () => {
    const draft = makeDraft(
      [
        {
          line_id: "ln-sealed",
          product_type: "plates",
          append_batch_id: "b1",
          name: "Плиты ПБ 28-5,3-8п",
          qty: 2,
        },
        {
          line_id: "ln-new",
          product_type: "plates",
          name: "Плиты ПБ 51-5,3-8п",
          qty: 1,
        },
      ],
      "plates",
    );

    const rows = getProductTypeOrderData(draft, "plates");
    expect(rows).toHaveLength(2);
    expect(rows.map((r) => r.line_id)).toEqual(["ln-sealed", "ln-new"]);
  });

  it("shows sealed same-type even when current cycle has no new lines yet", () => {
    const draft = makeDraft(
      [
        {
          line_id: "ln-sealed",
          product_type: "plates",
          append_batch_id: "b1",
          name: "Плиты ПБ 28-5,3-8п",
          qty: 2,
        },
      ],
      "plates",
    );

    expect(getProductTypeOrderData(draft, "plates")).toHaveLength(1);
    expect(getCurrentCycleOrderData(draft, "plates")).toEqual([]);
  });

  it("hides sealed prior plates during a piles cycle", () => {
    const draft = makeDraft([
      {
        line_id: "ln-plate",
        product_type: "plates",
        append_batch_id: "b1",
        name: "Плиты ПБ 45-12-8п",
        qty: 10,
      },
      {
        line_id: "ln-pile",
        product_type: "piles",
        mark: "С120.35-12",
        qty: 5,
      },
    ]);

    const piles = getProductTypeOrderData(draft, "piles");
    expect(piles).toHaveLength(1);
    expect(piles[0]?.line_id).toBe("ln-pile");
  });
});
