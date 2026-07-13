import { describe, expect, it } from "vitest";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import {
  filterDraftForBatchReview,
  getCurrentBatchReviewText,
  mergeEditedBatchIntoFullText,
  needsBatchReview,
} from "@/features/commercial-offer/lib/batchReview";

const makeDraft = (plate_batches: CommercialDraftDetails["metadata"]["plate_batches"]): CommercialDraftDetails =>
  ({
    draft_id: "d1",
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data: [{ name: "ПБ 78-12-8п", qty: 2 }],
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    wizard_state: {
      current_step: "plates",
      can_proceed_to: ["client"],
      next_required_action: "none",
      validation_errors: [],
    },
    metadata: {
      source_type: "text",
      original_text: "",
      ocr_text: "",
      input_text: "batch1\nbatch2",
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
      unparsed_lines: ["batch1 bad", "batch2 line"],
      normalized_text: "batch1 parsed\nbatch2 parsed",
      normalized_lines: [],
      wide_plate_lines: [
        { id: "w1", line: "batch1 wide", qty: 1 },
        { id: "w2", line: "batch2 line", qty: 1 },
      ],
      diagnostics: [],
      price_rows_count: 0,
      breakdown_tables_count: 0,
      total_sum: 0,
      plate_batches,
      wide_plates_resolved: false,
      last_source_filename: "",
      current_step: "plates",
      current_save_mode: null,
      execution_terms: "",
      logistics_cost: 0,
    },
  }) as CommercialDraftDetails;

describe("batchReview", () => {
  it("returns only the last batch text for review", () => {
    const draft = makeDraft([
      { source_type: "text", original_text: "", normalized_text: "ПБ 78-12-8п 2", ocr_text: "", filename: "" },
      { source_type: "image", original_text: "", normalized_text: "71-12-8 3", ocr_text: "71-12-8 3", filename: "f2.jpg" },
    ]);
    expect(getCurrentBatchReviewText(draft)).toBe("71-12-8 3");
  });

  it("falls back to cumulative normalized_text when batches are empty", () => {
    const draft = makeDraft([]);
    expect(getCurrentBatchReviewText(draft)).toBe("batch1 parsed\nbatch2 parsed");
  });

  it("detects pending review when new batch exceeds confirmed count", () => {
    const draft = makeDraft([
      { source_type: "text", original_text: "", normalized_text: "a", ocr_text: "", filename: "" },
      { source_type: "text", original_text: "", normalized_text: "b", ocr_text: "", filename: "" },
    ]);
    expect(needsBatchReview(draft, 1)).toBe(true);
    expect(needsBatchReview(draft, 2)).toBe(false);
  });

  it("merges edited last batch with previous batches", () => {
    const batches = [
      { source_type: "text" as const, original_text: "", normalized_text: "ПБ 78-12-8п 2", ocr_text: "", filename: "" },
      { source_type: "image" as const, original_text: "", normalized_text: "71-12-8 3", ocr_text: "", filename: "f.jpg" },
    ];
    expect(mergeEditedBatchIntoFullText(batches, "71-12-8 5")).toBe("ПБ 78-12-8п 2\n71-12-8 5");
  });

  it("filters unparsed and wide highlights to current batch lines only", () => {
    const draft = makeDraft([
      { source_type: "text", original_text: "", normalized_text: "batch1", ocr_text: "", filename: "" },
      { source_type: "text", original_text: "", normalized_text: "batch2 line", ocr_text: "", filename: "" },
    ]);
    const filtered = filterDraftForBatchReview(draft, "batch2 line");
    expect(filtered.metadata.unparsed_lines).toEqual(["batch2 line"]);
    expect(filtered.metadata.wide_plate_lines.map((item) => item.line)).toEqual(["batch2 line"]);
  });

  it("reflects wide-plate exclude in current batch review text", () => {
    const draft = makeDraft([
      {
        source_type: "image",
        original_text: "",
        normalized_text: "ПБ 63-15-8 3\nПБ 58-15-8 10",
        ocr_text: "",
        filename: "photo.jpg",
      },
    ]);
    draft.metadata.normalized_text = "ПБ 58-15-8 10";
    draft.metadata.plate_batches![0]!.normalized_text = "ПБ 58-15-8 10";
    expect(getCurrentBatchReviewText(draft)).toBe("ПБ 58-15-8 10");
  });
});
