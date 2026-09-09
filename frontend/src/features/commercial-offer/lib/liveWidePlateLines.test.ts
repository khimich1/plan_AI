import { describe, expect, it } from "vitest";

import {
  liveWidePlateLines,
  overlayDraftWithLiveWideLines,
  widePlateMatchKey,
} from "@/features/commercial-offer/lib/liveWidePlateLines";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makeDraft = (
  wide_plate_lines: CommercialDraftDetails["metadata"]["wide_plate_lines"],
  wide_plates_resolved = false,
): CommercialDraftDetails =>
  ({
    draft_id: "d1",
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data: [],
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    wizard_state: {
      current_step: "plates",
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
      wide_plate_lines,
      diagnostics: [],
      price_rows_count: 0,
      breakdown_tables_count: 0,
      total_sum: 0,
      plate_batches: [],
      wide_plates_resolved,
      last_source_filename: "",
      current_step: "plates",
      current_save_mode: null,
      execution_terms: "",
      logistics_cost: 0,
    },
  }) as CommercialDraftDetails;

describe("liveWidePlateLines", () => {
  it("treats 44-15-10п 5 as wide with qty 5", () => {
    const lines = liveWidePlateLines("44-15-10п 5");
    expect(lines).toHaveLength(1);
    expect(lines[0]?.qty).toBe(5);
    expect(lines[0]?.line).toBe("44-15-10п 5");
    expect(lines[0]?.id).toBe("live-wide-0");
  });

  it("after replacing text with 34-15-10п 15 reports wide qty 15", () => {
    expect(liveWidePlateLines("44-15-10п 5")[0]?.qty).toBe(5);
    const next = liveWidePlateLines("34-15-10п 15");
    expect(next).toHaveLength(1);
    expect(next[0]?.qty).toBe(15);
    expect(next[0]?.line).toBe("34-15-10п 15");
  });

  it("does not treat 34-12-10п 15 as wide", () => {
    expect(liveWidePlateLines("34-12-10п 15")).toEqual([]);
  });

  it("treats ПБ 34-15-10п 15 as wide", () => {
    const lines = liveWidePlateLines("ПБ 34-15-10п 15");
    expect(lines).toHaveLength(1);
    expect(lines[0]?.qty).toBe(15);
    expect(lines[0]?.line).toBe("ПБ 34-15-10п 15");
  });

  it("treats pasted text without qty as wide with qty 1", () => {
    const lines = liveWidePlateLines("ПБ 65-14-8\nПБ 62-15-8 2");
    expect(lines.map((item) => ({ line: item.line, qty: item.qty }))).toEqual([
      { line: "ПБ 65-14-8", qty: 1 },
      { line: "ПБ 62-15-8 2", qty: 2 },
    ]);
  });

  it("treats Excel paste with tabs and шт before qty as wide", () => {
    const text = "ПБ 65-14-8 2\nПБ 62-15-8\tшт\t2\nПБ 18-15-8\tшт\t38";
    const lines = liveWidePlateLines(text);
    expect(lines.map((item) => ({ line: item.line, qty: item.qty }))).toEqual([
      { line: "ПБ 65-14-8 2", qty: 2 },
      { line: "ПБ 62-15-8\tшт\t2", qty: 2 },
      { line: "ПБ 18-15-8\tшт\t38", qty: 38 },
    ]);
  });

  it("accepts шт suffix and missing space after ПБ", () => {
    const lines = liveWidePlateLines("ПБ65-14-8 2 шт\nПБ 18-15-8 2шт");
    expect(lines).toHaveLength(2);
    expect(lines[0]?.qty).toBe(2);
    expect(lines[1]?.qty).toBe(2);
  });
});

describe("overlayDraftWithLiveWideLines", () => {
  it("keeps server id and does not mark resolved when paste lacks qty", () => {
    const draft = makeDraft([{ id: "wide-1", line: "ПБ 65-14-8п", qty: 1 }]);
    const overlay = overlayDraftWithLiveWideLines(draft, "ПБ 65-14-8");
    expect(overlay.metadata.wide_plate_lines[0]?.id).toBe("wide-1");
    expect(overlay.metadata.wide_plates_resolved).toBe(false);
    expect(overlay.metadata.wide_plate_lines.length).toBeGreaterThan(0);
  });

  it("keeps resolved flag after apply even if lines remain wide", () => {
    const draft = makeDraft([{ id: "wide-1", line: "ПБ 62-15-8 2", qty: 2 }], true);
    const overlay = overlayDraftWithLiveWideLines(draft, "ПБ 62-15-8 2");
    expect(overlay.metadata.wide_plates_resolved).toBe(true);
  });
});

describe("widePlateMatchKey", () => {
  it("aligns raw paste with normalized server line", () => {
    expect(widePlateMatchKey("ПБ 65-14-8п 1")).toBe(widePlateMatchKey("ПБ 65-14-8"));
    expect(widePlateMatchKey("ПБ 18-15-8\tшт\t38")).toBe(widePlateMatchKey("ПБ 18-15-8 38"));
  });
});
