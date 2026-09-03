import { describe, expect, it } from "vitest";

import {
  buildPlateLineHighlightMap,
  mergeReviewHighlights,
  splitLineByDoborMarker,
} from "@/features/commercial-offer/lib/plateLineHighlights";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makeDraft = (overrides: Partial<CommercialDraftDetails["metadata"]> = {}): CommercialDraftDetails => ({
  draft_id: "draft-1",
  order: {},
  optimization: { total_plates: 0, total_cost: 0 },
  order_data: [],
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
    dobor_pairs: [],
    diagnostics: [],
    price_rows_count: 0,
    breakdown_tables_count: 0,
    total_sum: 0,
    plate_batches: [],
    wide_plates_resolved: true,
    last_source_filename: "",
    current_step: "plates",
    current_save_mode: null,
    execution_terms: "",
    logistics_cost: 0,
    ...overrides,
  },
  wizard_state: {
    current_step: "plates",
    can_proceed_to: [],
    next_required_action: "none",
    validation_errors: [],
  },
  files: [],
  saved_offer: null,
  totals: {},
  offer_identity: { offer_number: "1", offer_date: "01.01.2026", file_stem: "kp" },
});

describe("buildPlateLineHighlightMap", () => {
  it("highlights both dobor pair lines with partner title", () => {
    const draft = makeDraft({
      dobor_pairs: [
        {
          id: "dobor-1",
          source_line: "ПБ 57-7,2-8п + доб 5-шт",
          primary_line: "ПБ 57-7,2-8п 5",
          complement_line: "ПБ 57-4,8-8п 5",
        },
      ],
    });
    const textLines = ["ПБ 57-7,2-8п 5", "ПБ 57-4,8-8п 5"];

    const map = buildPlateLineHighlightMap(draft, textLines);

    expect(map.get(0)).toMatchObject({
      kind: "dobor",
      title: "Добор: пара с «ПБ 57-4,8-8п 5»",
      doborPairId: "dobor-1",
    });
    expect(map.get(1)).toMatchObject({
      kind: "dobor",
      title: "Добор: пара с «ПБ 57-7,2-8п 5»",
      doborPairId: "dobor-1",
    });
  });

  it("highlights 29-8-8п as invalid_width", () => {
    const draft = makeDraft({
      invalid_width_lines: [
        {
          id: "invalid-width-1",
          name: "Плиты ПБ 29-8-8п",
          line: "ПБ 29-8-8п 1",
          qty: 1,
          length_m: 2.9,
          width_m: 0.8,
          width_mm: 800,
          load_class: 800,
          replacements: [
            { width_mm: 720, width_label: "7,2" },
            { width_mm: 860, width_label: "8,6" },
          ],
        },
      ],
      invalid_widths_resolved: false,
    });

    const map = buildPlateLineHighlightMap(draft, ["ПБ 29-8-8п 1", "ПБ 29-12-8п 1", "ПБ 78-0.3-8п 1"]);

    expect(map.get(0)?.kind).toBe("invalid_width");
    expect(map.get(1)?.kind).toBeUndefined();
    expect(map.get(2)?.kind).toBeUndefined();
  });

  it("prefers wide over invalid_width when both match", () => {
    const draft = makeDraft({
      wide_plate_lines: [{ id: "wide-1", line: "ПБ 59-15-8п 2", qty: 2 }],
      invalid_width_lines: [
        {
          id: "invalid-width-1",
          name: "ПБ 59-15-8п",
          line: "ПБ 59-15-8п 2",
          qty: 2,
          length_m: 5.9,
          width_m: 1.5,
          width_mm: 1500,
          load_class: 800,
          replacements: [],
        },
      ],
      invalid_widths_resolved: false,
    });

    const map = buildPlateLineHighlightMap(draft, ["ПБ 59-15-8п 2"]);

    expect(map.get(0)?.kind).toBe("wide");
  });

  it("prefers wide over dobor when both match", () => {
    const draft = makeDraft({
      wide_plate_lines: [{ id: "wide-1", line: "ПБ 59-15-8п 2", qty: 2 }],
      dobor_pairs: [
        {
          id: "dobor-1",
          source_line: "ПБ 59-15-8п + доб 2",
          primary_line: "ПБ 59-15-8п 2",
          complement_line: "ПБ 59-0,0-8п 2",
        },
      ],
    });

    const map = buildPlateLineHighlightMap(draft, ["ПБ 59-15-8п 2"]);

    expect(map.get(0)?.kind).toBe("wide");
  });

  it("does not crash for orphaned dobor metadata", () => {
    const draft = makeDraft({
      dobor_pairs: [
        {
          id: "dobor-1",
          source_line: "ПБ 57-7,2-8п + доб 5-шт",
          primary_line: "ПБ 57-7,2-8п 5",
          complement_line: "ПБ 57-4,8-8п 5",
        },
      ],
    });

    const map = buildPlateLineHighlightMap(draft, ["ПБ 78-12-8п 2"]);

    expect(map.size).toBe(0);
  });
});

describe("splitLineByDoborMarker", () => {
  it("splits dobor marker variants for inline green highlight", () => {
    const cases: Array<{ line: string; marker: string }> = [
      { line: "ПБ 57-7,2-8п + доб 5", marker: " + доб" },
      { line: "ПБ 57-7,2-8п +доб 5", marker: " +доб" },
      { line: "ПБ 57-7,2-8п доб 5", marker: " доб" },
      { line: "ПБ 57-7,2-8п ДОБОР 5-шт", marker: " ДОБОР" },
    ];

    for (const { line, marker } of cases) {
      const segments = splitLineByDoborMarker(line);
      const markerSegments = segments.filter((segment) => segment.isMarker);
      expect(markerSegments).toHaveLength(1);
      expect(markerSegments[0]?.text).toBe(marker);
      expect(segments.map((segment) => segment.text).join("")).toBe(line);
    }
  });

  it("returns a single segment when no dobor marker is present", () => {
    expect(splitLineByDoborMarker("ПБ 57-7,2-8п 5")).toEqual([{ text: "ПБ 57-7,2-8п 5", isMarker: false }]);
  });
});

describe("mergeReviewHighlights (S3–S6)", () => {
  it("S3: lint parser-reject lines are unparsed (red)", () => {
    const map = mergeReviewHighlights(makeDraft(), "плохо\nПБ 78-12-8п 2", [
      { index: 0, text: "плохо", empty: false, ok: false, reason_text: "не совпал формат строки" },
      { index: 1, text: "ПБ 78-12-8п 2", empty: false, ok: true, reason_text: null },
    ]);

    expect(map.get(0)?.kind).toBe("unparsed");
    expect(map.get(0)?.title).toBe("не совпал формат строки");
    expect(map.get(1)).toBeUndefined();
  });

  it("S4: OCR correction stays yellow when the parser accepts the line", () => {
    const draft = makeDraft({
      ocr_corrections: [
        {
          row_index: 1,
          action: "changed_mark",
          before: { normalized_candidate: "ПБ 78-12-8р", qty: 2 },
          after: { normalized_candidate: "ПБ 78-12-8п", qty: 2 },
        },
      ],
    });
    const map = mergeReviewHighlights(draft, "ПБ 78-12-8п 2", [
      { index: 0, text: "ПБ 78-12-8п 2", empty: false, ok: true, reason_text: null },
    ]);

    expect(map.get(0)?.kind).toBe("correction");
  });

  it("S4: does not add a н-suffix yellow heuristic", () => {
    const map = mergeReviewHighlights(makeDraft(), "ПБ 78-12-8н 2", [
      { index: 0, text: "ПБ 78-12-8н 2", empty: false, ok: true, reason_text: null },
    ]);

    expect(map.size).toBe(0);
  });

  it("S5: parser-accepted 8н is not highlighted, even if draft unparsed_lines is stale", () => {
    const draft = makeDraft({ unparsed_lines: ["ПБ 78-12-8н 2"] });
    const map = mergeReviewHighlights(draft, "ПБ 78-12-8н 2", [
      { index: 0, text: "ПБ 78-12-8н 2", empty: false, ok: true, reason_text: null },
    ]);

    expect(map.get(0)).toBeUndefined();
  });

  it("S6: does not rewrite н to п — highlight map leaves the source line unchanged", () => {
    const line = "ПБ 78-12-8н 2";
    const map = mergeReviewHighlights(makeDraft(), line, [
      { index: 0, text: line, empty: false, ok: true, reason_text: null },
    ]);

    expect(map.size).toBe(0);
    expect(line).toBe("ПБ 78-12-8н 2");
  });
});
