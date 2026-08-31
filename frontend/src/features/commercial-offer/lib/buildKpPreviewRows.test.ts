import { describe, expect, it } from "vitest";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import { buildKpPreviewRows, parseWidePlateLine } from "@/features/commercial-offer/lib/buildKpPreviewRows";

const baseMetadata = (): CommercialDraftDetails["metadata"] => ({
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
  wide_plates_resolved: false,
  last_source_filename: "",
  current_step: "plates",
  current_save_mode: null,
  execution_terms: "",
  logistics_cost: 0,
});

type DraftOverrides = Partial<Omit<CommercialDraftDetails, "metadata" | "wizard_state">> & {
  metadata?: Partial<CommercialDraftDetails["metadata"]>;
  wizard_state?: Partial<CommercialDraftDetails["wizard_state"]>;
};

function makeDraft(overrides: DraftOverrides = {}): CommercialDraftDetails {
  const { metadata: metaOverrides, ...rest } = overrides;
  return {
    draft_id: "draft-test",
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
    ...rest,
    metadata: { ...baseMetadata(), ...metaOverrides },
  };
}

describe("parseWidePlateLine", () => {
  it("parses wide plate mark with qty suffix", () => {
    const parsed = parseWidePlateLine("ПБ 59-15-8п 2");
    expect(parsed).not.toBeNull();
    expect(parsed?.lengthM).toBeCloseTo(5.9);
    expect(parsed?.widthM).toBeCloseTo(1.5);
  });
});

describe("buildKpPreviewRows", () => {
  it("returns standard rows without flags", () => {
    const draft = makeDraft({
      order_data: [
        {
          name: "Плиты ПБ 71-12-8п",
          qty: 3,
          unit_price: 12450,
          length_m: 7.1,
          width_m: 1.2,
        },
      ],
    });

    const rows = buildKpPreviewRows(draft);
    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({
      name: "Плиты ПБ 71-12-8п",
      qty: 3,
      unitPrice: 12450,
      flag: null,
      lineId: null,
    });
  });

  it("passes through line_id as lineId", () => {
    const draft = makeDraft({
      order_data: [
        {
          line_id: "ln_plate_1",
          name: "Плиты ПБ 71-12-8п",
          qty: 3,
          unit_price: 12450,
          length_m: 7.1,
          width_m: 1.2,
        },
      ],
    });
    expect(buildKpPreviewRows(draft)[0]?.lineId).toBe("ln_plate_1");
  });

  it("marks direct wide plate when width_m > 1.2", () => {
    const draft = makeDraft({
      metadata: {
        wide_plate_lines: [{ id: "wide-1", line: "ПБ 59-15-8п 2", qty: 2 }],
      },
      order_data: [
        {
          name: "Плиты ПБ 59-15-8п",
          qty: 2,
          unit_price: 15000,
          length_m: 5.9,
          width_m: 1.5,
        },
      ],
    });

    const rows = buildKpPreviewRows(draft);
    expect(rows[0].flag).toBe("wide_direct");
    expect(rows[0].sourceLine).toBe("ПБ 59-15-8п 2");
  });

  it("marks split wide plate rows linked to wide_plate_lines", () => {
    const draft = makeDraft({
      metadata: {
        wide_plate_lines: [{ id: "wide-1", line: "ПБ 59-15-8п 2", qty: 2 }],
      },
      order_data: [
        {
          name: "Плиты ПБ 59-12-8п",
          qty: 2,
          unit_price: 12000,
          length_m: 5.9,
          width_m: 1.2,
        },
        {
          name: "Плиты ПБ 59-3-8п",
          qty: 2,
          unit_price: 4000,
          length_m: 5.9,
          width_m: 0.3,
        },
      ],
    });

    const rows = buildKpPreviewRows(draft);
    expect(rows).toHaveLength(2);
    expect(rows[0].flag).toBe("wide_split");
    expect(rows[0].sourceLine).toBe("ПБ 59-15-8п 2");
    expect(rows[1].flag).toBe("wide_split");
    expect(rows[1].sourceLine).toBe("ПБ 59-15-8п 2");
  });

  it("clears flags after wide plates are resolved", () => {
    const draft = makeDraft({
      metadata: {
        wide_plate_lines: [{ id: "wide-1", line: "ПБ 59-15-8п 2", qty: 2 }],
        wide_plates_resolved: true,
      },
      order_data: [
        {
          name: "Плиты ПБ 59-12-8п",
          qty: 2,
          unit_price: 12000,
          length_m: 5.9,
          width_m: 1.2,
        },
        {
          name: "Плиты ПБ 59-3-8п",
          qty: 2,
          unit_price: 4000,
          length_m: 5.9,
          width_m: 0.3,
        },
      ],
    });

    const rows = buildKpPreviewRows(draft);
    expect(rows.every((row) => row.flag === null)).toBe(true);
  });

  it("returns empty rows for empty order_data", () => {
    const draft = makeDraft({ order_data: [] });
    expect(buildKpPreviewRows(draft)).toEqual([]);
  });

  it("does not include unparsed lines in preview rows", () => {
    const draft = makeDraft({
      metadata: {
        unparsed_lines: ["непонятная строка (пропущено: ...)"],
      },
      order_data: [
        {
          name: "Плиты ПБ 71-12-8п",
          qty: 1,
          unit_price: 1000,
          length_m: 7.1,
          width_m: 1.2,
        },
      ],
    });

    const rows = buildKpPreviewRows(draft);
    expect(rows).toHaveLength(1);
    expect(rows.some((row) => row.name.includes("непонятная"))).toBe(false);
  });
});
