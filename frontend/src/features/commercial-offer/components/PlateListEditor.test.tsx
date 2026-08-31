import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { PlateListEditor } from "@/features/commercial-offer/components/PlateListEditor";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makeDraft = (): CommercialDraftDetails => ({
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

afterEach(() => {
  cleanup();
});

describe("PlateListEditor line numbers", () => {
  const listText = "ПБ 78-12-8н 2\n\n71-12-8 3";

  it("does not render gutter numbers by default", () => {
    render(<PlateListEditor draft={makeDraft()} value={listText} onChange={vi.fn()} />);

    expect(screen.queryAllByTestId("plate-line-number")).toHaveLength(0);
    expect(screen.getByPlaceholderText("Пока нет списка плит.")).toHaveValue(listText);
  });

  it("renders 1-based numbers for non-empty lines only and keeps value unprefixed", () => {
    render(
      <PlateListEditor draft={makeDraft()} value={listText} onChange={vi.fn()} showLineNumbers />,
    );

    const numbers = screen.getAllByTestId("plate-line-number").map((node) => node.textContent);
    expect(numbers).toEqual(["1. ", "2. "]);
    expect(screen.getByPlaceholderText("Пока нет списка плит.")).toHaveValue(listText);
    expect(screen.getByPlaceholderText("Пока нет списка плит.")).not.toHaveValue(
      expect.stringContaining("1."),
    );
  });
});

describe("PlateListEditor external highlights", () => {
  it("paints a line from an external map without a draft", () => {
    const highlights = new Map([
      [0, { kind: "unparsed" as const, title: "не совпал формат строки" }],
    ]);

    render(
      <PlateListEditor
        value={"плохо\nПБ 78-12-8п 2"}
        onChange={vi.fn()}
        highlights={highlights}
      />,
    );

    expect(screen.getByTitle("не совпал формат строки")).toBeInTheDocument();
    expect(screen.getByTitle("не совпал формат строки")).toHaveTextContent("плохо");
  });
});
