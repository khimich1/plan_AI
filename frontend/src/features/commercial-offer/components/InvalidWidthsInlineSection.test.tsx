import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import {
  defaultInvalidWidthDecision,
  InvalidWidthsInlineSection,
} from "@/features/commercial-offer/components/InvalidWidthsInlineSection";
import type { CommercialDraftDetails, InvalidWidthLine } from "@/features/commercial-offer/types/commercialOffer";

const invalidLine = (overrides: Partial<InvalidWidthLine> = {}): InvalidWidthLine => ({
  id: "invalid-width-1",
  name: "Плиты ПБ 29-8-8п",
  line: "ПБ 29-8-8п 1",
  qty: 1,
  length_m: 2.9,
  width_m: 0.8,
  width_mm: 800,
  load_class: 800,
  replacements: [
    { width_mm: 720, width_label: "7,2", price: 9000 },
    { width_mm: 860, width_label: "8,6", price: 10400 },
  ],
  ...overrides,
});

const makeDraft = (lines: InvalidWidthLine[], resolved = false): CommercialDraftDetails =>
  ({
    draft_id: "draft-1",
    order: {},
    optimization: {},
    order_data: [],
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    metadata: {
      invalid_width_lines: lines,
      invalid_widths_resolved: resolved,
    },
    wizard_state: {
      current_step: "plates",
      can_proceed_to: [],
      next_required_action: "resolve_invalid_widths",
      validation_errors: [],
    },
  }) as CommercialDraftDetails;

describe("InvalidWidthsInlineSection", () => {
  afterEach(() => {
    cleanup();
  });

  it("preselects the upper neighbor 860 for 800 mm", () => {
    const decision = defaultInvalidWidthDecision(invalidLine());
    expect(decision).toEqual({ action: "replace_width", widthMm: 860 });
  });

  it("renders two replacements plus exclude and Apply", () => {
    const onDecisionChange = vi.fn();
    const onApply = vi.fn();
    const line = invalidLine();
    render(
      <InvalidWidthsInlineSection
        draft={makeDraft([line])}
        decisions={{ [line.id]: defaultInvalidWidthDecision(line) }}
        isPending={false}
        onDecisionChange={onDecisionChange}
        onApply={onApply}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: /позиций требуют внимания/ }));
    expect(screen.getByLabelText(/8,6/)).toBeChecked();
    expect(screen.getByLabelText(/7,2/)).toBeInTheDocument();
    expect(screen.getByLabelText("Исключить позицию")).toBeInTheDocument();
    expect(screen.queryByText(/оставить как есть/i)).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Применить" })).toBeEnabled();
  });

  it("disables Apply when a replace_width decision has no width", () => {
    const line = invalidLine();
    render(
      <InvalidWidthsInlineSection
        draft={makeDraft([line])}
        decisions={{ [line.id]: { action: "replace_width", widthMm: null } }}
        isPending={false}
        onDecisionChange={vi.fn()}
        onApply={vi.fn()}
      />,
    );
    fireEvent.click(screen.getByRole("button", { name: /позиций требуют внимания/ }));
    expect(screen.getByRole("button", { name: "Применить" })).toBeDisabled();
  });
});
