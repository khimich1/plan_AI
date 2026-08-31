import { describe, expect, it } from "vitest";

import {
  buildBridgePileLinesFromOrderData,
  buildBridgePilePreviewRows,
} from "@/features/commercial-offer/lib/buildBridgePilePreviewRows";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const baseDraft = (overrides: Partial<CommercialDraftDetails> = {}): CommercialDraftDetails =>
  ({
    draft_id: "d1",
    order_data: [
      {
        mark: "C8-35T1",
        name: "C8-35T1",
        concrete_grade: "B25",
        available_grades: ["B25"],
        qty: 2,
        unit_price: 35695.27,
        product_kind: "bridge_pile",
      },
    ],
    metadata: {
      product_type: "bridge_piles",
      default_concrete_grade: "B25",
    },
    wizard_state: {
      current_step: "bridge_piles",
      can_proceed_to: ["client"],
      next_required_action: "none",
      validation_errors: [],
    },
    ...overrides,
  }) as CommercialDraftDetails;

describe("buildBridgePilePreviewRows", () => {
  it("maps order_data to preview rows with available grades", () => {
    const rows = buildBridgePilePreviewRows(baseDraft());
    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({
      mark: "C8-35T1",
      concrete_grade: "B25",
      available_grades: ["B25"],
      qty: 2,
      product_kind: "bridge_pile",
    });
  });

  it("buildBridgePileLinesFromOrderData joins mark grade qty", () => {
    const text = buildBridgePileLinesFromOrderData([
      { mark: "C8-35В4", concrete_grade: "B25", qty: 2, unit_price: 100, name: "C8-35В4", product_kind: "bridge_pile" },
      { mark: "C13-40T3", concrete_grade: "B30", qty: 3, unit_price: 200, name: "C13-40T3", product_kind: "bridge_pile" },
    ]);
    expect(text).toBe("C8-35В4 B25 2\nC13-40T3 B30 3");
  });
});
