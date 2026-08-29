import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { KpBridgePilePreviewPanel } from "@/features/commercial-offer/components/KpBridgePilePreviewPanel";
import { KpFbsPreviewPanel } from "@/features/commercial-offer/components/KpFbsPreviewPanel";
import { KpMarchPreviewPanel } from "@/features/commercial-offer/components/KpMarchPreviewPanel";
import { KpPilePreviewPanel } from "@/features/commercial-offer/components/KpPilePreviewPanel";
import type { CommercialDraftDetails, ProductType } from "@/features/commercial-offer/types/commercialOffer";

const makeDraft = (productType: ProductType, name: string): CommercialDraftDetails =>
  ({
    draft_id: `draft-${productType}`,
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data: [{ mark: name, name, qty: 1, unit_price: 1000, concrete_grade: "B25" }],
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    metadata: {
      product_type: productType,
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
      warnings: ["Не удалось распознать строк: 1"],
      unparsed_lines: ["плохо"],
      normalized_text: name,
      normalized_lines: [],
      wide_plate_lines: [],
      diagnostics: [],
      price_rows_count: 0,
      breakdown_tables_count: 0,
      total_sum: 0,
      plate_batches: [],
      wide_plates_resolved: true,
      last_source_filename: "",
      current_step: productType,
      current_save_mode: null,
      execution_terms: "",
      logistics_cost: 0,
    },
    wizard_state: {
      current_step: productType,
      can_proceed_to: ["client"],
      next_required_action: "none",
      validation_errors: [],
    },
  }) as CommercialDraftDetails;

describe("Kp*PreviewPanel unparsed UX", () => {
  it.each([
    ["piles", KpPilePreviewPanel, "С120.35-12"],
    ["fbs", KpFbsPreviewPanel, "ФБС 9.3.6-Т"],
    ["marches", KpMarchPreviewPanel, "1ЛМ 27-11-14-4"],
    ["bridge_piles", KpBridgePilePreviewPanel, "С7-35Т5"],
  ] as const)("%s hides unparsed list and count banner", (_type, Panel, name) => {
    render(<Panel draft={makeDraft(_type, name)} normalizedText={name} />);
    expect(screen.queryByText("Не попали в состав")).not.toBeInTheDocument();
    expect(screen.queryByText(/Не удалось распознать строк: 1/)).not.toBeInTheDocument();
  });
});
