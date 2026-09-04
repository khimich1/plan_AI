import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { SaveOfferSection } from "@/features/commercial-offer/components/SaveOfferSection";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const draft = {
  draft_id: "d1",
  order: {},
  optimization: {},
  order_data: [],
  files: [],
  saved_offer: null,
  totals: { total_qty: 0, subtotal: 0, vat_amount: 0, total_with_vat: 0 },
  offer_identity: { offer_number: "КП-1", offer_date: "02.09.2026", file_stem: "kp-1" },
  metadata: {
    source_type: "text",
    original_text: "",
    ocr_text: "",
    input_text: "",
    accumulated_text: "",
    manager_id: 1,
    manager_name: "Иванов",
    manager_phone: "",
    manager_email: "",
    client_name: "Клиент",
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
    product_type: "plates",
    wide_plates_resolved: true,
    last_source_filename: "",
    current_step: "result",
    current_save_mode: null,
    execution_terms: "",
    logistics_cost: 0,
    append_batches: [],
    resume_kp_id: null,
  },
  wizard_state: {
    current_step: "result",
    can_proceed_to: [],
    next_required_action: "none",
    validation_errors: [],
  },
} as CommercialDraftDetails;

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

describe("SaveOfferSection archive-only create save", () => {
  it("hides manufacturing terms and alternative save paths", () => {
    render(
      <SaveOfferSection
        draft={draft}
        lastSaveResult={null}
        isPending={false}
        onSave={vi.fn(async () => undefined)}
      />,
    );

    expect(screen.queryByText("Срок изготовления")).not.toBeInTheDocument();
    expect(screen.queryByText(/Другой вариант сохранения/i)).not.toBeInTheDocument();
    expect(screen.queryByText("В работе")).not.toBeInTheDocument();
    expect(screen.queryByText("Пропустить")).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "В архив" })).toBeInTheDocument();
  });

  it("saves only to archive without execution terms", async () => {
    const onSave = vi.fn(async () => undefined);
    render(
      <SaveOfferSection draft={draft} lastSaveResult={null} isPending={false} onSave={onSave} />,
    );

    fireEvent.click(screen.getByRole("button", { name: "В архив" }));

    await waitFor(() => {
      expect(onSave).toHaveBeenCalledWith({ mode: "archive", executionTermsInput: "" });
    });
  });

  it("disables primary save after create already saved to archive", () => {
    const savedDraft = {
      ...draft,
      saved_offer: {
        kp_id: 1,
        status: "в архиве",
        mode: "archive",
        execution_terms: "",
        saved_at: "2026-09-02T12:00:00",
      },
    } as CommercialDraftDetails;

    render(
      <SaveOfferSection draft={savedDraft} lastSaveResult={null} isPending={false} onSave={vi.fn()} />,
    );

    expect(screen.getByRole("button", { name: "Сохранено" })).toBeDisabled();
  });

  it("shows enabled Сохранить изменения when resume_kp_id is set even if saved_offer exists", async () => {
    const resumeDraft = {
      ...draft,
      saved_offer: {
        kp_id: 42,
        status: "в архиве",
        mode: "archive",
        execution_terms: "",
        saved_at: "2026-09-02T12:00:00",
      },
      metadata: { ...draft.metadata, resume_kp_id: 42 },
    } as CommercialDraftDetails;
    const onSave = vi.fn(async () => undefined);

    render(
      <SaveOfferSection draft={resumeDraft} lastSaveResult={null} isPending={false} onSave={onSave} />,
    );

    const button = screen.getByRole("button", { name: "Сохранить изменения" });
    expect(button).not.toBeDisabled();
    fireEvent.click(button);

    await waitFor(() => {
      expect(onSave).toHaveBeenCalledWith({ mode: "archive", executionTermsInput: "" });
    });
  });
});
