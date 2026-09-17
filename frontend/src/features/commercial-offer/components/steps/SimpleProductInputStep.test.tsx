import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SimpleProductInputStep } from "./SimpleProductInputStep";
import { PRODUCT_TYPE_CONFIG } from "@/features/commercial-offer/lib/productTypeConfig";
import { useSourceTextLint } from "@/features/commercial-offer/hooks/useSourceTextLint";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import type { PageSource } from "@/features/commercial-offer/lib/multiPageSource";
import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";

vi.mock("@/features/commercial-offer/hooks/useSourceTextLint", () => ({
  useSourceTextLint: vi.fn(),
}));

vi.mock("@/features/commercial-offer/api/commercialOfferApi", () => ({
  commercialOfferApi: {
    getPriceCatalog: vi.fn(),
  },
}));

const useLint = useSourceTextLint as unknown as ReturnType<typeof vi.fn>;
const noop = vi.fn();

const MARCH_LIST_PLACEHOLDER = PRODUCT_TYPE_CONFIG.marches.labels.placeholder;

afterEach(() => {
  cleanup();
});

beforeEach(() => {
  vi.clearAllMocks();
  useLint.mockReturnValue({ lines: [], isPending: false, isError: false });
});

describe("SimpleProductInputStep (marches) placeholder", () => {
  it("hints real catalog marks from march_prices, not invented ЛМ-1", () => {
    render(
      <SimpleProductInputStep
        productType="marches"
        draft={null}
        pendingBatchReview={false}
        sourceText=""
        batchReviewText=""
        normalizedText=""
        recognizedImageUrl={null}
        recognizedImageName={null}
        errorMessage={null}
        isRecognizing={false}
        onTextChange={noop}
        onBatchReviewTextChange={noop}
        pages={[]}
        activePageId={null}
        onAddFiles={noop}
        onRemovePage={noop}
        onSelectPage={noop}
        onRecognize={noop}
        onConfirmBatch={noop}
        onFinishInput={noop}
        onReset={noop}
      />,
    );

    expect(MARCH_LIST_PLACEHOLDER).toBe("1ЛМ 27-11-14-4 B25 5\nЛМ 2,8 3");
    const field = screen.getByRole("textbox");
    expect(field).toHaveAttribute("placeholder", MARCH_LIST_PLACEHOLDER);
    expect(field.getAttribute("placeholder")).not.toContain("ЛМ-1");
  });
});

describe("SimpleProductInputStep (marches) AI on batch-review", () => {
  const makeDraft = (text: string): CommercialDraftDetails =>
    ({
      draft_id: "draft-1",
      metadata: {
        product_type: "marches",
        normalized_text: text,
        march_batches: [{ batch_index: 0, normalized_text: text, source_kind: "image" }],
        ocr_corrections: [],
      },
    }) as CommercialDraftDetails;

  const makePage = (id: string): PageSource => ({
    id,
    file: new File(["x"], `${id}.png`, { type: "image/png" }),
    name: `${id}.png`,
    previewUrl: `blob:${id}`,
    status: "ready",
    batchReviewText: "1ЛМ 27-11-14-4 B25 5",
  });

  it("shows AI instruction controls on batch-review and applies when clicked", () => {
    const onApplyAi = vi.fn();

    render(
      <SimpleProductInputStep
        productType="marches"
        draft={makeDraft("1ЛМ 27-11-14-4 B25 5")}
        pendingBatchReview
        sourceText=""
        batchReviewText="1ЛМ 27-11-14-4 B25 5"
        normalizedText="1ЛМ 27-11-14-4 B25 5"
        recognizedImageUrl={null}
        recognizedImageName={null}
        errorMessage={null}
        isRecognizing={false}
        aiInstruction="убери строки с B15"
        onAiInstructionChange={noop}
        onApplyAi={onApplyAi}
        onTextChange={noop}
        onBatchReviewTextChange={noop}
        pages={[makePage("a")]}
        activePageId="a"
        onAddFiles={noop}
        onRemovePage={noop}
        onSelectPage={noop}
        onRecognize={noop}
        onConfirmBatch={noop}
        onFinishInput={noop}
        onReset={noop}
      />,
    );

    expect(screen.getByText("Инструкция для помощника")).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Применить инструкцию" }));
    expect(onApplyAi).toHaveBeenCalledTimes(1);
  });
});

describe("SimpleProductInputStep (marches) rerecognize button", () => {
  const makeDraft = (text: string): CommercialDraftDetails =>
    ({
      draft_id: "draft-1",
      metadata: {
        product_type: "marches",
        normalized_text: text,
        march_batches: [{ batch_index: 0, normalized_text: text, source_kind: "image" }],
        ocr_corrections: [],
      },
    }) as CommercialDraftDetails;

  const makePage = (id: string, status: PageSource["status"] = "ready"): PageSource => ({
    id,
    file: new File(["x"], `${id}.png`, { type: "image/png" }),
    name: `${id}.png`,
    previewUrl: `blob:${id}`,
    status,
    batchReviewText: "1ЛМ 27-11-14-4 B25 5",
  });

  const baseProps = {
    productType: "marches" as const,
    draft: makeDraft("1ЛМ 27-11-14-4 B25 5"),
    pendingBatchReview: true as const,
    sourceText: "",
    batchReviewText: "1ЛМ 27-11-14-4 B25 5",
    normalizedText: "1ЛМ 27-11-14-4 B25 5",
    recognizedImageUrl: "blob:photo",
    recognizedImageName: "a.png",
    errorMessage: null,
    isRecognizing: false,
    onTextChange: noop,
    onBatchReviewTextChange: noop,
    onAddFiles: noop,
    onRemovePage: noop,
    onSelectPage: noop,
    onRecognize: noop,
    onRerecognize: noop,
    onConfirmBatch: noop,
    onFinishInput: noop,
    onReset: noop,
  };

  it("shows Перераспознать on ready and confirmed pages", () => {
    const { rerender } = render(
      <SimpleProductInputStep {...baseProps} pages={[makePage("a", "ready")]} activePageId="a" />,
    );
    expect(screen.getByRole("button", { name: "Перераспознать" })).toBeEnabled();
    expect(screen.queryByText("Нестандартная ширина")).not.toBeInTheDocument();

    rerender(
      <SimpleProductInputStep {...baseProps} pages={[makePage("a", "confirmed")]} activePageId="a" />,
    );
    expect(screen.getByRole("button", { name: "Перераспознать" })).toBeEnabled();
    expect(screen.queryByText("Нестандартная ширина")).not.toBeInTheDocument();
  });
});

describe("SimpleProductInputStep price catalog", () => {
  const unpricedDraft = {
    draft_id: "draft-piles-1",
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data: [
      {
        line_id: "ln1",
        product_type: "piles",
        mark: "C110.30-6",
        name: "C110.30-6",
        qty: 2,
        unit_price: null,
        concrete_grade: "B25",
      },
      {
        line_id: "ln2",
        product_type: "piles",
        mark: "C110.40-8.1",
        name: "C110.40-8.1",
        qty: 1,
        unit_price: null,
        concrete_grade: "B25",
      },
    ],
    files: [],
    saved_offer: null,
    totals: {},
    offer_identity: { offer_number: "", offer_date: "", file_stem: "" },
    metadata: {
      product_type: "piles",
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
      wide_plates_resolved: true,
      last_source_filename: "",
      current_step: "piles",
      current_save_mode: null,
      execution_terms: "",
      logistics_cost: 0,
      default_concrete_grade: "B25",
    },
    wizard_state: {
      current_step: "piles",
      can_proceed_to: [],
      next_required_action: "none",
      validation_errors: ["Нет цен для позиций: C110.30-6"],
    },
  } as CommercialDraftDetails;

  it("fetches the price catalog when Прайс is opened", async () => {
    vi.mocked(commercialOfferApi.getPriceCatalog).mockResolvedValue({
      items: [{ mark: "С110.30-9", concrete_grade: "B25", price: 27585.43 }],
    });

    render(
      <SimpleProductInputStep
        productType="piles"
        draft={unpricedDraft}
        pendingBatchReview={false}
        sourceText=""
        batchReviewText=""
        normalizedText=""
        recognizedImageUrl={null}
        recognizedImageName={null}
        errorMessage={null}
        isRecognizing={false}
        onTextChange={noop}
        onBatchReviewTextChange={noop}
        pages={[]}
        activePageId={null}
        onAddFiles={noop}
        onRemovePage={noop}
        onSelectPage={noop}
        onRecognize={noop}
        onConfirmBatch={noop}
        onFinishInput={noop}
        onReset={noop}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: "Прайс" }));
    await waitFor(() => {
      expect(commercialOfferApi.getPriceCatalog).toHaveBeenCalledWith({
        productType: "piles",
        q: "C110",
      });
    });
    expect(await screen.findByText("С110.30-9")).toBeInTheDocument();
  });
});
