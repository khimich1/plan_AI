import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { MARCH_LIST_PLACEHOLDER, MarchInputStep } from "./MarchInputStep";
import { useSourceTextLint } from "@/features/commercial-offer/hooks/useSourceTextLint";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import type { PageSource } from "@/features/commercial-offer/lib/multiPageSource";

vi.mock("@/features/commercial-offer/hooks/useSourceTextLint", () => ({
  useSourceTextLint: vi.fn(),
}));

const useLint = useSourceTextLint as unknown as ReturnType<typeof vi.fn>;
const noop = vi.fn();

afterEach(() => {
  cleanup();
});

beforeEach(() => {
  vi.clearAllMocks();
  useLint.mockReturnValue({ lines: [], isPending: false, isError: false });
});

describe("MarchInputStep placeholder", () => {
  it("hints real catalog marks from march_prices, not invented ЛМ-1", () => {
    render(
      <MarchInputStep
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
        onFinishMarches={noop}
        onReset={noop}
      />,
    );

    expect(MARCH_LIST_PLACEHOLDER).toBe("1ЛМ 27-11-14-4 B25 5\nЛМ 2,8 3");
    const field = screen.getByRole("textbox");
    expect(field).toHaveAttribute("placeholder", MARCH_LIST_PLACEHOLDER);
    expect(field.getAttribute("placeholder")).not.toContain("ЛМ-1");
  });
});

describe("MarchInputStep AI on batch-review", () => {
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
      <MarchInputStep
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
        onFinishMarches={noop}
        onReset={noop}
      />,
    );

    expect(screen.getByText("Инструкция для помощника")).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Применить инструкцию" }));
    expect(onApplyAi).toHaveBeenCalledTimes(1);
  });
});
