import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { PlateInputStep } from "@/features/commercial-offer/components/steps/PlateInputStep";
import { useSourceTextLint } from "@/features/commercial-offer/hooks/useSourceTextLint";
import { OCR_VERIFY_FAILED_REVIEW_MESSAGE } from "@/features/commercial-offer/lib/ocrVerifyFailed";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import type { PageSource } from "@/features/commercial-offer/lib/multiPageSource";

vi.mock("@/features/commercial-offer/hooks/useSourceTextLint", () => ({
  useSourceTextLint: vi.fn(),
}));

const useLint = useSourceTextLint as unknown as ReturnType<typeof vi.fn>;
const noop = vi.fn();

const makeDraft = (text: string): CommercialDraftDetails =>
  ({
    draft_id: "draft-1",
    metadata: {
      product_type: "plates",
      normalized_text: text,
      plate_batches: [{ batch_index: 0, normalized_text: text, source_kind: "image" }],
      ocr_corrections: [],
    },
  }) as CommercialDraftDetails;

const makePage = (id: string, status: PageSource["status"] = "ready"): PageSource => ({
  id,
  file: new File(["x"], `${id}.png`, { type: "image/png" }),
  name: `${id}.png`,
  previewUrl: `blob:${id}`,
  status,
  batchReviewText: "ПБ 78-12-8п 2",
});

afterEach(() => {
  cleanup();
});

beforeEach(() => {
  vi.clearAllMocks();
  useLint.mockReturnValue({ lines: [], isPending: false, isError: false });
});

describe("PlateInputStep AI on batch-review", () => {
  it("shows AI instruction controls on batch-review and applies when clicked", () => {
    const onAiInstructionChange = vi.fn();
    const onApplyAi = vi.fn();

    render(
      <PlateInputStep
        draft={makeDraft("ПБ 78-12-8п 2")}
        pendingBatchReview
        sourceText=""
        batchReviewText="ПБ 78-12-8п 2"
        normalizedText="ПБ 78-12-8п 2"
        pages={[makePage("a")]}
        activePageId="a"
        recognizedImageUrl={null}
        recognizedImageName={null}
        errorMessage={null}
        isRecognizing={false}
        aiInstruction="замени н на п"
        onAiInstructionChange={onAiInstructionChange}
        onApplyAi={onApplyAi}
        onTextChange={noop}
        onBatchReviewTextChange={noop}
        onAddFiles={noop}
        onRemovePage={noop}
        onSelectPage={noop}
        onRecognize={noop}
        onConfirmBatch={noop}
        onFinishPlates={noop}
        onReset={noop}
      />,
    );

    expect(screen.getByText("Инструкция для помощника")).toBeInTheDocument();
    expect(screen.getByPlaceholderText(/замени н на п/i)).toHaveValue("замени н на п");
    expect(screen.getByDisplayValue("ПБ 78-12-8п 2")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Применить инструкцию" }));
    expect(onApplyAi).toHaveBeenCalledTimes(1);
  });

  // S21 — AI must stay enabled while progressive OCR tail is still busy
  it("keeps AI apply enabled when page1 is ready and recognizing tail is busy", () => {
    render(
      <PlateInputStep
        draft={makeDraft("ПБ 78-12-8п 2")}
        pendingBatchReview
        sourceText=""
        batchReviewText="ПБ 78-12-8п 2"
        normalizedText="ПБ 78-12-8п 2"
        pages={[makePage("a", "ready"), makePage("b", "pending")]}
        activePageId="a"
        recognizedImageUrl={null}
        recognizedImageName={null}
        errorMessage={null}
        isRecognizing
        isAiProcessing={false}
        aiInstruction="замени н на п"
        onAiInstructionChange={noop}
        onApplyAi={noop}
        onTextChange={noop}
        onBatchReviewTextChange={noop}
        onAddFiles={noop}
        onRemovePage={noop}
        onSelectPage={noop}
        onRecognize={noop}
        onConfirmBatch={noop}
        onFinishPlates={noop}
        onReset={noop}
      />,
    );

    expect(screen.getByRole("button", { name: "Применить инструкцию" })).not.toBeDisabled();
  });

  it("S1: review list beside the photo shows updated text after Apply success", () => {
    const { rerender } = render(
      <PlateInputStep
        draft={makeDraft("ПБ 78-12-8н 2")}
        pendingBatchReview
        sourceText=""
        batchReviewText="ПБ 78-12-8н 2"
        normalizedText="ПБ 78-12-8н 2"
        pages={[makePage("a")]}
        activePageId="a"
        recognizedImageUrl={null}
        recognizedImageName={null}
        errorMessage={null}
        isRecognizing={false}
        aiInstruction=""
        onAiInstructionChange={noop}
        onApplyAi={noop}
        onTextChange={noop}
        onBatchReviewTextChange={noop}
        onAddFiles={noop}
        onRemovePage={noop}
        onSelectPage={noop}
        onRecognize={noop}
        onConfirmBatch={noop}
        onFinishPlates={noop}
        onReset={noop}
      />,
    );

    expect(screen.getByDisplayValue("ПБ 78-12-8н 2")).toBeInTheDocument();

    rerender(
      <PlateInputStep
        draft={makeDraft("ПБ 78-12-8п 2")}
        pendingBatchReview
        sourceText=""
        batchReviewText="ПБ 78-12-8п 2"
        normalizedText="ПБ 78-12-8п 2"
        pages={[makePage("a")]}
        activePageId="a"
        recognizedImageUrl={null}
        recognizedImageName={null}
        errorMessage={null}
        isRecognizing={false}
        aiInstruction=""
        onAiInstructionChange={noop}
        onApplyAi={noop}
        onTextChange={noop}
        onBatchReviewTextChange={noop}
        onAddFiles={noop}
        onRemovePage={noop}
        onSelectPage={noop}
        onRecognize={noop}
        onConfirmBatch={noop}
        onFinishPlates={noop}
        onReset={noop}
      />,
    );

    expect(screen.getByDisplayValue("ПБ 78-12-8п 2")).toBeInTheDocument();
    expect(screen.queryByDisplayValue("ПБ 78-12-8н 2")).not.toBeInTheDocument();
  });

  it("S2: shows a visible error when Apply fails", () => {
    render(
      <PlateInputStep
        draft={makeDraft("ПБ 78-12-8н 2")}
        pendingBatchReview
        sourceText=""
        batchReviewText="ПБ 78-12-8н 2"
        normalizedText="ПБ 78-12-8н 2"
        pages={[makePage("a")]}
        activePageId="a"
        recognizedImageUrl={null}
        recognizedImageName={null}
        errorMessage="Не удалось применить инструкцию"
        isRecognizing={false}
        aiInstruction="замени н на п"
        onAiInstructionChange={noop}
        onApplyAi={noop}
        onTextChange={noop}
        onBatchReviewTextChange={noop}
        onAddFiles={noop}
        onRemovePage={noop}
        onSelectPage={noop}
        onRecognize={noop}
        onConfirmBatch={noop}
        onFinishPlates={noop}
        onReset={noop}
      />,
    );

    expect(screen.getByText("Не удалось применить инструкцию")).toBeInTheDocument();
    expect(screen.getByDisplayValue("ПБ 78-12-8н 2")).toBeInTheDocument();
  });
});

describe("PlateInputStep review highlights (S3–S5)", () => {
  const reviewProps = {
    pendingBatchReview: true as const,
    sourceText: "",
    pages: [makePage("a")],
    activePageId: "a",
    recognizedImageUrl: null,
    recognizedImageName: null,
    errorMessage: null,
    isRecognizing: false,
    onAiInstructionChange: noop,
    onApplyAi: noop,
    onTextChange: noop,
    onBatchReviewTextChange: noop,
    onAddFiles: noop,
    onRemovePage: noop,
    onSelectPage: noop,
    onRecognize: noop,
    onConfirmBatch: noop,
    onFinishPlates: noop,
    onReset: noop,
  };

  it("S3: paints parser-reject lines red on batch-review", () => {
    useLint.mockReturnValue({
      isPending: false,
      isError: false,
      lines: [
        { index: 0, text: "плохо", empty: false, ok: false, reason_text: "не совпал формат строки" },
        { index: 1, text: "ПБ 78-12-8п 2", empty: false, ok: true, reason_text: null },
      ],
    });

    render(
      <PlateInputStep
        {...reviewProps}
        draft={makeDraft("плохо\nПБ 78-12-8п 2")}
        batchReviewText="плохо\nПБ 78-12-8п 2"
        normalizedText="плохо\nПБ 78-12-8п 2"
      />,
    );

    expect(screen.getByTitle("не совпал формат строки")).toHaveTextContent("плохо");
  });

  it("S5: does not highlight parser-accepted 8н", () => {
    useLint.mockReturnValue({
      isPending: false,
      isError: false,
      lines: [{ index: 0, text: "ПБ 78-12-8н 2", empty: false, ok: true, reason_text: null }],
    });

    render(
      <PlateInputStep
        {...reviewProps}
        draft={makeDraft("ПБ 78-12-8н 2")}
        batchReviewText="ПБ 78-12-8н 2"
        normalizedText="ПБ 78-12-8н 2"
      />,
    );

    expect(screen.getByDisplayValue("ПБ 78-12-8н 2")).toBeInTheDocument();
    expect(screen.queryByTitle("Строка не попала в расчёт — проверьте вручную")).not.toBeInTheDocument();
    expect(screen.queryByTitle("Строка исправлена при распознавании — сверьте с фото")).not.toBeInTheDocument();
  });
});

describe("PlateInputStep ocr verify-failed banner", () => {
  const oldCopy = "Повторная проверка распознавания не удалась — сверьте список плит с исходным фото вручную.";
  const newCopy = OCR_VERIFY_FAILED_REVIEW_MESSAGE;

  const reviewBase = {
    pendingBatchReview: true as const,
    sourceText: "",
    recognizedImageUrl: null,
    recognizedImageName: null,
    errorMessage: null,
    isRecognizing: false,
    onAiInstructionChange: noop,
    onApplyAi: noop,
    onTextChange: noop,
    onBatchReviewTextChange: noop,
    onAddFiles: noop,
    onRemovePage: noop,
    onSelectPage: noop,
    onRecognize: noop,
    onConfirmBatch: noop,
    onFinishPlates: noop,
    onReset: noop,
  };

  it("shows the softened copy for the active page that failed verify", () => {
    const draft = makeDraft("ПБ 78-12-8п 2");
    draft.metadata.ocr_verify_failed = true;
    const failedPage = { ...makePage("a"), ocrVerifyFailed: true };

    render(
      <PlateInputStep
        {...reviewBase}
        draft={draft}
        batchReviewText="ПБ 78-12-8п 2"
        normalizedText="ПБ 78-12-8п 2"
        pages={[failedPage]}
        activePageId="a"
      />,
    );

    expect(screen.getByText(newCopy)).toBeInTheDocument();
    expect(screen.queryByText(oldCopy)).not.toBeInTheDocument();
  });

  it("does not show the banner when the draft flag is stale and the active page succeeded", () => {
    const draft = makeDraft("ПБ 78-12-8п 2");
    draft.metadata.ocr_verify_failed = true;
    const pages = [
      { ...makePage("a"), ocrVerifyFailed: false },
      { ...makePage("b"), ocrVerifyFailed: true },
    ];

    render(
      <PlateInputStep
        {...reviewBase}
        draft={draft}
        batchReviewText="ПБ 78-12-8п 2"
        normalizedText="ПБ 78-12-8п 2"
        pages={pages}
        activePageId="a"
      />,
    );

    expect(screen.queryByText(newCopy)).not.toBeInTheDocument();
    expect(screen.queryByText(oldCopy)).not.toBeInTheDocument();
  });

  it("falls back to the draft flag when there are no pages", () => {
    const draft = makeDraft("ПБ 78-12-8п 2");
    draft.metadata.ocr_verify_failed = true;

    render(
      <PlateInputStep
        {...reviewBase}
        draft={draft}
        batchReviewText="ПБ 78-12-8п 2"
        normalizedText="ПБ 78-12-8п 2"
        pages={[]}
        activePageId={null}
      />,
    );

    expect(screen.getByText(newCopy)).toBeInTheDocument();
  });
});

describe("PlateInputStep source image queue CTA", () => {
  const baseProps = {
    draft: makeDraft("ПБ 78-12-8п 2"),
    pendingBatchReview: false,
    sourceText: "",
    batchReviewText: "",
    normalizedText: "ПБ 78-12-8п 2",
    pages: [] as PageSource[],
    activePageId: null as string | null,
    recognizedImageUrl: null,
    recognizedImageName: null,
    errorMessage: null,
    isRecognizing: false,
    onTextChange: noop,
    onBatchReviewTextChange: noop,
    onAddFiles: noop,
    onRemovePage: noop,
    onSelectPage: noop,
    onRecognize: noop,
    onConfirmBatch: noop,
    onFinishPlates: noop,
    onReset: noop,
  };

  it("hides CTA when sourceQueue is empty", () => {
    render(<PlateInputStep {...baseProps} sourceQueue={[]} />);

    expect(screen.queryByRole("button", { name: /Исходные фото/i })).not.toBeInTheDocument();
  });

  it("shows CTA with count and opens Drawer on click", () => {
    render(
      <PlateInputStep
        {...baseProps}
        sourceQueue={[
          { id: "a", url: "blob:a", name: "a.png" },
          { id: "b", url: "blob:b", name: "b.png" },
        ]}
      />,
    );

    const cta = screen.getByRole("button", { name: "Исходные фото (2)" });
    expect(cta).toBeInTheDocument();

    fireEvent.click(cta);

    expect(screen.getByRole("dialog")).toBeInTheDocument();
    expect(screen.getByRole("img", { name: "a.png" })).toHaveAttribute("src", "blob:a");
  });
});

describe("PlateInputStep rerecognize button", () => {
  const onRecognize = vi.fn();
  const onRerecognize = vi.fn();

  const retryProps = {
    draft: makeDraft("ПБ 78-12-8п 2"),
    pendingBatchReview: true as const,
    sourceText: "",
    batchReviewText: "ПБ 78-12-8п 2",
    normalizedText: "ПБ 78-12-8п 2",
    recognizedImageUrl: "blob:photo",
    recognizedImageName: "a.png",
    errorMessage: null,
    isRecognizing: false,
    onAiInstructionChange: noop,
    onApplyAi: noop,
    onTextChange: noop,
    onBatchReviewTextChange: noop,
    onAddFiles: noop,
    onRemovePage: noop,
    onSelectPage: noop,
    onRecognize,
    onRerecognize,
    onConfirmBatch: noop,
    onFinishPlates: noop,
    onReset: noop,
  };

  beforeEach(() => {
    onRecognize.mockClear();
    onRerecognize.mockClear();
  });

  it("shows Перераспознать under selected photo when page is ready", () => {
    render(
      <PlateInputStep
        {...retryProps}
        pages={[makePage("a", "ready")]}
        activePageId="a"
      />,
    );

    expect(screen.getByRole("button", { name: "Перераспознать" })).toBeInTheDocument();
    expect(screen.getByText(/Ctrl \+ колёсико/)).toBeInTheDocument();
  });

  it("shows button for error and confirmed pages", () => {
    const { rerender } = render(
      <PlateInputStep
        {...retryProps}
        pages={[makePage("a", "error")]}
        activePageId="a"
      />,
    );
    expect(screen.getByRole("button", { name: "Перераспознать" })).toBeEnabled();

    rerender(
      <PlateInputStep
        {...retryProps}
        pages={[makePage("a", "confirmed")]}
        activePageId="a"
      />,
    );
    expect(screen.getByRole("button", { name: "Перераспознать" })).toBeEnabled();
  });

  it("disables button when page is running or pending", () => {
    const { rerender } = render(
      <PlateInputStep
        {...retryProps}
        pages={[makePage("a", "running")]}
        activePageId="a"
        isRerecognizing
      />,
    );
    expect(screen.getByRole("button", { name: "Распознавание..." })).toBeDisabled();

    rerender(
      <PlateInputStep
        {...retryProps}
        pages={[makePage("a", "pending")]}
        activePageId="a"
        isRerecognizing={false}
      />,
    );
    expect(screen.getByRole("button", { name: "Перераспознать" })).toBeDisabled();
  });

  it("click calls onRerecognize", () => {
    render(
      <PlateInputStep
        {...retryProps}
        pages={[makePage("a", "ready")]}
        activePageId="a"
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: "Перераспознать" }));
    expect(onRerecognize).toHaveBeenCalledTimes(1);
    expect(onRecognize).not.toHaveBeenCalled();
  });
});

describe("PlateInputStep live wide overlay", () => {
  const liveWideProps = {
    pendingBatchReview: true as const,
    sourceText: "",
    pages: [makePage("a")],
    activePageId: "a",
    recognizedImageUrl: "blob:photo",
    recognizedImageName: "a.png",
    errorMessage: null,
    isRecognizing: false,
    onAiInstructionChange: noop,
    onApplyAi: noop,
    onTextChange: noop,
    onBatchReviewTextChange: noop,
    onAddFiles: noop,
    onRemovePage: noop,
    onSelectPage: noop,
    onRecognize: noop,
    onConfirmBatch: noop,
    onFinishPlates: noop,
    onReset: noop,
    onWidePlateDecisionChange: noop,
    onApplyWidePlates: noop,
  };

  it("uses live editor 34-15-10п 15 instead of stale OCR 44-15-10п 5", () => {
    const draft = makeDraft("44-15-10п 5");
    draft.metadata.wide_plate_lines = [{ id: "stale", line: "44-15-10п 5", qty: 5 }];
    draft.metadata.wide_plates_resolved = false;

    render(
      <PlateInputStep
        {...liveWideProps}
        draft={draft}
        batchReviewText="34-15-10п 15"
        normalizedText="44-15-10п 5"
      />,
    );

    expect(screen.getByText("Нестандартная ширина")).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: /позиций требуют внимания/i }));
    expect(screen.getByText(/Количество:\s*15/)).toBeInTheDocument();
    expect(screen.queryByText(/Количество:\s*5/)).not.toBeInTheDocument();
    expect(screen.queryByText("44-15-10п 5")).not.toBeInTheDocument();
  });

  it("hides wide card and does not block confirm when width is 12 dm", () => {
    const draft = makeDraft("44-15-10п 5");
    draft.metadata.wide_plate_lines = [{ id: "stale", line: "44-15-10п 5", qty: 5 }];
    draft.metadata.wide_plates_resolved = false;

    render(
      <PlateInputStep
        {...liveWideProps}
        draft={draft}
        batchReviewText="34-12-10п 15"
        normalizedText="44-15-10п 5"
      />,
    );

    expect(screen.queryByText("Нестандартная ширина")).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Список верен" })).toBeEnabled();
  });
});

describe("PlateInputStep live OCR banners", () => {
  it("shows page ocrCorrections instead of stale draft hydrate", () => {
    const draft = makeDraft("two\nlines");
    draft.metadata.ocr_corrections = [{ action: "replaced", reason: "stale-2-line" }];
    draft.metadata.ocr_verify_failed = true;
    const page = {
      ...makePage("a"),
      ocrCorrections: [{ action: "replaced", reason: "fresh-full-list" }],
      ocrVerifyFailed: false,
    };

    render(
      <PlateInputStep
        draft={draft}
        pendingBatchReview
        sourceText=""
        batchReviewText="full list"
        normalizedText="two\nlines"
        pages={[page]}
        activePageId="a"
        recognizedImageUrl="blob:photo"
        recognizedImageName="a.png"
        errorMessage={null}
        isRecognizing={false}
        onTextChange={noop}
        onBatchReviewTextChange={noop}
        onAddFiles={noop}
        onRemovePage={noop}
        onSelectPage={noop}
        onRecognize={noop}
        onConfirmBatch={noop}
        onFinishPlates={noop}
        onReset={noop}
      />,
    );

    expect(screen.getByText(/fresh-full-list/)).toBeInTheDocument();
    expect(screen.queryByText(/stale-2-line/)).not.toBeInTheDocument();
    expect(screen.queryByText(OCR_VERIFY_FAILED_REVIEW_MESSAGE)).not.toBeInTheDocument();
  });
});
