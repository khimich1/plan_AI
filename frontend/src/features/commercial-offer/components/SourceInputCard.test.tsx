import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  SOURCE_LINT_ERROR_TITLE,
  SOURCE_LINT_PENDING_TITLE,
  SOURCE_LINT_RED_TITLE,
  SourceInputCard,
  resolveSourceSubmitDisabled,
} from "@/features/commercial-offer/components/SourceInputCard";
import { useSourceTextLint } from "@/features/commercial-offer/hooks/useSourceTextLint";
import {
  OCR_WAIT_MESSAGE,
  type PageSource,
  type PageStatus,
} from "@/features/commercial-offer/lib/multiPageSource";

vi.mock("@/features/commercial-offer/hooks/useSourceTextLint", () => ({
  useSourceTextLint: vi.fn(),
}));

const useLint = useSourceTextLint as unknown as ReturnType<typeof vi.fn>;

const makePage = (
  id: string,
  name = `${id}.png`,
  status: PageStatus = "pending",
): PageSource => ({
  id,
  file: new File(["x"], name, { type: "image/png" }),
  name,
  previewUrl: `blob:${id}`,
  status,
  batchReviewText: "",
});

const baseProps = {
  productType: "plates" as const,
  hasDraft: false,
  pages: [] as PageSource[],
  activePageId: null as string | null,
  isRecognizing: false,
  listLabel: "Список плит",
  placeholder: "ПБ 78-12-8п 2",
  emptySubtitle: "Вставьте текст списка плит или загрузите фото таблицы.",
  onTextChange: vi.fn(),
  onAddFiles: vi.fn(),
  onRemovePage: vi.fn(),
  onSelectPage: vi.fn(),
  onRecognize: vi.fn(),
};

afterEach(() => {
  cleanup();
});

beforeEach(() => {
  vi.clearAllMocks();
  useLint.mockReturnValue({ lines: [], isPending: false, isError: false });
});

describe("SourceInputCard button gate", () => {
  it("disables process button while lint is pending and sets title", () => {
    useLint.mockReturnValue({ lines: [], isPending: true, isError: false });

    render(<SourceInputCard {...baseProps} sourceText="ПБ 78-12-8п 2" />);

    const button = screen.getByRole("button", { name: "Обработать текст" });
    expect(button).toBeDisabled();
    expect(button).toHaveAttribute("title", SOURCE_LINT_PENDING_TITLE);
  });

  it("disables process button when a non-empty line is not ok", () => {
    useLint.mockReturnValue({
      isPending: false,
      isError: false,
      lines: [
        { index: 0, text: "ПБ 78-12-8п 2", empty: false, ok: true, reason_text: null },
        { index: 1, text: "плохо", empty: false, ok: false, reason_text: "не совпал формат" },
      ],
    });

    render(<SourceInputCard {...baseProps} sourceText={"ПБ 78-12-8п 2\nплохо"} />);

    const button = screen.getByRole("button", { name: "Обработать текст" });
    expect(button).toBeDisabled();
    expect(button).toHaveAttribute("title", SOURCE_LINT_RED_TITLE);
  });

  it("enables process button when all non-empty lines are ok", () => {
    useLint.mockReturnValue({
      isPending: false,
      isError: false,
      lines: [
        { index: 0, text: "ПБ 78-12-8п 2", empty: false, ok: true, reason_text: null },
        { index: 1, text: "", empty: true, ok: true, reason_text: null },
      ],
    });

    render(<SourceInputCard {...baseProps} sourceText={"ПБ 78-12-8п 2\n"} />);

    const button = screen.getByRole("button", { name: "Обработать текст" });
    expect(button).toBeEnabled();
    expect(button).not.toHaveAttribute("title");
  });

  it("does not lint-disable photo-only recognize when pages are present", () => {
    useLint.mockReturnValue({ lines: [], isPending: true, isError: false });

    render(
      <SourceInputCard {...baseProps} sourceText="" pages={[makePage("table", "table.png")]} activePageId="table" />,
    );

    const button = screen.getByRole("button", { name: "Распознать фото" });
    expect(button).toBeEnabled();
    expect(button).not.toHaveAttribute("title");
    expect(useLint).toHaveBeenCalledWith(expect.objectContaining({ enabled: false, text: "" }));
  });

  it("disables process button on lint error with a distinct title", () => {
    useLint.mockReturnValue({ lines: [], isPending: false, isError: true });

    render(<SourceInputCard {...baseProps} sourceText="ПБ 78-12-8п 2" />);

    const button = screen.getByRole("button", { name: "Обработать текст" });
    expect(button).toBeDisabled();
    expect(button).toHaveAttribute("title", SOURCE_LINT_ERROR_TITLE);
  });

  it("paints a red line overlay with the parse reason as title", () => {
    useLint.mockReturnValue({
      isPending: false,
      isError: false,
      lines: [{ index: 0, text: "плохо", empty: false, ok: false, reason_text: "не совпал формат строки" }],
    });

    render(<SourceInputCard {...baseProps} sourceText="плохо" />);

    expect(screen.getByTitle("не совпал формат строки")).toHaveTextContent("плохо");
  });

  it("shows gallery previews instead of file-name Alert", () => {
    render(
      <SourceInputCard
        {...baseProps}
        sourceText=""
        pages={[makePage("a", "a.png"), makePage("b", "b.png")]}
        activePageId="a"
      />,
    );

    expect(screen.queryByText(/Выбран файл/i)).not.toBeInTheDocument();
    expect(screen.getByRole("img", { name: "a.png" })).toBeInTheDocument();
    expect(screen.getByRole("img", { name: "b.png" })).toBeInTheDocument();
  });

  it("uses multiple file input on first-source card", () => {
    const { container } = render(<SourceInputCard {...baseProps} sourceText="" />);
    const input = container.querySelector('input[type="file"]');
    expect(input).toHaveAttribute("multiple");
  });

  it("does not show Распознавание label when pages are pending but isRecognizing is false", () => {
    render(
      <SourceInputCard
        {...baseProps}
        sourceText=""
        pages={[makePage("a", "a.png"), makePage("b", "b.png")]}
        activePageId="a"
        isRecognizing={false}
      />,
    );

    const button = screen.getByRole("button", { name: "Распознать фото" });
    expect(button).toBeEnabled();
    expect(screen.queryByRole("button", { name: "Распознавание..." })).not.toBeInTheDocument();
  });

  it("shows Распознавание and disables when isRecognizing is true with pages", () => {
    render(
      <SourceInputCard
        {...baseProps}
        sourceText=""
        pages={[makePage("a", "a.png")]}
        activePageId="a"
        isRecognizing
      />,
    );

    const button = screen.getByRole("button", { name: "Распознавание..." });
    expect(button).toBeDisabled();
  });
});

describe("SourceInputCard OCR wait banner", () => {
  it("does not show wait banner before recognition starts", () => {
    render(
      <SourceInputCard
        {...baseProps}
        sourceText=""
        pages={[makePage("a", "a.png"), makePage("b", "b.png")]}
        activePageId="a"
        recognitionStarted={false}
        isRecognizing={false}
      />,
    );

    expect(screen.queryByText(OCR_WAIT_MESSAGE)).not.toBeInTheDocument();
  });

  it("shows wait banner when started and no page is ready yet", () => {
    render(
      <SourceInputCard
        {...baseProps}
        sourceText=""
        pages={[makePage("a", "a.png", "running"), makePage("b", "b.png", "pending")]}
        activePageId="a"
        recognitionStarted
        isRecognizing
      />,
    );

    expect(screen.getByRole("status")).toHaveTextContent(OCR_WAIT_MESSAGE);
  });

  it("hides wait banner after the first page becomes ready", () => {
    render(
      <SourceInputCard
        {...baseProps}
        sourceText=""
        pages={[makePage("a", "a.png", "ready"), makePage("b", "b.png", "pending")]}
        activePageId="a"
        recognitionStarted
        isRecognizing
      />,
    );

    expect(screen.queryByText(OCR_WAIT_MESSAGE)).not.toBeInTheDocument();
  });
});

describe("resolveSourceSubmitDisabled footer gate", () => {
  const redGate = {
    sourceText: "плохо",
    canSubmit: false,
    blockReason: SOURCE_LINT_RED_TITLE,
  };

  it("disables append when lint reports red lines", () => {
    const result = resolveSourceSubmitDisabled("плохо", false, false, true, redGate);
    expect(result.disabled).toBe(true);
    expect(result.title).toBe(SOURCE_LINT_RED_TITLE);
  });

  it("treats a stale gate (text moved on) as pending", () => {
    const result = resolveSourceSubmitDisabled("ПБ 78-12-8п 2", false, false, true, redGate);
    expect(result.disabled).toBe(true);
    expect(result.title).toBe(SOURCE_LINT_PENDING_TITLE);
  });

  it("enables append when gate says the current text can submit", () => {
    const result = resolveSourceSubmitDisabled("ПБ 78-12-8п 2", false, false, true, {
      sourceText: "ПБ 78-12-8п 2",
      canSubmit: true,
      blockReason: undefined,
    });
    expect(result.disabled).toBe(false);
    expect(result.title).toBeUndefined();
  });

  it("does not apply lint when only a photo is selected", () => {
    const result = resolveSourceSubmitDisabled("", true, false, true, {
      sourceText: "",
      canSubmit: false,
      blockReason: SOURCE_LINT_PENDING_TITLE,
    });
    expect(result.disabled).toBe(false);
  });

  it("passes through the network-error title from the card gate", () => {
    const result = resolveSourceSubmitDisabled("ПБ 78-12-8п 2", false, false, true, {
      sourceText: "ПБ 78-12-8п 2",
      canSubmit: false,
      blockReason: SOURCE_LINT_ERROR_TITLE,
    });
    expect(result.disabled).toBe(true);
    expect(result.title).toBe(SOURCE_LINT_ERROR_TITLE);
  });
});
