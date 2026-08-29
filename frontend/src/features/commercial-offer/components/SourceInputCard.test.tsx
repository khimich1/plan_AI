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

vi.mock("@/features/commercial-offer/hooks/useSourceTextLint", () => ({
  useSourceTextLint: vi.fn(),
}));

const useLint = useSourceTextLint as unknown as ReturnType<typeof vi.fn>;

const baseProps = {
  productType: "plates" as const,
  hasDraft: false,
  selectedImageName: null as string | null,
  isRecognizing: false,
  listLabel: "Список плит",
  placeholder: "ПБ 78-12-8п 2",
  emptySubtitle: "Вставьте текст списка плит или загрузите фото таблицы.",
  onTextChange: vi.fn(),
  onFileChange: vi.fn(),
  onImagePaste: vi.fn(),
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

  it("does not lint-disable photo-only recognize", () => {
    useLint.mockReturnValue({ lines: [], isPending: true, isError: false });

    render(
      <SourceInputCard {...baseProps} sourceText="" selectedImageName="table.png" />,
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
});

describe("resolveSourceSubmitDisabled footer gate", () => {
  const redGate = {
    sourceText: "плохо",
    canSubmit: false,
    blockReason: SOURCE_LINT_RED_TITLE,
  };

  it("disables append when lint reports red lines", () => {
    const result = resolveSourceSubmitDisabled("плохо", null, false, true, redGate);
    expect(result.disabled).toBe(true);
    expect(result.title).toBe(SOURCE_LINT_RED_TITLE);
  });

  it("treats a stale gate (text moved on) as pending", () => {
    const result = resolveSourceSubmitDisabled("ПБ 78-12-8п 2", null, false, true, redGate);
    expect(result.disabled).toBe(true);
    expect(result.title).toBe(SOURCE_LINT_PENDING_TITLE);
  });

  it("enables append when gate says the current text can submit", () => {
    const result = resolveSourceSubmitDisabled("ПБ 78-12-8п 2", null, false, true, {
      sourceText: "ПБ 78-12-8п 2",
      canSubmit: true,
      blockReason: undefined,
    });
    expect(result.disabled).toBe(false);
    expect(result.title).toBeUndefined();
  });

  it("does not apply lint when only a photo is selected", () => {
    const result = resolveSourceSubmitDisabled("", "table.png", false, true, {
      sourceText: "",
      canSubmit: false,
      blockReason: SOURCE_LINT_PENDING_TITLE,
    });
    expect(result.disabled).toBe(false);
  });

  it("passes through the network-error title from the card gate", () => {
    const result = resolveSourceSubmitDisabled("ПБ 78-12-8п 2", null, false, true, {
      sourceText: "ПБ 78-12-8п 2",
      canSubmit: false,
      blockReason: SOURCE_LINT_ERROR_TITLE,
    });
    expect(result.disabled).toBe(true);
    expect(result.title).toBe(SOURCE_LINT_ERROR_TITLE);
  });
});
