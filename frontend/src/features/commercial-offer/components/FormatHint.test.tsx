import { act, cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { FormatHint } from "@/features/commercial-offer/components/FormatHint";
import { FORMAT_HINT_SHARED } from "@/features/commercial-offer/lib/productTypeConfig";

const SAMPLE = "ПБ 78-12-8п 2\n71-12-8 3\nПБ 66-12-8п 4";
const EXPLANATION = "Марка, затем количество штук. Нагрузка входит в марку (`8п`).";

const samplePre = () => {
  const pre = document.querySelector("pre");
  expect(pre).not.toBeNull();
  expect(pre?.textContent).toBe(SAMPLE);
  return pre as HTMLPreElement;
};

afterEach(() => {
  cleanup();
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

beforeEach(() => {
  vi.restoreAllMocks();
});

describe("FormatHint", () => {
  it("starts collapsed: toggle shows ▸ Подсказка and expanded block is absent", () => {
    render(<FormatHint explanation={EXPLANATION} sample={SAMPLE} />);

    const toggle = screen.getByRole("button", { name: "▸ Подсказка" });
    expect(toggle).toHaveAttribute("aria-expanded", "false");
    expect(screen.queryByText(FORMAT_HINT_SHARED)).not.toBeInTheDocument();
    expect(screen.queryByText(EXPLANATION)).not.toBeInTheDocument();
    expect(document.querySelector("pre")).toBeNull();
    expect(screen.queryByRole("button", { name: "Скопировать образец" })).not.toBeInTheDocument();
  });

  it("expands on click to show shared shell, explanation, and sample", () => {
    render(<FormatHint explanation={EXPLANATION} sample={SAMPLE} />);

    fireEvent.click(screen.getByRole("button", { name: "▸ Подсказка" }));

    const toggle = screen.getByRole("button", { name: "▾ Подсказка" });
    expect(toggle).toHaveAttribute("aria-expanded", "true");
    expect(screen.getByText(FORMAT_HINT_SHARED)).toBeInTheDocument();
    expect(screen.getByText(EXPLANATION)).toBeInTheDocument();
    samplePre();
    expect(screen.getByRole("button", { name: "Скопировать образец" })).toBeInTheDocument();
  });

  it("collapses again on a second click", () => {
    render(<FormatHint explanation={EXPLANATION} sample={SAMPLE} />);

    fireEvent.click(screen.getByRole("button", { name: "▸ Подсказка" }));
    fireEvent.click(screen.getByRole("button", { name: "▾ Подсказка" }));

    expect(screen.getByRole("button", { name: "▸ Подсказка" })).toHaveAttribute(
      "aria-expanded",
      "false",
    );
    expect(screen.queryByText(FORMAT_HINT_SHARED)).not.toBeInTheDocument();
    expect(screen.queryByText(EXPLANATION)).not.toBeInTheDocument();
  });

  it("copies only the sample via clipboard.writeText", async () => {
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", {
      ...navigator,
      clipboard: { writeText },
    });

    render(<FormatHint explanation={EXPLANATION} sample={SAMPLE} />);
    fireEvent.click(screen.getByRole("button", { name: "▸ Подсказка" }));
    fireEvent.click(screen.getByRole("button", { name: "Скопировать образец" }));

    await waitFor(() => {
      expect(writeText).toHaveBeenCalledTimes(1);
    });
    expect(writeText).toHaveBeenCalledWith(SAMPLE);
    await waitFor(() => {
      expect(screen.getByRole("button", { name: "Скопировано" })).toBeInTheDocument();
    });
  });

  it("resets copy button label after a short success delay", async () => {
    vi.useFakeTimers();
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", {
      ...navigator,
      clipboard: { writeText },
    });

    render(<FormatHint explanation={EXPLANATION} sample={SAMPLE} />);
    fireEvent.click(screen.getByRole("button", { name: "▸ Подсказка" }));

    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Скопировать образец" }));
      await Promise.resolve();
    });

    expect(screen.getByRole("button", { name: "Скопировано" })).toBeInTheDocument();

    await act(async () => {
      await vi.advanceTimersByTimeAsync(2000);
    });

    expect(screen.getByRole("button", { name: "Скопировать образец" })).toBeInTheDocument();
  });

  it("shows failure text when clipboard.writeText rejects", async () => {
    const writeText = vi.fn().mockRejectedValue(new Error("denied"));
    vi.stubGlobal("navigator", {
      ...navigator,
      clipboard: { writeText },
    });

    render(<FormatHint explanation={EXPLANATION} sample={SAMPLE} />);
    fireEvent.click(screen.getByRole("button", { name: "▸ Подсказка" }));
    fireEvent.click(screen.getByRole("button", { name: "Скопировать образец" }));

    await waitFor(() => {
      expect(screen.getByText("Не удалось скопировать")).toBeInTheDocument();
    });
    samplePre();
  });

  it("has no list-field change prop and leaves the sample as display-only text", async () => {
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", {
      ...navigator,
      clipboard: { writeText },
    });

    const { container } = render(<FormatHint explanation={EXPLANATION} sample={SAMPLE} />);
    fireEvent.click(screen.getByRole("button", { name: "▸ Подсказка" }));
    fireEvent.click(screen.getByRole("button", { name: "Скопировать образец" }));

    await waitFor(() => {
      expect(writeText).toHaveBeenCalledWith(SAMPLE);
    });
    expect(container.querySelector("textarea")).toBeNull();
    samplePre();
  });
});
