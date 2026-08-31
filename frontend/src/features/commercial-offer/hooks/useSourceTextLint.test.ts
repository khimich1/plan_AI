import { act, cleanup, renderHook } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import { SOURCE_TEXT_LINT_DEBOUNCE_MS, useSourceTextLint } from "./useSourceTextLint";

vi.mock("@/features/commercial-offer/api/commercialOfferApi", () => ({
  commercialOfferApi: {
    parseSource: vi.fn(),
  },
}));

const parseSource = commercialOfferApi.parseSource as unknown as ReturnType<typeof vi.fn>;

const okResponse = {
  product_type: "plates" as const,
  lines: [{ index: 0, text: "ПБ 78-12-8п 2", empty: false, ok: true, reason_text: null }],
  unparsed_lines: [],
};

afterEach(() => {
  cleanup();
  vi.useRealTimers();
  vi.clearAllMocks();
});

beforeEach(() => {
  parseSource.mockReset();
});

describe("useSourceTextLint", () => {
  it("does not fetch for empty text and is not pending", () => {
    vi.useFakeTimers();
    const { result } = renderHook(() =>
      useSourceTextLint({ text: "  \n", productType: "plates", enabled: true }),
    );

    expect(result.current.isPending).toBe(false);
    expect(result.current.lines).toEqual([]);
    act(() => {
      vi.advanceTimersByTime(SOURCE_TEXT_LINT_DEBOUNCE_MS + 50);
    });
    expect(parseSource).not.toHaveBeenCalled();
  });

  it("sets pending immediately and fetches once after debounce", async () => {
    vi.useFakeTimers();
    parseSource.mockResolvedValue(okResponse);

    const { result, rerender } = renderHook(
      ({ text }: { text: string }) =>
        useSourceTextLint({ text, productType: "plates", enabled: true }),
      { initialProps: { text: "ПБ" } },
    );

    expect(result.current.isPending).toBe(true);
    rerender({ text: "ПБ 78-12-8п 2" });
    expect(result.current.isPending).toBe(true);

    await act(async () => {
      vi.advanceTimersByTime(SOURCE_TEXT_LINT_DEBOUNCE_MS - 1);
    });
    expect(parseSource).not.toHaveBeenCalled();

    await act(async () => {
      await vi.advanceTimersByTimeAsync(1);
    });
    expect(parseSource).toHaveBeenCalledTimes(1);
    expect(parseSource).toHaveBeenCalledWith(
      { text: "ПБ 78-12-8п 2", productType: "plates" },
      expect.objectContaining({ signal: expect.any(AbortSignal) }),
    );
    expect(result.current.isPending).toBe(false);
    expect(result.current.lines[0]?.ok).toBe(true);
  });

  it("discards a stale response from an older request", async () => {
    vi.useFakeTimers();
    let resolveFirst: ((value: typeof okResponse) => void) | undefined;
    parseSource.mockImplementationOnce(
      () =>
        new Promise((resolve) => {
          resolveFirst = resolve;
        }),
    );
    parseSource.mockResolvedValueOnce({
      product_type: "plates",
      lines: [{ index: 0, text: "ПБ 66-12-8п 4", empty: false, ok: true, reason_text: null }],
      unparsed_lines: [],
    });

    const { result, rerender } = renderHook(
      ({ text }: { text: string }) =>
        useSourceTextLint({ text, productType: "plates", enabled: true }),
      { initialProps: { text: "first" } },
    );

    await act(async () => {
      await vi.advanceTimersByTimeAsync(SOURCE_TEXT_LINT_DEBOUNCE_MS);
    });
    rerender({ text: "second" });
    await act(async () => {
      await vi.advanceTimersByTimeAsync(SOURCE_TEXT_LINT_DEBOUNCE_MS);
    });

    await act(async () => {
      resolveFirst?.({
        product_type: "plates",
        lines: [{ index: 0, text: "stale", empty: false, ok: false, reason_text: "stale" }],
        unparsed_lines: ["stale"],
      });
    });

    expect(result.current.lines[0]?.text).toBe("ПБ 66-12-8п 4");
    expect(result.current.lines[0]?.ok).toBe(true);
  });

  it("does not fetch when enabled is false (photo-only)", () => {
    vi.useFakeTimers();
    const { result } = renderHook(() =>
      useSourceTextLint({ text: "ПБ 78-12-8п 2", productType: "plates", enabled: false }),
    );

    expect(result.current.isPending).toBe(false);
    act(() => {
      vi.advanceTimersByTime(SOURCE_TEXT_LINT_DEBOUNCE_MS + 50);
    });
    expect(parseSource).not.toHaveBeenCalled();
  });

  it("sets isError when parseSource rejects and does not treat abort as error", async () => {
    vi.useFakeTimers();
    parseSource.mockRejectedValue(new Error("network down"));

    const { result } = renderHook(() =>
      useSourceTextLint({ text: "ПБ 78-12-8п 2", productType: "plates", enabled: true }),
    );

    await act(async () => {
      await vi.advanceTimersByTimeAsync(SOURCE_TEXT_LINT_DEBOUNCE_MS);
    });

    expect(result.current.isPending).toBe(false);
    expect(result.current.isError).toBe(true);
  });
});

