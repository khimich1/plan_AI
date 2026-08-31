import { act, cleanup, renderHook, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { useMultiPageRecognize } from "@/features/commercial-offer/hooks/useMultiPageRecognize";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import { MAX_PAGES } from "@/features/commercial-offer/lib/multiPageSource";

const makeFile = (name: string) => new File([`content-${name}`], name, { type: "image/png" });

const makeDraft = (text: string, draftId = "draft-1"): CommercialDraftDetails =>
  ({
    draft_id: draftId,
    metadata: {
      product_type: "plates",
      normalized_text: text,
      plate_batches: [{ batch_index: 0, normalized_text: text, source_kind: "image" }],
    },
  }) as CommercialDraftDetails;

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

beforeEach(() => {
  const originalUrl = globalThis.URL;
  vi.spyOn(originalUrl, "createObjectURL").mockImplementation((file: Blob | MediaSource) => {
    const name = file instanceof File ? file.name : "blob";
    return `blob:${name}`;
  });
  vi.spyOn(originalUrl, "revokeObjectURL").mockImplementation(() => undefined);
});

describe("useMultiPageRecognize", () => {
  it("does not call OCR before start", async () => {
    const recognizePage = vi.fn();
    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    act(() => {
      result.current.addFiles([makeFile("a.png"), makeFile("b.png")]);
    });

    expect(result.current.pages).toHaveLength(2);
    expect(result.current.pages.every((page) => page.status === "pending")).toBe(true);
    expect(recognizePage).not.toHaveBeenCalled();
  });

  it("marks first page ready while second is still pending/running", async () => {
    let resolveSecond: ((value: { draft: CommercialDraftDetails; batchReviewText: string }) => void) | undefined;
    const recognizePage = vi
      .fn()
      .mockImplementationOnce(async () => ({
        draft: makeDraft("page1"),
        batchReviewText: "page1",
      }))
      .mockImplementationOnce(
        () =>
          new Promise((resolve) => {
            resolveSecond = resolve;
          }),
      );

    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    act(() => {
      result.current.addFiles([makeFile("a.png"), makeFile("b.png")]);
    });

    act(() => {
      void result.current.start({ productType: "plates" });
    });

    await waitFor(() => {
      expect(result.current.pages[0]?.status).toBe("ready");
    });
    expect(result.current.pages[0]?.batchReviewText).toBe("page1");
    expect(result.current.activeId).toBe(result.current.pages[0]?.id);
    expect(["pending", "running"]).toContain(result.current.pages[1]?.status);

    await act(async () => {
      resolveSecond?.({ draft: makeDraft("page2"), batchReviewText: "page2" });
    });

    await waitFor(() => {
      expect(result.current.pages[1]?.status).toBe("ready");
    });
    expect(recognizePage).toHaveBeenCalledTimes(2);
    expect(recognizePage.mock.calls[0]?.[0].isFirst).toBe(true);
    expect(recognizePage.mock.calls[1]?.[0].isFirst).toBe(false);
  });

  it("continues to next page after an error", async () => {
    const recognizePage = vi
      .fn()
      .mockRejectedValueOnce(new Error("ocr failed"))
      .mockResolvedValueOnce({ draft: makeDraft("page2"), batchReviewText: "page2" });

    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    act(() => {
      result.current.addFiles([makeFile("bad.png"), makeFile("good.png")]);
    });

    act(() => {
      void result.current.start({ productType: "plates" });
    });

    await waitFor(() => {
      expect(result.current.pages[0]?.status).toBe("error");
      expect(result.current.pages[1]?.status).toBe("ready");
    });
    expect(result.current.pages[0]?.errorMessage).toMatch(/ocr failed/i);
    expect(recognizePage).toHaveBeenCalledTimes(2);
  });

  it("remove ready/error then add file appends pending at the end without touching others", async () => {
    const recognizePage = vi
      .fn()
      .mockResolvedValueOnce({ draft: makeDraft("a"), batchReviewText: "a" })
      .mockResolvedValueOnce({ draft: makeDraft("b"), batchReviewText: "b" });

    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    act(() => {
      result.current.addFiles([makeFile("a.png"), makeFile("b.png")]);
    });

    act(() => {
      void result.current.start({ productType: "plates" });
    });

    await waitFor(() => {
      expect(result.current.pages.every((page) => page.status === "ready")).toBe(true);
    });

    const keptId = result.current.pages[1]!.id;
    const keptText = result.current.pages[1]!.batchReviewText;
    const removedId = result.current.pages[0]!.id;

    act(() => {
      result.current.remove(removedId);
      result.current.addFiles([makeFile("c.png")]);
    });

    expect(result.current.pages).toHaveLength(2);
    expect(result.current.pages[0]?.id).toBe(keptId);
    expect(result.current.pages[0]?.status).toBe("ready");
    expect(result.current.pages[0]?.batchReviewText).toBe(keptText);
    expect(result.current.pages[1]?.name).toBe("c.png");
    expect(["pending", "running"]).toContain(result.current.pages[1]?.status);
  });

  it("enforces soft-cap of 12 pages", () => {
    const recognizePage = vi.fn();
    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    act(() => {
      result.current.addFiles(Array.from({ length: MAX_PAGES }, (_, i) => makeFile(`p${i}.png`)));
    });
    expect(result.current.pages).toHaveLength(MAX_PAGES);

    act(() => {
      result.current.addFiles([makeFile("overflow.png")]);
    });

    expect(result.current.pages).toHaveLength(MAX_PAGES);
    expect(result.current.softCapMessage).toMatch(/12/);
    expect(recognizePage).not.toHaveBeenCalled();
  });

  it("confirmActive marks page confirmed and advances to next ready", async () => {
    const recognizePage = vi
      .fn()
      .mockResolvedValueOnce({ draft: makeDraft("a"), batchReviewText: "a" })
      .mockResolvedValueOnce({ draft: makeDraft("b"), batchReviewText: "b" });

    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    act(() => {
      result.current.addFiles([makeFile("a.png"), makeFile("b.png")]);
    });

    act(() => {
      void result.current.start({ productType: "plates" });
    });

    await waitFor(() => {
      expect(result.current.pages.every((page) => page.status === "ready")).toBe(true);
    });

    const firstId = result.current.pages[0]!.id;
    const secondId = result.current.pages[1]!.id;

    act(() => {
      result.current.setActive(firstId);
    });

    let confirmResult: { allConfirmed: boolean; nextId: string | null } | undefined;
    act(() => {
      confirmResult = result.current.confirmActive();
    });

    expect(result.current.pages[0]?.status).toBe("confirmed");
    expect(confirmResult?.nextId).toBe(secondId);
    expect(result.current.activeId).toBe(secondId);
    expect(confirmResult?.allConfirmed).toBe(false);
  });

  // S12: remove all pages after start resets multi-session
  it("resets hasStarted when all pages are removed after start", async () => {
    const recognizePage = vi.fn().mockResolvedValue({
      draft: makeDraft("a"),
      batchReviewText: "a",
    });
    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    act(() => {
      result.current.addFiles([makeFile("a.png")]);
    });
    act(() => {
      void result.current.start({ productType: "plates" });
    });
    await waitFor(() => {
      expect(result.current.pages[0]?.status).toBe("ready");
      expect(result.current.hasStarted).toBe(true);
    });

    act(() => {
      result.current.remove(result.current.pages[0]!.id);
    });

    expect(result.current.pages).toHaveLength(0);
    expect(result.current.hasStarted).toBe(false);
    expect(result.current.allConfirmed).toBe(false);
  });

  // S15: remove/addFiles return next pages (no stale React state)
  it("remove and addFiles return the next pages list synchronously", () => {
    const recognizePage = vi.fn();
    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    let afterAdd: ReturnType<typeof result.current.addFiles> = [];
    act(() => {
      afterAdd = result.current.addFiles([makeFile("a.png"), makeFile("b.png")]);
    });
    expect(afterAdd).toHaveLength(2);
    expect(afterAdd.map((p) => p.name)).toEqual(["a.png", "b.png"]);

    const removeId = afterAdd[0]!.id;
    let afterRemove: ReturnType<typeof result.current.remove> = [];
    act(() => {
      afterRemove = result.current.remove(removeId);
    });
    expect(afterRemove).toHaveLength(1);
    expect(afterRemove[0]?.name).toBe("b.png");

    let afterAddTail: ReturnType<typeof result.current.addFiles> = [];
    act(() => {
      afterAddTail = result.current.addFiles([makeFile("c.png")]);
    });
    expect(afterAddTail.map((p) => p.name)).toEqual(["b.png", "c.png"]);
  });

  // S16: one object URL per page; revoke once on remove
  it("reuses previewUrl on ready and revokes a single object URL on remove", async () => {
    const createPreviewUrl = vi.fn((file: File) => `blob:${file.name}`);
    const revokePreviewUrl = vi.fn();
    const recognizePage = vi.fn().mockResolvedValue({
      draft: makeDraft("a"),
      batchReviewText: "a",
    });
    const { result } = renderHook(() =>
      useMultiPageRecognize({ recognizePage, createPreviewUrl, revokePreviewUrl }),
    );

    act(() => {
      result.current.addFiles([makeFile("a.png")]);
    });
    expect(createPreviewUrl).toHaveBeenCalledTimes(1);

    act(() => {
      void result.current.start({ productType: "plates" });
    });
    await waitFor(() => {
      expect(result.current.pages[0]?.status).toBe("ready");
    });

    expect(createPreviewUrl).toHaveBeenCalledTimes(1);
    expect(result.current.pages[0]?.recognizedImageUrl).toBe(result.current.pages[0]?.previewUrl);

    act(() => {
      result.current.remove(result.current.pages[0]!.id);
    });
    expect(revokePreviewUrl).toHaveBeenCalledTimes(1);
    expect(revokePreviewUrl).toHaveBeenCalledWith("blob:a.png");
  });

  // S18: pending pages before start must not count as recognizing
  it("isRecognizing stays false after addFiles until start", async () => {
    let resolveOcr: ((value: { draft: CommercialDraftDetails; batchReviewText: string }) => void) | undefined;
    const recognizePage = vi.fn(
      () =>
        new Promise<{ draft: CommercialDraftDetails; batchReviewText: string }>((resolve) => {
          resolveOcr = resolve;
        }),
    );
    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    act(() => {
      result.current.addFiles([makeFile("a.png"), makeFile("b.png")]);
    });

    expect(result.current.hasStarted).toBe(false);
    expect(result.current.hasPendingOrRunning).toBe(true);
    expect(result.current.isRecognizing).toBe(false);

    act(() => {
      void result.current.start({ productType: "plates" });
    });

    await waitFor(() => {
      expect(result.current.isRecognizing).toBe(true);
      expect(result.current.pages[0]?.status).toBe("running");
    });

    await act(async () => {
      resolveOcr?.({ draft: makeDraft("a"), batchReviewText: "a" });
    });

    await waitFor(() => {
      expect(result.current.pages[0]?.status).toBe("ready");
      expect(result.current.isRecognizing).toBe(true); // page b still pending/running
    });

    await act(async () => {
      resolveOcr?.({ draft: makeDraft("b"), batchReviewText: "b" });
    });

    await waitFor(() => {
      expect(result.current.pages.every((page) => page.status === "ready")).toBe(true);
    });
    expect(result.current.isRecognizing).toBe(false);
  });

  it("stores ocrVerifyFailed per page from that page's OCR metadata", async () => {
    const recognizePage = vi
      .fn()
      .mockImplementationOnce(async () => ({
        draft: makeDraft("page1"),
        batchReviewText: "page1",
      }))
      .mockImplementationOnce(async () => {
        const draft = makeDraft("page2");
        draft.metadata.ocr_verify_failed = true;
        return { draft, batchReviewText: "page2" };
      });

    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    act(() => {
      result.current.addFiles([makeFile("a.png"), makeFile("b.png")]);
    });

    act(() => {
      void result.current.start({ productType: "plates" });
    });

    await waitFor(() => {
      expect(result.current.pages.every((page) => page.status === "ready")).toBe(true);
    });

    expect(result.current.pages[0]?.ocrVerifyFailed).toBe(false);
    expect(result.current.pages[1]?.ocrVerifyFailed).toBe(true);
  });
});
