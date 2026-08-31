import { act, renderHook } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";

import { planApplyAiSessionSync } from "@/features/commercial-offer/lib/applyAiSession";
import { useMultiPageRecognize } from "@/features/commercial-offer/hooks/useMultiPageRecognize";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makeOcrDraft = (lastBatchText: string): CommercialDraftDetails =>
  ({
    draft_id: "d1",
    metadata: {
      product_type: "plates",
      normalized_text: `page1\n${lastBatchText}`,
      plate_batches: [
        {
          source_type: "image",
          original_text: "",
          normalized_text: "page1",
          ocr_text: "",
          filename: "a.png",
        },
        {
          source_type: "image",
          original_text: "",
          normalized_text: lastBatchText,
          ocr_text: "",
          filename: "b.png",
        },
      ],
    },
  }) as CommercialDraftDetails;

const makeAiDraft = (text: string): CommercialDraftDetails =>
  ({
    draft_id: "d1",
    metadata: {
      product_type: "plates",
      source_type: "ai",
      normalized_text: text,
      input_text: text,
      plate_batches: [
        {
          source_type: "ai",
          original_text: "замени н на п",
          normalized_text: text,
          ocr_text: text,
          filename: "a.png",
        },
      ],
    },
  }) as CommercialDraftDetails;

describe("planApplyAiSessionSync (S20 / S1)", () => {
  it("preserves multi session and syncs active page text when hasStarted", () => {
    const draft = makeAiDraft("ПБ 78-12-8п 2");
    const plan = planApplyAiSessionSync({
      multiHasStarted: true,
      activePageId: "page-b",
      draft,
    });

    expect(plan.preserveMultiSession).toBe(true);
    expect(plan.nextActivePageText).toBe("ПБ 78-12-8п 2");
  });

  it("S1: uses draft normalized_text when last OCR batch is still stale", () => {
    const draft = {
      draft_id: "d1",
      metadata: {
        product_type: "plates",
        source_type: "ai",
        normalized_text: "ПБ 78-12-8п 2",
        input_text: "ПБ 78-12-8п 2",
        plate_batches: [
          {
            source_type: "image",
            original_text: "",
            normalized_text: "ПБ 78-12-8н 2",
            ocr_text: "ПБ 78-12-8н 2",
            filename: "a.png",
          },
        ],
      },
    } as CommercialDraftDetails;

    const plan = planApplyAiSessionSync({
      multiHasStarted: true,
      activePageId: "page-a",
      draft,
    });

    expect(plan.preserveMultiSession).toBe(true);
    expect(plan.nextActivePageText).toBe("ПБ 78-12-8п 2");
  });

  it("S1: non-multi Apply still returns review text for editor/store sync", () => {
    const plan = planApplyAiSessionSync({
      multiHasStarted: false,
      activePageId: null,
      draft: makeAiDraft("ПБ 60-12-8п 7"),
    });

    expect(plan.preserveMultiSession).toBe(false);
    expect(plan.nextActivePageText).toBe("ПБ 60-12-8п 7");
  });

  it("S20: 2+ ready pages stay intact after Apply AI sync (no reset)", async () => {
    const recognizePage = vi
      .fn()
      .mockResolvedValueOnce({ draft: makeOcrDraft("page1"), batchReviewText: "page1" })
      .mockResolvedValueOnce({ draft: makeOcrDraft("page2"), batchReviewText: "page2" });

    const { result } = renderHook(() => useMultiPageRecognize({ recognizePage }));

    act(() => {
      result.current.addFiles([
        new File(["a"], "a.png", { type: "image/png" }),
        new File(["b"], "b.png", { type: "image/png" }),
      ]);
    });
    await act(async () => {
      await result.current.start({ productType: "plates" });
    });

    expect(result.current.pages).toHaveLength(2);
    expect(result.current.hasStarted).toBe(true);

    const draftAfterAi = makeAiDraft("AI fixed");
    const plan = planApplyAiSessionSync({
      multiHasStarted: result.current.hasStarted,
      activePageId: result.current.activeId,
      draft: draftAfterAi,
    });

    expect(plan.preserveMultiSession).toBe(true);
    // Simulate wizard: sync text, do NOT reset
    if (plan.nextActivePageText !== null && result.current.activeId) {
      act(() => {
        result.current.updatePageText(result.current.activeId!, plan.nextActivePageText!);
      });
    }

    expect(result.current.pages).toHaveLength(2);
    expect(result.current.hasStarted).toBe(true);
    expect(result.current.pages.some((p) => p.batchReviewText === "AI fixed")).toBe(true);
  });
});
