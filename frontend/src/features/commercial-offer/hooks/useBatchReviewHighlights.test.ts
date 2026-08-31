import { renderHook } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";

import { useBatchReviewHighlights } from "@/features/commercial-offer/hooks/useBatchReviewHighlights";
import { useSourceTextLint } from "@/features/commercial-offer/hooks/useSourceTextLint";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

vi.mock("@/features/commercial-offer/hooks/useSourceTextLint", () => ({
  useSourceTextLint: vi.fn(),
}));

const useLint = useSourceTextLint as unknown as ReturnType<typeof vi.fn>;

const draft = { draft_id: "d1", metadata: { unparsed_lines: [] } } as CommercialDraftDetails;

describe("useBatchReviewHighlights", () => {
  it("merges live lint rejects into the review highlight map", () => {
    useLint.mockReturnValue({
      isPending: false,
      isError: false,
      lines: [{ index: 0, text: "плохо", empty: false, ok: false, reason_text: "не совпал формат строки" }],
    });

    const { result } = renderHook(() =>
      useBatchReviewHighlights({
        text: "плохо",
        productType: "plates",
        draft,
        enabled: true,
      }),
    );

    expect(result.current.get(0)?.kind).toBe("unparsed");
  });

  it("does not highlight when lint is disabled", () => {
    useLint.mockReturnValue({
      isPending: false,
      isError: false,
      lines: [{ index: 0, text: "плохо", empty: false, ok: false, reason_text: "x" }],
    });

    const { result } = renderHook(() =>
      useBatchReviewHighlights({
        text: "плохо",
        productType: "plates",
        draft,
        enabled: false,
      }),
    );

    expect(result.current.size).toBe(0);
  });
});
