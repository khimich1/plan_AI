import { describe, expect, it } from "vitest";

import {
  applyWideResolveToActivePage,
  applyWideResolveToQueuePages,
  resolveAppliedQueuePageText,
  syncWidePlateQueuePages,
} from "@/features/commercial-offer/lib/syncWidePlateQueuePages";

describe("syncWidePlateQueuePages", () => {
  it("pairs each page with its own batch and keeps the inactive page", () => {
    const result = syncWidePlateQueuePages({
      pages: [
        { id: "page-1", batchReviewText: "old wide mark" },
        { id: "page-2", batchReviewText: "old second" },
      ],
      activePageId: "page-1",
      draft: {
        metadata: {
          normalized_text: "ПБ 60-12-8 2\nПБ 76-12-8 13\nПБ 76-3.0-8 13",
          plate_batches: [
            { normalized_text: "ПБ 60-12-8 2" },
            { normalized_text: "ПБ 76-12-8 13\nПБ 76-3.0-8 13" },
          ],
        },
      },
    });

    expect(result.ok).toBe(true);
    if (!result.ok) {
      return;
    }
    expect(result.pages).toEqual([
      { id: "page-1", batchReviewText: "ПБ 60-12-8 2" },
      { id: "page-2", batchReviewText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13" },
    ]);
    expect(result.pages[1]?.batchReviewText).not.toBe(result.activePageText);
    expect(result.activePageText).toBe("ПБ 60-12-8 2");
    expect(result.activePageText).not.toBe("ПБ 60-12-8 2\nПБ 76-12-8 13\nПБ 76-3.0-8 13");
  });

  it("returns failure when plate batch count differs from page count", () => {
    const result = syncWidePlateQueuePages({
      pages: [
        { id: "page-1", batchReviewText: "keep-me" },
        { id: "page-2", batchReviewText: "keep-me-too" },
      ],
      activePageId: "page-2",
      draft: {
        metadata: {
          plate_batches: [{ normalized_text: "only-last-batch" }],
        },
      },
    });

    expect(result).toEqual({
      ok: false,
      message: expect.any(String),
    });
    expect(result).not.toHaveProperty("pages");
  });

  it("writes a four-line replacement onto the first page and leaves the second", () => {
    const replacement = [
      "ПБ 76-12-8 13",
      "ПБ 76-3.0-8 13",
      "ПБ 74-12-8 4",
      "ПБ 74-3.0-8 4",
    ].join("\n");
    const second = "ПБ 58-12-8 2";
    const result = syncWidePlateQueuePages({
      pages: [
        { id: "page-1", batchReviewText: "ПБ 76-15-8 13" },
        { id: "page-2", batchReviewText: second },
      ],
      activePageId: "page-1",
      draft: {
        metadata: {
          plate_batches: [{ normalized_text: replacement }, { normalized_text: second }],
        },
      },
    });

    expect(result.ok).toBe(true);
    if (!result.ok) {
      return;
    }
    expect(result.pages).toHaveLength(2);
    expect(result.pages[0]?.batchReviewText).toBe(replacement);
    expect(result.pages[0]?.batchReviewText).not.toContain("76-15-8");
    expect(result.pages[1]?.batchReviewText).toBe(second);
  });

  it("keeps an edited page and applies every replacement line", () => {
    const replacement = [
      "ПБ 76-12-8 13",
      "ПБ 76-3.0-8 13",
      "ПБ 74-12-8 4",
      "ПБ 74-3.0-8 4",
    ].join("\n");
    const text = resolveAppliedQueuePageText({
      pageText: "ПБ 76-15-8 13\nПБ 60-12-8 5",
      batchBefore: "ПБ 76-15-8 13\nПБ 60-12-8 1",
      batchAfter: `${replacement}\nПБ 60-12-8 1`,
      decisions: [
        {
          sourceLine: "ПБ 76-15-8п 13",
          action: "replace",
          replacementText: replacement,
        },
      ],
    });

    expect(text).toBe(`${replacement}\nПБ 60-12-8 5`);
    expect(text).not.toContain("76-15-8");
    expect(text).not.toContain("ПБ 60-12-8 1");
  });

  it("does not copy an edited page onto the next photo", () => {
    const replacement = [
      "ПБ 76-12-8 13",
      "ПБ 76-3.0-8 13",
      "ПБ 74-12-8 4",
      "ПБ 74-3.0-8 4",
    ].join("\n");
    const second = "ПБ 58-12-8 2";
    const pages = applyWideResolveToQueuePages({
      pages: [
        { id: "page-1", batchReviewText: "ПБ 76-15-8 13\nПБ 60-12-8 5" },
        { id: "page-2", batchReviewText: second },
      ],
      batchesBefore: [
        { normalized_text: "ПБ 76-15-8 13\nПБ 60-12-8 1" },
        { normalized_text: second },
      ],
      resolvedPages: [
        { id: "page-1", batchReviewText: `${replacement}\nПБ 60-12-8 1` },
        { id: "page-2", batchReviewText: second },
      ],
      decisions: [
        {
          sourceLine: "ПБ 76-15-8п 13",
          action: "replace",
          replacementText: replacement,
        },
      ],
    });

    expect(pages[0]?.batchReviewText).toBe(`${replacement}\nПБ 60-12-8 5`);
    expect(pages[0]?.batchReviewText).not.toContain("76-15-8");
    expect(pages[1]?.batchReviewText).toBe(second);
  });

  it("uses the resolved batch when the page text was not edited", () => {
    const before = "ПБ 76-15-8 13";
    const after = "ПБ 76-12-8 13\nПБ 76-3.0-8 13\nПБ 74-12-8 4\nПБ 74-3.0-8 4";
    expect(
      resolveAppliedQueuePageText({
        pageText: before,
        batchBefore: before,
        batchAfter: after,
        decisions: [
          {
            sourceLine: "ПБ 76-15-8п 13",
            action: "replace",
            replacementText: "ignored",
          },
        ],
      }),
    ).toBe(after);
  });

  it("writes the replacement only onto the open photo", () => {
    const second = "ПБ 56-15-8 20\nПБ 45-15-10 3";
    const pages = applyWideResolveToActivePage({
      pages: [
        { id: "page-1", batchReviewText: "ПБ 76-15-8 13" },
        { id: "page-2", batchReviewText: second },
      ],
      activePageId: "page-1",
      batchesBefore: [{ normalized_text: "ПБ 76-15-8 13" }, { normalized_text: second }],
      resolvedPages: [
        { id: "page-1", batchReviewText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13" },
        { id: "page-2", batchReviewText: "ПБ 56-12-8 20\nПБ 45-12-10 3" },
      ],
      decisions: [
        {
          sourceLine: "ПБ 76-15-8 13",
          action: "replace",
          replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
        },
      ],
    });

    expect(pages[0]?.batchReviewText).toBe("ПБ 76-12-8 13\nПБ 76-3.0-8 13");
    expect(pages[1]?.batchReviewText).toBe(second);
  });
});
