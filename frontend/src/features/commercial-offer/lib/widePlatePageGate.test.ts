import { describe, expect, it } from "vitest";

import {
  photoQueueHasUndecidedWidePlate,
  undecidedWideKeys,
  wideKeysInText,
} from "@/features/commercial-offer/lib/widePlatePageGate";

describe("wide plate page gate", () => {
  it("treats a known wide mark as undecided until this photo applies", () => {
    expect(undecidedWideKeys("ПБ 56-15-8 20", [])).toEqual(wideKeysInText("ПБ 56-15-8 20"));
  });

  it("does not reopen a mark this photo already applied", () => {
    const [key] = wideKeysInText("ПБ 56-15-8 20");
    expect(undecidedWideKeys("ПБ 56-15-8 20", [key!])).toEqual([]);
  });

  it("keeps the second photo open when only the first photo applied the same mark", () => {
    const pages = [
      { id: "first", batchReviewText: "ПБ 56-15-8 20" },
      { id: "second", batchReviewText: "ПБ 56-15-8 4" },
    ];
    const settled = { first: wideKeysInText(pages[0]!.batchReviewText) };
    expect(photoQueueHasUndecidedWidePlate(pages, settled)).toBe(true);
    expect(undecidedWideKeys(pages[1]!.batchReviewText, settled.second ?? [])).toHaveLength(1);
  });
});
