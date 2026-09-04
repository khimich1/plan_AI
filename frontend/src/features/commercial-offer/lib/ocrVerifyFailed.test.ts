import { describe, expect, it } from "vitest";

import {
  resolveActivePageOcrCorrections,
  resolveActivePageOcrVerifyFailed,
} from "@/features/commercial-offer/lib/ocrVerifyFailed";

describe("resolveActivePageOcrVerifyFailed", () => {
  it("uses draft flag when there are no pages (legacy single review)", () => {
    expect(resolveActivePageOcrVerifyFailed([], null, true)).toBe(true);
    expect(resolveActivePageOcrVerifyFailed([], null, false)).toBe(false);
  });

  it("uses the active page flag and ignores a stale draft flag", () => {
    const pages = [
      { id: "a", ocrVerifyFailed: false },
      { id: "b", ocrVerifyFailed: true },
    ];
    expect(resolveActivePageOcrVerifyFailed(pages, "a", true)).toBe(false);
    expect(resolveActivePageOcrVerifyFailed(pages, "b", false)).toBe(true);
  });

  it("is false when pages exist but the active page is missing or unset", () => {
    expect(resolveActivePageOcrVerifyFailed([{ id: "a", ocrVerifyFailed: true }], null, true)).toBe(
      false,
    );
    expect(resolveActivePageOcrVerifyFailed([{ id: "a", ocrVerifyFailed: true }], "missing", true)).toBe(
      false,
    );
  });
});

describe("resolveActivePageOcrCorrections", () => {
  it("uses draft fallback when there are no pages", () => {
    const draftFallback = [{ action: "replaced" as const, reason: "old" }];
    expect(resolveActivePageOcrCorrections([], null, draftFallback)).toEqual(draftFallback);
  });

  it("lets the active page corrections beat a stale draft hydrate", () => {
    const draftFallback = [{ action: "replaced", reason: "stale-2-line" }];
    const pages = [
      {
        id: "a",
        ocrCorrections: [
          { action: "replaced", reason: "fresh-full-list" },
        ],
      },
    ];
    expect(resolveActivePageOcrCorrections(pages, "a", draftFallback)).toEqual([
      { action: "replaced", reason: "fresh-full-list" },
    ]);
  });

  it("uses an empty page array after retry instead of stale draft corrections", () => {
    const draftFallback = [{ action: "replaced", reason: "stale" }];
    expect(resolveActivePageOcrCorrections([{ id: "a", ocrCorrections: [] }], "a", draftFallback)).toEqual(
      [],
    );
  });
});
