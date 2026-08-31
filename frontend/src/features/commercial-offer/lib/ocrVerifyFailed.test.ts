import { describe, expect, it } from "vitest";

import { resolveActivePageOcrVerifyFailed } from "@/features/commercial-offer/lib/ocrVerifyFailed";

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
