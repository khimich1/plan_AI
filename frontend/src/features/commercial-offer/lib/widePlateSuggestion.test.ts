import { describe, expect, it } from "vitest";

import { buildAutoSplitSuggestion } from "@/features/commercial-offer/lib/widePlateSuggestion";

describe("buildAutoSplitSuggestion", () => {
  it("splits a bare mark 34-15-10п 15 into 12 + 3 with qty 15", () => {
    const suggestion = buildAutoSplitSuggestion("34-15-10п 15", 5);
    expect(suggestion).toMatch(/34/);
    expect(suggestion).toMatch(/12/);
    expect(suggestion).toMatch(/3(?:\.0)?/);
    expect(suggestion).toMatch(/15/);
    expect(suggestion).not.toMatch(/ПБ 60-12/);
  });
});
