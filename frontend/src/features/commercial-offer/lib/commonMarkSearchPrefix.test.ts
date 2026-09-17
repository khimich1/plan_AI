import { describe, expect, it } from "vitest";
import {
  catalogRowSharesUnpricedStem,
  commonMarkSearchPrefix,
  normalizeMarkForSearch,
} from "@/features/commercial-offer/lib/commonMarkSearchPrefix";

describe("commonMarkSearchPrefix", () => {
  it("returns C110 for C110.30-6 and C110.40-8.1", () => {
    expect(commonMarkSearchPrefix(["C110.30-6", "C110.40-8.1"])).toBe("C110");
  });

  it("stems a single missing mark so neighbors stay searchable", () => {
    expect(commonMarkSearchPrefix(["C110.30-6"])).toBe("C110.30");
  });

  it("treats Cyrillic С as C", () => {
    expect(commonMarkSearchPrefix(["С110.30-6", "C110.40-8.1"])).toBe("C110");
  });

  it("returns empty when the shared prefix is shorter than 4", () => {
    expect(commonMarkSearchPrefix(["C110.30-6", "ЛС11"])).toBe("");
  });

  it("returns empty for an empty list", () => {
    expect(commonMarkSearchPrefix([])).toBe("");
  });
});

describe("normalizeMarkForSearch", () => {
  it("collapses spaces and folds C↔С", () => {
    expect(normalizeMarkForSearch(" с 110.30-6 ")).toBe("C110.30-6");
  });
});

describe("catalogRowSharesUnpricedStem", () => {
  it("highlights C110.30-9 against missing C110.30-6", () => {
    expect(catalogRowSharesUnpricedStem("С110.30-9", ["C110.30-6"])).toBe(true);
  });

  it("does not highlight a different family", () => {
    expect(catalogRowSharesUnpricedStem("С120.35-12", ["C110.30-6"])).toBe(false);
  });
});
