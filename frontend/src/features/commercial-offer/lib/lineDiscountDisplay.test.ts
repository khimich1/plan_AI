import { describe, expect, it } from "vitest";
import { discountedUnitPrice } from "@/features/commercial-offer/lib/lineDiscountDisplay";

describe("discountedUnitPrice", () => {
  it("applies 10% discount", () => {
    expect(discountedUnitPrice(42508, 10)).toBeCloseTo(38257.2, 5);
  });

  it("returns list price at 0%", () => {
    expect(discountedUnitPrice(42508, 0)).toBe(42508);
  });

  it("returns zero at 100%", () => {
    expect(discountedUnitPrice(42508, 100)).toBe(0);
  });

  it("returns null for invalid price", () => {
    expect(discountedUnitPrice("not-a-price", 10)).toBeNull();
  });
});
