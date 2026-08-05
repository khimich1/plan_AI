import { describe, expect, it } from "vitest";
import {
  discountPercentFromTargetSum,
  requiresHighDiscountConfirmation,
  targetSumFromDiscountPercent,
} from "@/features/commercial-offer/lib/discountFromTargetSum";

describe("target-sum discount math", () => {
  it("calculates discounts while preserving delivery", () => {
    expect(
      discountPercentFromTargetSum({
        targetTotalWithVat: 2_000_000,
        baseProductsTotalWithVat: 2_400_000,
        deliveryTotal: 100_000,
      }),
    ).toEqual({ ok: true, discountPercent: 20.83 });
    expect(
      targetSumFromDiscountPercent({
        discountPercent: 20.83,
        baseProductsTotalWithVat: 2_400_000,
        deliveryTotal: 100_000,
      }),
    ).toBe(2_000_080);
  });

  it("handles minimum and maximum targets", () => {
    expect(
      discountPercentFromTargetSum({
        targetTotalWithVat: 100,
        baseProductsTotalWithVat: 900,
        deliveryTotal: 100,
      }),
    ).toEqual({ ok: true, discountPercent: 100 });
    expect(
      discountPercentFromTargetSum({
        targetTotalWithVat: 1000,
        baseProductsTotalWithVat: 900,
        deliveryTotal: 100,
      }),
    ).toEqual({ ok: true, discountPercent: 0 });
  });

  it("rejects zero base and targets beyond supported bounds", () => {
    expect(
      discountPercentFromTargetSum({ targetTotalWithVat: 100, baseProductsTotalWithVat: 0, deliveryTotal: 0 }),
    ).toMatchObject({ ok: false });
    expect(
      discountPercentFromTargetSum({ targetTotalWithVat: 99, baseProductsTotalWithVat: 900, deliveryTotal: 100 }),
    ).toMatchObject({ ok: false });
    expect(
      discountPercentFromTargetSum({ targetTotalWithVat: 1001, baseProductsTotalWithVat: 900, deliveryTotal: 100 }),
    ).toMatchObject({ ok: false });
  });

  it("requires approval strictly above 16 percent", () => {
    expect(requiresHighDiscountConfirmation(16)).toBe(false);
    expect(requiresHighDiscountConfirmation(16.01)).toBe(true);
  });
});
