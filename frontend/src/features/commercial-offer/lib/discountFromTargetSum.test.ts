import { describe, expect, it } from "vitest";
import {
  discountPercentFromTargetSum,
  requiresHighDiscountConfirmation,
  targetSumFromDiscountPercent,
  totalWithDiscountPercent,
} from "@/features/commercial-offer/lib/discountFromTargetSum";

describe("target-sum discount math", () => {
  it("calculates a discount that reconstructs the target sum exactly", () => {
    const target = 2_000_000;
    const base = 2_400_000;
    const delivery = 100_000;
    const result = discountPercentFromTargetSum({
      targetTotalWithVat: target,
      baseProductsTotalWithVat: base,
      deliveryTotal: delivery,
    });
    expect(result.ok).toBe(true);
    if (!result.ok) {
      return;
    }
    expect(
      totalWithDiscountPercent({
        baseProductsTotalWithVat: base,
        deliveryTotal: delivery,
        discountPercent: result.discountPercent,
      }),
    ).toBe(target);
    expect(
      targetSumFromDiscountPercent({
        discountPercent: result.discountPercent,
        baseProductsTotalWithVat: base,
        deliveryTotal: delivery,
      }),
    ).toBe(target);
  });

  it("hits exact round targets like 2_500_000", () => {
    const target = 2_500_000;
    const base = 2_901_234.56;
    const delivery = 50_000;
    const result = discountPercentFromTargetSum({
      targetTotalWithVat: target,
      baseProductsTotalWithVat: base,
      deliveryTotal: delivery,
    });
    expect(result).toMatchObject({ ok: true });
    if (!result.ok) {
      return;
    }
    expect(
      targetSumFromDiscountPercent({
        discountPercent: result.discountPercent,
        baseProductsTotalWithVat: base,
        deliveryTotal: delivery,
      }),
    ).toBe(target);
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
