import { toNumber } from "@/features/commercial-offer/lib/formatOfferNumbers";

export const discountedUnitPrice = (
  unitPrice: unknown,
  discountPercent: unknown,
): number | null => {
  const price = toNumber(unitPrice);
  const percent = toNumber(discountPercent) ?? 0;
  if (price === null) {
    return null;
  }
  const clamped = Math.min(Math.max(percent, 0), 100);
  return price * (1 - clamped / 100);
};
