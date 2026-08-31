export const APPROVAL_THRESHOLD_PERCENT = 16;
export const HIGH_DISCOUNT_CONFIRMATION_KEYWORD = "ПОДТВЕРЖДАЮ";
export const HIGH_DISCOUNT_WARNING =
  "ВНИМАНИЕ: ДЛЯ СКИДКИ ВЫШЕ 16% НЕОБХОДИМО ОДОБРЕНИЕ РУКОВОДСТВА.";

/** Max decimals stored/sent for target-sum-derived discounts (2 is too coarse for exact ₽). */
export const DISCOUNT_PERCENT_PRECISION = 8;

export type DiscountFromTargetSumResult =
  | { ok: true; discountPercent: number }
  | { ok: false; error: string };

export const roundMoney = (value: number): number => Math.round((value + Number.EPSILON) * 100) / 100;

/** Round % for manual entry display (2 dp). Not used when deriving % from a target sum. */
export const roundPercent = (value: number): number => Math.round((value + Number.EPSILON) * 100) / 100;

export const roundPercentPrecise = (value: number): number => {
  const factor = 10 ** DISCOUNT_PERCENT_PRECISION;
  return Math.round((value + Number.EPSILON) * factor) / factor;
};

/** Compact display: up to 8 decimals, trim trailing zeros. */
export const formatDiscountPercentInput = (value: number): string => {
  if (!Number.isFinite(value)) {
    return "";
  }
  const trimmed = String(roundPercentPrecise(value));
  return trimmed.replace(".", ",");
};

export const baseProductsTotal = (items: Array<Record<string, unknown>>): number =>
  roundMoney(
    items.reduce((total, item) => {
      const unitPrice = asFiniteNumber(item.unit_price);
      const quantity = asFiniteNumber(item.qty);
      return unitPrice === null || quantity === null ? total : total + unitPrice * quantity;
    }, 0),
  );

/** Same aggregation as server: Σ(unit_price × qty × factor), then + delivery, money-round. */
export const totalWithDiscountPercent = ({
  baseProductsTotalWithVat,
  deliveryTotal,
  discountPercent,
}: {
  baseProductsTotalWithVat: number;
  deliveryTotal: number;
  discountPercent: number;
}): number =>
  roundMoney(baseProductsTotalWithVat * (1 - discountPercent / 100) + deliveryTotal);

/**
 * Find discount % so that round(base×(1−%/100)+delivery) === target (exact ₽).
 * 2-decimal % is too coarse (~tens–hundreds of ₽ drift); we keep high precision and refine.
 */
export const discountPercentFromTargetSum = ({
  targetTotalWithVat,
  baseProductsTotalWithVat,
  deliveryTotal,
}: {
  targetTotalWithVat: number;
  baseProductsTotalWithVat: number;
  deliveryTotal: number;
}): DiscountFromTargetSumResult => {
  if (![targetTotalWithVat, baseProductsTotalWithVat, deliveryTotal].every(Number.isFinite)) {
    return { ok: false, error: "Введите корректную целевую сумму." };
  }
  if (baseProductsTotalWithVat <= 0) {
    return { ok: false, error: "Нет позиций для расчёта скидки." };
  }
  if (deliveryTotal < 0) {
    return { ok: false, error: "Не удалось корректно определить стоимость доставки." };
  }
  const target = roundMoney(targetTotalWithVat);
  if (target < deliveryTotal) {
    return {
      ok: false,
      error: `Целевая сумма не может быть меньше доставки (${roundMoney(deliveryTotal)} ₽).`,
    };
  }

  const maximumTotal = roundMoney(baseProductsTotalWithVat + deliveryTotal);
  if (target > maximumTotal) {
    return {
      ok: false,
      error: `Целевая сумма не может превышать ${maximumTotal} ₽ без наценки.`,
    };
  }

  if (target === maximumTotal) {
    return { ok: true, discountPercent: 0 };
  }
  if (target === roundMoney(deliveryTotal)) {
    return { ok: true, discountPercent: 100 };
  }

  // Exact algebraic % for the aggregated base (matches server when no per-line rounding).
  let discountPercent = 100 * (1 - (target - deliveryTotal) / baseProductsTotalWithVat);

  // Binary search refine: guarantee roundMoney(projected) === target despite float noise.
  let lo = 0;
  let hi = 100;
  for (let i = 0; i < 80; i += 1) {
    const mid = (lo + hi) / 2;
    const projected = totalWithDiscountPercent({
      baseProductsTotalWithVat,
      deliveryTotal,
      discountPercent: mid,
    });
    if (projected > target) {
      lo = mid; // total too high → need larger discount
    } else {
      hi = mid;
    }
  }
  discountPercent = hi;

  // Prefer the smallest % that still hits target (if a plateau exists).
  const hits = (percent: number) =>
    totalWithDiscountPercent({
      baseProductsTotalWithVat,
      deliveryTotal,
      discountPercent: percent,
    }) === target;

  if (!hits(discountPercent)) {
    // Fallback: algebraic + precise round
    discountPercent = roundPercentPrecise(100 * (1 - (target - deliveryTotal) / baseProductsTotalWithVat));
  } else {
    discountPercent = roundPercentPrecise(discountPercent);
    // If rounding the % breaks the hit, keep unrounded hi from search
    if (!hits(discountPercent)) {
      discountPercent = hi;
    }
  }

  if (!hits(discountPercent)) {
    return {
      ok: false,
      error: "Не удалось подобрать скидку ровно под целевую сумму. Попробуйте другую сумму.",
    };
  }

  return { ok: true, discountPercent };
};

export const targetSumFromDiscountPercent = ({
  discountPercent,
  baseProductsTotalWithVat,
  deliveryTotal,
}: {
  discountPercent: number;
  baseProductsTotalWithVat: number;
  deliveryTotal: number;
}): number | null => {
  if (
    ![discountPercent, baseProductsTotalWithVat, deliveryTotal].every(Number.isFinite) ||
    discountPercent < 0 ||
    discountPercent > 100 ||
    baseProductsTotalWithVat < 0 ||
    deliveryTotal < 0
  ) {
    return null;
  }
  return totalWithDiscountPercent({
    baseProductsTotalWithVat,
    deliveryTotal,
    discountPercent,
  });
};

export const requiresHighDiscountConfirmation = (discountPercent: number): boolean =>
  Number.isFinite(discountPercent) && discountPercent > APPROVAL_THRESHOLD_PERCENT;

const asFiniteNumber = (value: unknown): number | null => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string" && value.trim()) {
    const parsed = Number(value.replace(",", "."));
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
};
