export const APPROVAL_THRESHOLD_PERCENT = 16;
export const HIGH_DISCOUNT_CONFIRMATION_KEYWORD = "ПОДТВЕРЖДАЮ";
export const HIGH_DISCOUNT_WARNING =
  "ВНИМАНИЕ: ДЛЯ СКИДКИ ВЫШЕ 16% НЕОБХОДИМО ОДОБРЕНИЕ РУКОВОДСТВА.";

export type DiscountFromTargetSumResult =
  | { ok: true; discountPercent: number }
  | { ok: false; error: string };

export const roundMoney = (value: number): number => Math.round((value + Number.EPSILON) * 100) / 100;
export const roundPercent = roundMoney;

export const baseProductsTotal = (items: Array<Record<string, unknown>>): number =>
  roundMoney(
    items.reduce((total, item) => {
      const unitPrice = asFiniteNumber(item.unit_price);
      const quantity = asFiniteNumber(item.qty);
      return unitPrice === null || quantity === null ? total : total + unitPrice * quantity;
    }, 0),
  );

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
  if (targetTotalWithVat < deliveryTotal) {
    return {
      ok: false,
      error: `Целевая сумма не может быть меньше доставки (${roundMoney(deliveryTotal)} ₽).`,
    };
  }

  const maximumTotal = roundMoney(baseProductsTotalWithVat + deliveryTotal);
  if (targetTotalWithVat > maximumTotal) {
    return {
      ok: false,
      error: `Целевая сумма не может превышать ${maximumTotal} ₽ без наценки.`,
    };
  }

  return {
    ok: true,
    discountPercent: roundPercent(100 * (1 - (targetTotalWithVat - deliveryTotal) / baseProductsTotalWithVat)),
  };
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
  return roundMoney(baseProductsTotalWithVat * (1 - discountPercent / 100) + deliveryTotal);
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
