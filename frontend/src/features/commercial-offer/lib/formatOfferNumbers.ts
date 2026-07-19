export const toNumber = (value: unknown): number | null => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string" && value.trim()) {
    const normalized = value.replace(/\s+/g, "").replace(",", ".");
    const parsed = Number(normalized);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }
  return null;
};

export const formatOfferNumber = (value: unknown): string => {
  const parsed = toNumber(value);
  if (parsed === null) {
    return "0";
  }
  return parsed.toLocaleString("ru-RU", { maximumFractionDigits: 2 });
};

export const formatOfferSum = (qtyValue: unknown, unitPriceValue: unknown): string => {
  const qty = toNumber(qtyValue);
  const unitPrice = toNumber(unitPriceValue);
  if (qty === null || unitPrice === null) {
    return "0";
  }
  return (qty * unitPrice).toLocaleString("ru-RU", { maximumFractionDigits: 2 });
};

/** Серверные итоги (`draft.totals`): не пересчитываем НДС на клиенте. */
export const formatTotalsMoney = (value: number | undefined): string => {
  if (value === undefined || Number.isNaN(value)) {
    return "—";
  }
  return value.toLocaleString("ru-RU", { maximumFractionDigits: 2 });
};
