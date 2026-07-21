import type { BreakdownTable } from "@/features/commercial-offer/types/commercialOffer";

export const normalizePlateName = (name: string): string =>
  name.trim().replace(/\s*\(нагрузка\?\)\s*$/i, "");

export const findBreakdownTable = (
  tables: BreakdownTable[],
  orderItemName: string,
): BreakdownTable | undefined => {
  const normalized = normalizePlateName(orderItemName);
  if (!normalized) {
    return undefined;
  }

  const exact = tables.find((table) => normalizePlateName(table.name) === normalized);
  if (exact) {
    return exact;
  }

  return tables.find((table) => {
    const tableName = normalizePlateName(table.name);
    return tableName.includes(normalized) || normalized.includes(tableName);
  });
};

export const isBreakdownTotalRow = (component: string): boolean => {
  const label = component.trim().toLowerCase();
  return (
    label.startsWith("итого") ||
    label.startsWith("округлено") ||
    label.startsWith("за ") && label.includes("плит")
  );
};
