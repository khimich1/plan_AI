import type { GsmWaybill, WaybillWarningCode } from "@/features/gsm/types/gsm";

const HARD_WARNING_CODES = new Set<string>(["manual_intervention", "unsolvable"]);
const SOFT_WARNING_CODES = new Set<string>([
  "balance_route",
  "hook_above_threshold",
  "weekend_anchor",
]);

export type ExportHardBlock = {
  blocked: boolean;
  dates: string[];
};

const hasHardWarning = (codes: readonly WaybillWarningCode[]): boolean =>
  codes.some((code) => HARD_WARNING_CODES.has(code));

const hasSoftWarning = (codes: readonly WaybillWarningCode[]): boolean =>
  codes.some((code) => SOFT_WARNING_CODES.has(code));

/** Hard stop: manual_intervention / unsolvable on a day or period. */
export const exportHardBlock = (
  waybills: GsmWaybill[],
  periodWarnings: readonly WaybillWarningCode[] = [],
): ExportHardBlock => {
  const dates = waybills.filter((day) => hasHardWarning(day.warnings)).map((day) => day.date);
  return {
    blocked: dates.length > 0 || hasHardWarning(periodWarnings),
    dates,
  };
};

/** Yellow warnings that need a confirm, but do not disable export. */
export const hasSoftExportWarnings = (
  waybills: GsmWaybill[],
  periodWarnings: readonly WaybillWarningCode[] = [],
): boolean => {
  if (exportHardBlock(waybills, periodWarnings).blocked) {
    return false;
  }
  return waybills.some((day) => hasSoftWarning(day.warnings)) || hasSoftWarning(periodWarnings);
};

export const hasExportedDays = (waybills: GsmWaybill[]): boolean =>
  waybills.some((day) => day.status === "exported");

/** Confirm copy before export; empty means export immediately. */
export const exportConfirmMessages = (
  waybills: GsmWaybill[],
  periodWarnings: readonly WaybillWarningCode[] = [],
): string[] => {
  const messages: string[] = [];
  if (hasSoftExportWarnings(waybills, periodWarnings)) {
    messages.push("В периоде есть предупреждения. Экспортировать всё равно?");
  }
  if (hasExportedDays(waybills)) {
    messages.push("Период уже экспортировался. Скачать снова?");
  }
  return messages;
};

/** Human-readable reason why «Экспорт zip» is disabled, or null if allowed. */
export const exportDisabledReason = (
  waybills: GsmWaybill[],
  periodWarnings: readonly WaybillWarningCode[] = [],
): string | null => {
  if (waybills.length === 0) {
    return "Нет путевых листов за период.";
  }
  const hard = exportHardBlock(waybills, periodWarnings);
  if (!hard.blocked) {
    return null;
  }
  if (hard.dates.length > 0) {
    return `Исправьте дни ручной доработки: ${hard.dates.join(", ")}.`;
  }
  return "Период нерешаем — исправьте дни вручную.";
};
