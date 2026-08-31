import { openBeforeMonthLabel } from "@/features/gsm/lib/fleetStatus";
import type { FleetOverviewRow, GsmWaybill, WaybillWarningCode } from "@/features/gsm/types/gsm";

const HARD_WARNING_CODES = new Set<string>(["manual_intervention", "unsolvable"]);
const SOFT_WARNING_CODES = new Set<string>([
  "balance_route",
  "borrowed_route",
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

export type BulkExportExclusion = {
  vehicleId: number;
  label: string;
  reason: string;
};

export type BulkExportPlan = {
  cleanIds: number[];
  excluded: BulkExportExclusion[];
  alreadyExportedCount: number;
  cleanCount: number;
};

const vehicleLabel = (row: FleetOverviewRow): string =>
  `${row.vehicle.name} (${row.vehicle.plate_number})`;

/** Split selected overview rows: red_days block zip; order is vehicle id. */
export const planBulkExport = (
  rows: FleetOverviewRow[],
  selectedIds: number[],
): BulkExportPlan => {
  const selected = new Set(selectedIds);
  const ordered = rows
    .filter((row) => selected.has(row.vehicle.id))
    .sort((a, b) => a.vehicle.id - b.vehicle.id);
  const excluded: BulkExportExclusion[] = [];
  const clean: FleetOverviewRow[] = [];
  for (const row of ordered) {
    if (row.red_days > 0) {
      excluded.push({
        vehicleId: row.vehicle.id,
        label: vehicleLabel(row),
        reason: `Исправьте дни ручной доработки (${row.red_days}).`,
      });
    } else {
      clean.push(row);
    }
  }
  return {
    cleanIds: clean.map((row) => row.vehicle.id),
    excluded,
    alreadyExportedCount: clean.filter((row) => row.exported_count > 0).length,
    cleanCount: clean.length,
  };
};

/** Confirm copy by counters; no per-day details. */
export const bulkExportConfirmMessages = (plan: BulkExportPlan): string[] => {
  if (plan.cleanCount === 0 || plan.alreadyExportedCount === 0) {
    return [];
  }
  return [`${plan.alreadyExportedCount} из ${plan.cleanCount} уже экспортировались. Скачать снова?`];
};

const periodMonth = (periodFrom: string): string => periodFrom.slice(0, 7);

const alreadyExportedForKit = (row: FleetOverviewRow): boolean =>
  row.wb_count > 0 && row.exported_count === row.wb_count;

const scopeKitRows = (
  rows: FleetOverviewRow[],
  selectedIds: number[] | null,
): FleetOverviewRow[] => {
  if (selectedIds === null) {
    return rows;
  }
  const selected = new Set(selectedIds);
  return rows.filter((row) => selected.has(row.vehicle.id));
};

/**
 * Split overview rows for the month-close kit (usage report + waybills).
 * `selectedIds === null` means all rows; `[]` means none.
 */
export const planKit = (
  rows: FleetOverviewRow[],
  selectedIds: number[] | null,
  periodFrom: string,
): BulkExportPlan => {
  const ordered = [...scopeKitRows(rows, selectedIds)].sort((a, b) => a.vehicle.id - b.vehicle.id);
  const currentYm = periodMonth(periodFrom);
  const excluded: BulkExportExclusion[] = [];
  const clean: FleetOverviewRow[] = [];
  for (const row of ordered) {
    if (row.red_days > 0) {
      excluded.push({
        vehicleId: row.vehicle.id,
        label: vehicleLabel(row),
        reason: `Исправьте дни ручной доработки (${row.red_days}).`,
      });
    } else if (row.open_before > 0 && currentYm !== row.open_before_month) {
      const tailYm = row.open_before_month ?? currentYm;
      excluded.push({
        vehicleId: row.vehicle.id,
        label: vehicleLabel(row),
        reason: `Сначала выгрузите ${openBeforeMonthLabel(tailYm)}`,
      });
    } else if (row.chain_broken) {
      excluded.push({
        vehicleId: row.vehicle.id,
        label: vehicleLabel(row),
        reason: `Пересчитайте ${openBeforeMonthLabel(currentYm)}: бак не сходится с предыдущим`,
      });
    } else {
      clean.push(row);
    }
  }
  return {
    cleanIds: clean.map((row) => row.vehicle.id),
    excluded,
    alreadyExportedCount: clean.filter(alreadyExportedForKit).length,
    cleanCount: clean.length,
  };
};

/** Confirm copy for kit download; mirrors bulk zip confirms. */
export const bulkKitConfirmMessages = bulkExportConfirmMessages;

export type BulkGenerateSkip = {
  vehicleId: number;
  label: string;
  reason: string;
};

export type BulkGeneratePlan = {
  eligibleIds: number[];
  skipped: BulkGenerateSkip[];
};

const isTailMonth = (row: FleetOverviewRow, periodFrom: string): boolean =>
  row.open_before_month != null && periodMonth(periodFrom) === row.open_before_month;

/**
 * Split selected overview rows for bulk generate.
 * Skip open_before / chain_broken unless the selected period is that vehicle's tail month.
 */
export const planBulkGenerate = (
  rows: FleetOverviewRow[],
  selectedIds: number[],
  periodFrom: string,
): BulkGeneratePlan => {
  const selected = new Set(selectedIds);
  const ordered = rows
    .filter((row) => selected.has(row.vehicle.id))
    .sort((a, b) => a.vehicle.id - b.vehicle.id);
  const currentYm = periodMonth(periodFrom);
  const skipped: BulkGenerateSkip[] = [];
  const eligibleIds: number[] = [];
  for (const row of ordered) {
    if (isTailMonth(row, periodFrom)) {
      eligibleIds.push(row.vehicle.id);
      continue;
    }
    if (row.open_before > 0) {
      const tailYm = row.open_before_month ?? currentYm;
      skipped.push({
        vehicleId: row.vehicle.id,
        label: vehicleLabel(row),
        reason: `сначала выгрузите ${openBeforeMonthLabel(tailYm)}`,
      });
    } else if (row.chain_broken) {
      skipped.push({
        vehicleId: row.vehicle.id,
        label: vehicleLabel(row),
        reason: `пересчитайте ${openBeforeMonthLabel(currentYm)}`,
      });
    } else {
      eligibleIds.push(row.vehicle.id);
    }
  }
  return { eligibleIds, skipped };
};
