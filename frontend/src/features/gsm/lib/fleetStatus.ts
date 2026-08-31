import type { VehiclePeriodStatus } from "@/features/gsm/types/gsm";

export type FleetStatusTone = "muted" | "warning" | "danger" | "info" | "success";

export type FleetStatusMeta = {
  label: string;
  tone: FleetStatusTone;
};

const META: Record<VehiclePeriodStatus, FleetStatusMeta> = {
  no_data: { label: "Нет данных", tone: "muted" },
  needs_generation: { label: "Требуется генерация", tone: "warning" },
  has_red_days: { label: "Есть красные дни", tone: "danger" },
  drafts_pending: { label: "Черновики", tone: "warning" },
  pending_export: { label: "Готово к экспорту", tone: "info" },
  ready: { label: "Выгружено", tone: "success" },
};

export const fleetStatusMeta = (status: VehiclePeriodStatus): FleetStatusMeta => META[status];

export const TONE_COLORS: Record<FleetStatusTone, { bg: string; fg: string }> = {
  muted: { bg: "#f2f4f7", fg: "#475467" },
  warning: { bg: "#fef0c7", fg: "#93370d" },
  danger: { bg: "#fee4e2", fg: "#b42318" },
  info: { bg: "#d1e9ff", fg: "#175cd3" },
  success: { bg: "#d1fadf", fg: "#067647" },
};

export const currentMonthBounds = (now = new Date()): { from: string; to: string } => {
  const y = now.getFullYear();
  const m = now.getMonth();
  const pad = (n: number) => String(n).padStart(2, "0");
  const iso = (year: number, monthIndex: number, day: number) =>
    `${year}-${pad(monthIndex + 1)}-${pad(day)}`;
  const lastDay = new Date(y, m + 1, 0).getDate();
  return { from: iso(y, m, 1), to: iso(y, m, lastDay) };
};

export const previousMonthBounds = (fromIso: string): { from: string; to: string } => {
  const [year, month] = fromIso.split("-").map(Number);
  const prev = new Date(year, month - 2, 1);
  return currentMonthBounds(prev);
};

export const litersDiffHidden = (wbCount: number): boolean => wbCount === 0;

export const litersDiffOk = (diff: number): boolean => Math.abs(diff) <= 0.01;

export const openBeforeSummary = (rows: { open_before: number }[]): { pl: number; vehicles: number } => {
  const pl = rows.reduce((sum, row) => sum + row.open_before, 0);
  const vehicles = rows.filter((row) => row.open_before > 0).length;
  return { pl, vehicles };
};

/** Max `open_before_month` among rows with a tail (nearest unclosed month). */
export const fleetOpenBeforeMonth = (
  rows: { open_before: number; open_before_month: string | null }[],
): string | null => {
  let max: string | null = null;
  for (const row of rows) {
    if (row.open_before > 0 && row.open_before_month) {
      if (max == null || row.open_before_month > max) {
        max = row.open_before_month;
      }
    }
  }
  return max;
};

/** Capitalized ru-RU month name from YYYY-MM, e.g. «Июль». */
export const openBeforeMonthLabel = (yyyyMm: string): string => {
  const [year, month] = yyyyMm.split("-").map(Number);
  const raw = new Date(year, month - 1, 1).toLocaleDateString("ru-RU", { month: "long" });
  if (!raw) {
    return yyyyMm;
  }
  return raw.charAt(0).toUpperCase() + raw.slice(1);
};
