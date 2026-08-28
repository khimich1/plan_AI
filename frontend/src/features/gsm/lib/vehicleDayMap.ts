import type { GsmWaybill } from "@/features/gsm/types/gsm";

const ANCHOR_SERVICES = new Set(["fuel", "wash"]);

export type VehicleDayCell = {
  date: string;
  hasTx: boolean;
  hasPl: boolean;
  isGap: boolean;
  isRed: boolean;
  waybill: GsmWaybill | null;
};

const addDays = (iso: string, days: number): string => {
  const d = new Date(`${iso}T12:00:00`);
  d.setDate(d.getDate() + days);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
};

/** Monday=0 … Sunday=6 for YYYY-MM-DD (local noon). */
const mondayBasedDow = (iso: string): number => {
  const d = new Date(`${iso}T12:00:00`);
  return (d.getDay() + 6) % 7;
};

export const buildVehicleDayCells = (
  periodFrom: string,
  periodTo: string,
  waybills: GsmWaybill[],
  transactions: { ts: string; service_type: string }[],
): VehicleDayCell[] => {
  if (periodFrom > periodTo) return [];

  const txDates = new Set<string>();
  for (const tx of transactions) {
    if (!ANCHOR_SERVICES.has(tx.service_type)) continue;
    txDates.add(tx.ts.slice(0, 10));
  }

  const waybillsByDate = new Map<string, GsmWaybill[]>();
  for (const wb of waybills) {
    const list = waybillsByDate.get(wb.date) ?? [];
    list.push(wb);
    waybillsByDate.set(wb.date, list);
  }

  const cells: VehicleDayCell[] = [];
  for (let date = periodFrom; date <= periodTo; date = addDays(date, 1)) {
    const dayWaybills = waybillsByDate.get(date) ?? [];
    const waybill =
      dayWaybills.length === 0
        ? null
        : [...dayWaybills].sort((a, b) => a.id - b.id)[0];
    const hasTx = txDates.has(date);
    const hasPl = waybill != null;
    const isGap = hasTx && !hasPl;
    const isRed =
      hasPl && (waybill?.warnings.includes("manual_intervention") ?? false);
    cells.push({ date, hasTx, hasPl, isGap, isRed, waybill });
  }
  return cells;
};

/** Week slots Mon–Sun; null = padding outside [from, to]. */
export const layoutWeekSlots = (
  cells: VehicleDayCell[],
): Array<VehicleDayCell | null> => {
  if (cells.length === 0) return [];

  const first = cells[0].date;
  const last = cells[cells.length - 1].date;
  const leading = mondayBasedDow(first);
  const trailing = 6 - mondayBasedDow(last);

  return [
    ...Array.from({ length: leading }, () => null),
    ...cells,
    ...Array.from({ length: trailing }, () => null),
  ];
};
