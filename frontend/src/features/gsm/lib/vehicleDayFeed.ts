import type { GsmTransaction, GsmWaybill } from "@/features/gsm/types/gsm";

const ANCHOR_SERVICES = new Set(["fuel", "wash"]);

export type VehicleDayFeed = {
  date: string; // YYYY-MM-DD
  txs: GsmTransaction[];
  waybills: GsmWaybill[];
  isGap: boolean;
};

const parseMonth = (month: string): { year: number; monthIndex: number } => {
  const year = Number(month.slice(0, 4));
  const monthIndex = Number(month.slice(5, 7)) - 1;
  return { year, monthIndex };
};

const pad2 = (value: number): string => String(value).padStart(2, "0");

export const monthBounds = (month: string): { from: string; to: string } => {
  const { year, monthIndex } = parseMonth(month);
  const lastDay = new Date(year, monthIndex + 1, 0).getDate();
  return {
    from: `${month}-01`,
    to: `${month}-${pad2(lastDay)}`,
  };
};

export const shiftMonth = (month: string, delta: number): string => {
  const { year, monthIndex } = parseMonth(month);
  const d = new Date(year, monthIndex + delta, 1);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
};

export const buildVehicleDayFeed = (
  periodFrom: string,
  periodTo: string,
  waybills: GsmWaybill[],
  transactions: GsmTransaction[],
): VehicleDayFeed[] => {
  if (periodFrom > periodTo) return [];

  const txsByDate = new Map<string, GsmTransaction[]>();
  for (const tx of transactions) {
    const date = tx.ts.slice(0, 10);
    if (date < periodFrom || date > periodTo) continue;
    const list = txsByDate.get(date) ?? [];
    list.push(tx);
    txsByDate.set(date, list);
  }

  const waybillsByDate = new Map<string, GsmWaybill[]>();
  for (const wb of waybills) {
    if (wb.date < periodFrom || wb.date > periodTo) continue;
    const list = waybillsByDate.get(wb.date) ?? [];
    list.push(wb);
    waybillsByDate.set(wb.date, list);
  }

  const dates = [...new Set([...txsByDate.keys(), ...waybillsByDate.keys()])].sort();

  return dates.map((date) => {
    const txs = [...(txsByDate.get(date) ?? [])].sort((a, b) =>
      a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : 0,
    );
    const dayWaybills = [...(waybillsByDate.get(date) ?? [])].sort((a, b) => a.id - b.id);
    const isGap =
      dayWaybills.length === 0 && txs.some((t) => ANCHOR_SERVICES.has(t.service_type));
    return { date, txs, waybills: dayWaybills, isGap };
  });
};
