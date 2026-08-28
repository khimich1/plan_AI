import { findGsmHomeTwin, isGsmHomeBase } from "@/features/gsm/lib/gsmHome";
import type { GsmRoute, GsmSeasonMode, GsmVehicle, GsmWaybill } from "@/features/gsm/types/gsm";

const PROTECTED = new Set(["confirmed", "exported"]);

export type DownstreamPreviewDay = {
  id: number;
  date: string;
  status: string;
  km: number;
  fuel_start: number | null;
  fuel_end: number | null;
  odometer_start: number | null;
  odometer_end: number | null;
  changed: boolean;
};

export type DownstreamPreviewResult = {
  days: DownstreamPreviewDay[];
  error: string | null;
};

const round2 = (n: number): number => Math.round(n * 100) / 100;

export const burnForKm = (km: number, normPer100km: number): number =>
  round2((km * normPer100km) / 100);

/**
 * Mirror of the backend season rule: settings hold the latest dated switch, seasons
 * alternate starting from summer. Days on/after the switch take `seasonMode`, earlier
 * days take the opposite season; without a recorded switch every day is summer.
 */
export const normForDate = (
  dayIso: string,
  vehicle: Pick<GsmVehicle, "norm_summer" | "norm_winter">,
  seasonMode: GsmSeasonMode,
  seasonSwitchedAt: string | null,
): number => {
  if (!seasonSwitchedAt) {
    return vehicle.norm_summer;
  }
  const isWinter =
    dayIso >= seasonSwitchedAt ? seasonMode === "winter" : seasonMode === "summer";
  return isWinter ? vehicle.norm_winter : vehicle.norm_summer;
};

/**
 * Client-side preview of fuel/odo chain after editing a day's km.
 * Mirrors backend `_rechain_downstream`: drafts recompute; confirmed/exported stay fixed.
 */
export const previewDownstream = (params: {
  edited: GsmWaybill;
  editedKm: number;
  periodWaybills: GsmWaybill[];
  vehicle: Pick<GsmVehicle, "norm_summer" | "norm_winter" | "tank_volume_liters">;
  seasonMode: GsmSeasonMode;
  seasonSwitchedAt: string | null;
}): DownstreamPreviewResult => {
  const { edited, editedKm, periodWaybills, vehicle, seasonMode, seasonSwitchedAt } = params;
  const sorted = [...periodWaybills].sort((a, b) => a.date.localeCompare(b.date));
  const tank = vehicle.tank_volume_liters;

  const applyDay = (
    dayIso: string,
    fuelStart: number,
    fuelIssued: number,
    km: number,
    odoStart: number,
  ): { fuel_end: number; odometer_end: number } | { error: string } => {
    const norm = normForDate(dayIso, vehicle, seasonMode, seasonSwitchedAt);
    const burn = burnForKm(km, norm);
    const fuelEnd = round2(fuelStart + fuelIssued - burn);
    if (fuelEnd < 0 || fuelEnd > tank) {
      return {
        error: `Остаток ${fuelEnd} л вне коридора [0…${tank}] на ${dayIso}`,
      };
    }
    return { fuel_end: fuelEnd, odometer_end: odoStart + km };
  };

  const days: DownstreamPreviewDay[] = [];
  let error: string | null = null;
  let fuelCur: number | null = null;
  let odoCur: number | null = null;
  let pastEdited = false;

  for (const row of sorted) {
    if (row.date < edited.date) {
      days.push({
        id: row.id,
        date: row.date,
        status: row.status,
        km: row.km,
        fuel_start: row.fuel_start,
        fuel_end: row.fuel_end,
        odometer_start: row.odometer_start,
        odometer_end: row.odometer_end,
        changed: false,
      });
      continue;
    }

    if (row.id === edited.id || row.date === edited.date) {
      const fuelStart = edited.fuel_start ?? 0;
      const fuelIssued = edited.fuel_issued ?? 0;
      const odoStart = edited.odometer_start ?? 0;
      const applied = applyDay(edited.date, fuelStart, fuelIssued, editedKm, odoStart);
      if ("error" in applied) {
        error = applied.error;
        days.push({
          id: edited.id,
          date: edited.date,
          status: edited.status,
          km: editedKm,
          fuel_start: fuelStart,
          fuel_end: null,
          odometer_start: odoStart,
          odometer_end: null,
          changed: true,
        });
        pastEdited = true;
        fuelCur = null;
        odoCur = null;
        continue;
      }
      const changed =
        editedKm !== edited.km ||
        applied.fuel_end !== edited.fuel_end ||
        applied.odometer_end !== edited.odometer_end;
      days.push({
        id: edited.id,
        date: edited.date,
        status: edited.status,
        km: editedKm,
        fuel_start: fuelStart,
        fuel_end: applied.fuel_end,
        odometer_start: odoStart,
        odometer_end: applied.odometer_end,
        changed,
      });
      fuelCur = applied.fuel_end;
      odoCur = applied.odometer_end;
      pastEdited = true;
      continue;
    }

    if (!pastEdited) {
      days.push({
        id: row.id,
        date: row.date,
        status: row.status,
        km: row.km,
        fuel_start: row.fuel_start,
        fuel_end: row.fuel_end,
        odometer_start: row.odometer_start,
        odometer_end: row.odometer_end,
        changed: false,
      });
      continue;
    }

    if (PROTECTED.has(row.status)) {
      days.push({
        id: row.id,
        date: row.date,
        status: row.status,
        km: row.km,
        fuel_start: row.fuel_start,
        fuel_end: row.fuel_end,
        odometer_start: row.odometer_start,
        odometer_end: row.odometer_end,
        changed: false,
      });
      if (row.fuel_end != null && row.odometer_end != null) {
        fuelCur = row.fuel_end;
        odoCur = row.odometer_end;
      }
      continue;
    }

    if (fuelCur == null || odoCur == null || error) {
      days.push({
        id: row.id,
        date: row.date,
        status: row.status,
        km: row.km,
        fuel_start: row.fuel_start,
        fuel_end: row.fuel_end,
        odometer_start: row.odometer_start,
        odometer_end: row.odometer_end,
        changed: false,
      });
      continue;
    }

    const fuelIssued = row.fuel_issued ?? 0;
    const applied = applyDay(row.date, fuelCur, fuelIssued, row.km, odoCur);
    if ("error" in applied) {
      error = applied.error;
      days.push({
        id: row.id,
        date: row.date,
        status: row.status,
        km: row.km,
        fuel_start: fuelCur,
        fuel_end: null,
        odometer_start: odoCur,
        odometer_end: null,
        changed: true,
      });
      fuelCur = null;
      odoCur = null;
      continue;
    }

    const changed =
      applied.fuel_end !== row.fuel_end ||
      applied.odometer_end !== row.odometer_end ||
      fuelCur !== row.fuel_start ||
      odoCur !== row.odometer_start;

    days.push({
      id: row.id,
      date: row.date,
      status: row.status,
      km: row.km,
      fuel_start: fuelCur,
      fuel_end: applied.fuel_end,
      odometer_start: odoCur,
      odometer_end: applied.odometer_end,
      changed,
    });
    fuelCur = applied.fuel_end;
    odoCur = applied.odometer_end;
  }

  return { days, error };
};

/** Previous day in the period list used to autofill manual fuel/odo start. */
export const previousWaybillChainStart = (
  waybills: GsmWaybill[],
  dateIso: string,
): { fuel_start: number | null; odometer_start: number | null } => {
  const prev = [...waybills]
    .filter((w) => w.date < dateIso)
    .sort((a, b) => b.date.localeCompare(a.date))[0];
  if (!prev) {
    return { fuel_start: null, odometer_start: null };
  }
  return {
    fuel_start: prev.fuel_end,
    odometer_start: prev.odometer_end,
  };
};

export const libraryRouteToLegs = (
  route: {
    id: number;
    addr_a: string;
    addr_b: string;
    km: number;
    typical_station_ids?: number[];
  },
  catalog?: GsmRoute[],
) => {
  const aHome = isGsmHomeBase(route.addr_a);
  const bHome = isGsmHomeBase(route.addr_b);
  const flip = !aHome && bHome;
  const from = flip ? route.addr_b : route.addr_a;
  const to = flip ? route.addr_a : route.addr_b;
  const twin = flip && catalog ? findGsmHomeTwin(route, catalog) : null;
  return [
    {
      from,
      to,
      km: route.km,
      route_id: twin?.id ?? route.id,
      station_id: route.typical_station_ids?.[0] ?? null,
    },
  ];
};
