import type { GsmWaybill, WaybillRouteLeg, WaybillWarningCode } from "@/features/gsm/types/gsm";

export type WarningMeta = {
  short: string;
  reason: string;
};

const KNOWN: Record<string, WarningMeta> = {
  weekend_anchor: {
    short: "Выходной",
    reason: "Якорный день (заправка или мойка) выпал на выходной или праздник РФ.",
  },
  hook_above_threshold: {
    short: "Крюк",
    reason: "Крюк маршрута до АЗС превышает порог — день лучше проверить вручную.",
  },
  unsolvable: {
    short: "Нерешаемо",
    reason:
      "Недостаточно будних дней между якорями, чтобы удержать остаток бака в коридоре [0…бак].",
  },
  manual_intervention: {
    short: "Ручная доработка",
    reason: "Баланс не сошёлся автоматически; бухгалтер правит день вручную.",
  },
  balance_route: {
    short: "Маршрут для баланса",
    reason: "Выбран удлинённый маршрут для освобождения места в баке.",
  },
};

export const warningMeta = (code: WaybillWarningCode): WarningMeta =>
  KNOWN[code] ?? {
    short: String(code),
    reason: `Предупреждение генератора: ${code}.`,
  };

export const routeFrom = (leg: WaybillRouteLeg): string => leg.from ?? leg.from_addr ?? "—";

export const routeTo = (leg: WaybillRouteLeg): string => leg.to ?? leg.to_addr ?? "—";

export const formatRouteSummary = (route: WaybillRouteLeg[]): string => {
  if (!route.length) {
    return "—";
  }
  const first = route[0];
  if (route.length === 1) {
    return `${routeFrom(first)} → ${routeTo(first)}`;
  }
  return `${routeFrom(first)} → … → ${routeTo(route[route.length - 1])} (${route.length} плеч)`;
};

/** Heuristic: fuel day or day with anchor-specific warnings / station on route. */
export const isAnchorDay = (waybill: GsmWaybill): boolean => {
  if ((waybill.fuel_issued ?? 0) > 0) {
    return true;
  }
  if (waybill.warnings.some((w) => w === "weekend_anchor" || w === "hook_above_threshold")) {
    return true;
  }
  return waybill.route.some((leg) => leg.station_id != null);
};

/** Day saved as draft because the generator could not keep the tank in corridor. */
export const isProblematicDay = (waybill: GsmWaybill): boolean =>
  waybill.warnings.includes("manual_intervention");

export const formatLiters = (value: number | null | undefined): string => {
  if (value == null || !Number.isFinite(value)) {
    return "—";
  }
  return `${value.toLocaleString("ru-RU", { maximumFractionDigits: 2 })} л`;
};

export const formatKm = (value: number | null | undefined): string => {
  if (value == null || !Number.isFinite(value)) {
    return "—";
  }
  return `${value.toLocaleString("ru-RU")} км`;
};

export const formatOdometer = (value: number | null | undefined): string => {
  if (value == null || !Number.isFinite(value)) {
    return "—";
  }
  return value.toLocaleString("ru-RU");
};
