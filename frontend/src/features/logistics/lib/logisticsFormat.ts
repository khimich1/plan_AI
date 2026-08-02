import type {
  BoolFlag,
  DeliveryType,
  ShipmentStatus,
  VehicleClass,
} from "@/features/logistics/types/logistics";

export const VEHICLE_CLASS_LIMITS_FALLBACK_KG: Record<VehicleClass, number> = {
  t20: 19_800,
  t30plus: 30_000,
};

export const isFlagOn = (value: BoolFlag | null | undefined): boolean => Boolean(value);

export const deliveryTypeLabel = (value: DeliveryType): string =>
  value === "pickup" ? "Самовывоз" : "Доставка";

export const deliveryTypeShort = (value: DeliveryType): string =>
  value === "pickup" ? "С" : "Д";

export const shipmentStatusLabel = (value: ShipmentStatus): string =>
  value === "done" ? "Обработано" : "В работе";

export const vehicleClassLabel = (value: VehicleClass | null | undefined): string => {
  if (value === "t20") return "до 19,8 т";
  if (value === "t30plus") return "30 т+";
  return "—";
};

export const formatWeightKg = (value: number | null | undefined): string => {
  if (value == null || !Number.isFinite(value)) {
    return "—";
  }
  return `${Math.round(value).toLocaleString("ru-RU")} кг`;
};

export const formatCost = (value: number | null | undefined): string => {
  if (value == null || !Number.isFinite(value)) {
    return "—";
  }
  return `${value.toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽`;
};

export const formatDate = (iso: string | null | undefined): string => {
  if (!iso) {
    return "—";
  }
  const [y, m, d] = iso.slice(0, 10).split("-");
  if (!y || !m || !d) {
    return iso;
  }
  return `${d}.${m}.${y}`;
};

export const todayIsoDate = (): string => {
  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${now.getFullYear()}-${month}-${day}`;
};
