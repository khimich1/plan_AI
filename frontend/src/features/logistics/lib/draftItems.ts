import type {
  AvailablePlate,
  PileCatalogEntry,
  ProposedItem,
  ShipmentItem,
  ShipmentItemInput,
  ShipmentItemType,
} from "@/features/logistics/types/logistics";

export type DraftItem = {
  key: string;
  item_type: ShipmentItemType;
  completed_plate_id: number | null;
  kp_id: number | null;
  mark: string;
  plate_name: string;
  dims: string;
  qty: number;
  unit_weight_kg: number | null;
  weight_kg: number | null;
  weight_manual: boolean;
  note: string;
};

let draftKeyCounter = 0;
const nextKey = (): string => {
  draftKeyCounter += 1;
  return `draft-${draftKeyCounter}`;
};

const formatDims = (
  lengthM: number | null,
  widthM: number | null,
  loadClass: number | null,
): string => {
  const l = lengthM != null ? `${lengthM}` : "—";
  const w = widthM != null ? `${widthM}` : "—";
  const load = loadClass != null ? String(loadClass) : "—";
  return `${l}×${w} / ${load}`;
};

export const draftItemLabel = (item: DraftItem): string =>
  item.item_type === "plate" ? item.plate_name : item.mark || "Свободная строка";

export const draftRowWeightKg = (item: DraftItem): number | null => {
  if (item.weight_kg != null) {
    return item.weight_kg;
  }
  if (item.unit_weight_kg != null) {
    return item.unit_weight_kg * item.qty;
  }
  return null;
};

export const draftTotalWeightKg = (items: DraftItem[]): number =>
  items.reduce((sum, item) => sum + (draftRowWeightKg(item) ?? 0), 0);

export const draftFromSaved = (item: ShipmentItem): DraftItem => {
  const autoWeight = item.unit_weight_kg != null ? item.unit_weight_kg * item.qty : null;
  // Сервер не помечает ручной вес: если weight ≠ unit×qty, считаем правку ручной
  const freeManual =
    item.item_type === "free" &&
    item.weight_kg != null &&
    (autoWeight == null || Math.abs(item.weight_kg - autoWeight) > 0.5);
  return {
    key: `saved-${item.id}`,
    item_type: item.item_type,
    completed_plate_id: item.completed_plate_id,
    kp_id: item.kp_id,
    mark: item.mark ?? "",
    plate_name: item.plate_name ?? "",
    dims: formatDims(item.length_m, item.width_m, item.load_class),
    qty: item.qty,
    unit_weight_kg: item.unit_weight_kg,
    weight_kg: item.item_type === "plate" || freeManual ? item.weight_kg : null,
    weight_manual: freeManual,
    note: item.note ?? "",
  };
};

export const draftFromProposed = (item: ProposedItem): DraftItem => ({
  key: nextKey(),
  item_type: item.item_type,
  completed_plate_id: item.completed_plate_id,
  kp_id: item.kp_id,
  mark: "",
  plate_name: item.plate_name,
  dims: formatDims(item.length_m, item.width_m, item.load_class),
  qty: item.qty,
  unit_weight_kg: item.unit_weight_kg,
  weight_kg: item.weight_kg,
  weight_manual: false,
  note: "",
});

export const draftFromAvailablePlate = (plate: AvailablePlate, kpId: number): DraftItem => ({
  key: nextKey(),
  item_type: "plate",
  completed_plate_id: plate.completed_plate_id,
  kp_id: kpId,
  mark: "",
  plate_name: plate.plate_name,
  dims: formatDims(plate.length_m, plate.width_m, plate.load_class),
  qty: 1,
  unit_weight_kg: plate.unit_weight_kg,
  weight_kg: null,
  weight_manual: false,
  note: "",
});

export const draftFreeRow = (kpId: number | null): DraftItem => ({
  key: nextKey(),
  item_type: "free",
  completed_plate_id: null,
  kp_id: kpId,
  mark: "",
  plate_name: "",
  dims: "",
  qty: 1,
  unit_weight_kg: null,
  weight_kg: null,
  weight_manual: false,
  note: "",
});

export const applyPileCatalogEntry = (item: DraftItem, entry: PileCatalogEntry): DraftItem => ({
  ...item,
  mark: entry.mark,
  unit_weight_kg: entry.weight_kg,
  weight_kg: item.weight_manual ? item.weight_kg : null,
});

export const draftsToPayload = (items: DraftItem[]): ShipmentItemInput[] =>
  items.map((item, index) => ({
    item_type: item.item_type,
    completed_plate_id: item.item_type === "plate" ? item.completed_plate_id : undefined,
    kp_id: item.kp_id ?? undefined,
    mark: item.item_type === "free" && item.mark.trim() ? item.mark.trim() : undefined,
    qty: item.qty,
    weight_kg:
      item.item_type === "free" && item.weight_manual && item.weight_kg != null
        ? item.weight_kg
        : undefined,
    sort_order: index,
    note: item.note.trim() ? item.note.trim() : undefined,
  }));

/** Ключ сохранённого состава: меняется при refetch карточки после confirm — триггер переинициализации черновика. */
export const savedItemsVersion = (items: ShipmentItem[]): string =>
  JSON.stringify(
    items.map((i) => [i.id, i.item_type, i.completed_plate_id, i.qty, i.weight_kg, i.note, i.sort_order]),
  );
