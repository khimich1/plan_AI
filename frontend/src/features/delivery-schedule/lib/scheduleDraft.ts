import type {
  BatchDraftOut,
  BatchIn,
  BatchOut,
  DeliverySchedulePut,
  DeliveryScheduleView,
  TrafficLightStatus,
} from "@/features/delivery-schedule/types/deliverySchedule";

export type OfferPlateForSchedule = {
  id: number;
  plate_name: string;
  qty: number;
  position_number?: number | null;
};

export type BatchDraftItem = {
  plate_id: number;
  qty: number;
};

export type BatchDraft = {
  /** Локальный ключ для React (не id БД). */
  key: string;
  name: string;
  deliver_from: string;
  deliver_to: string;
  produce_by: string;
  items: BatchDraftItem[];
  status: TrafficLightStatus | null;
  ready_date: string | null;
  hint: string | null;
  changed: boolean;
};

export type PositionSplit = {
  plate_id: number;
  plate_name: string;
  position_number: number | null;
  ordered: number;
  allocated: number;
  remaining: number;
};

let draftKeySeq = 0;

export const nextDraftKey = (): string => {
  draftKeySeq += 1;
  return `batch-${draftKeySeq}-${Date.now()}`;
};

export const emptyBatchDraft = (name = ""): BatchDraft => ({
  key: nextDraftKey(),
  name,
  deliver_from: "",
  deliver_to: "",
  produce_by: "",
  items: [],
  status: null,
  ready_date: null,
  hint: null,
  changed: false,
});

export const batchOutToDraft = (batch: BatchOut): BatchDraft => ({
  key: nextDraftKey(),
  name: batch.name,
  deliver_from: batch.deliver_from,
  deliver_to: batch.deliver_to,
  produce_by: batch.produce_by,
  items: batch.items.map((item) => ({ plate_id: item.plate_id, qty: item.qty })),
  status: batch.status,
  ready_date: batch.ready_date,
  hint: batch.hint,
  changed: batch.changed,
});

export const viewToDrafts = (view: DeliveryScheduleView | null | undefined): BatchDraft[] => {
  if (!view?.batches?.length) {
    return [];
  }
  return [...view.batches]
    .sort((a, b) => a.sort_order - b.sort_order || a.id - b.id)
    .map(batchOutToDraft);
};

/** Черновик с POST /import → локальные drafts редактора (светофор ещё не посчитан). */
export const importBatchesToDrafts = (batches: BatchDraftOut[]): BatchDraft[] =>
  batches.map((batch) => ({
    key: nextDraftKey(),
    name: batch.name,
    deliver_from: batch.deliver_from,
    deliver_to: batch.deliver_to,
    produce_by: batch.produce_by,
    items: batch.items.map((item) => ({ plate_id: item.plate_id, qty: item.qty })),
    status: null,
    ready_date: null,
    hint: null,
    changed: false,
  }));

export const allocatedByPlate = (batches: BatchDraft[]): Map<number, number> => {
  const map = new Map<number, number>();
  for (const batch of batches) {
    for (const item of batch.items) {
      if (item.qty <= 0) {
        continue;
      }
      map.set(item.plate_id, (map.get(item.plate_id) ?? 0) + item.qty);
    }
  }
  return map;
};

export const computePositionSplits = (
  plates: OfferPlateForSchedule[],
  batches: BatchDraft[],
): PositionSplit[] => {
  const allocated = allocatedByPlate(batches);
  return plates.map((plate) => {
    const used = allocated.get(plate.id) ?? 0;
    return {
      plate_id: plate.id,
      plate_name: plate.plate_name,
      position_number: plate.position_number ?? null,
      ordered: plate.qty,
      allocated: used,
      remaining: plate.qty - used,
    };
  });
};

/** Сколько позиций полностью разбито (allocated === ordered и ordered > 0). */
export const splitProgress = (
  splits: PositionSplit[],
): { allocatedPositions: number; totalPositions: number } => {
  const withQty = splits.filter((s) => s.ordered > 0);
  const fully = withQty.filter((s) => s.allocated === s.ordered).length;
  return { allocatedPositions: fully, totalPositions: withQty.length };
};

export type QtyValidationIssue = {
  plate_id: number;
  plate_name: string;
  ordered: number;
  allocated: number;
};

export const findQtyOverflows = (
  plates: OfferPlateForSchedule[],
  batches: BatchDraft[],
): QtyValidationIssue[] => {
  const splits = computePositionSplits(plates, batches);
  return splits
    .filter((s) => s.allocated > s.ordered)
    .map((s) => ({
      plate_id: s.plate_id,
      plate_name: s.plate_name,
      ordered: s.ordered,
      allocated: s.allocated,
    }));
};

export const draftsToPut = (
  batches: BatchDraft[],
  meta?: { invoice_number?: string | null; contract_number?: string | null },
): DeliverySchedulePut => {
  const payloadBatches: BatchIn[] = batches.map((batch, index) => ({
    name: batch.name.trim() || `Партия ${index + 1}`,
    deliver_from: batch.deliver_from,
    deliver_to: batch.deliver_to,
    produce_by: batch.produce_by,
    items: batch.items
      .filter((item) => item.qty >= 1)
      .map((item) => ({ plate_id: item.plate_id, qty: item.qty })),
    sort_order: index,
  }));
  return {
    invoice_number: meta?.invoice_number ?? null,
    contract_number: meta?.contract_number ?? null,
    batches: payloadBatches,
  };
};

export const validateBatchDates = (batch: BatchDraft): string | null => {
  if (!batch.deliver_from || !batch.deliver_to || !batch.produce_by) {
    return "Укажите все три даты партии";
  }
  if (batch.deliver_from > batch.deliver_to) {
    return "«Поставка с» не может быть позже «Поставка по»";
  }
  return null;
};
