import type { BatchDraft, OfferPlateForSchedule } from "@/features/delivery-schedule/lib/scheduleDraft";

export const SECTION_IN_SCHEDULE = "Уже в поставках";
export const SECTION_REMAINDER = "Остаток";

export type TemplateSectionTitle = typeof SECTION_IN_SCHEDULE | typeof SECTION_REMAINDER;

export type TemplateRow =
  | { kind: "section"; title: TemplateSectionTitle }
  | {
      kind: "item";
      batchName: string;
      deliverFrom: string;
      deliverTo: string;
      produceBy: string;
      plateName: string;
      qty: number;
    };

const emptyItemDates = {
  batchName: "",
  deliverFrom: "",
  deliverTo: "",
  produceBy: "",
} as const;

const plateMark = (plate: OfferPlateForSchedule | undefined): string =>
  (plate?.plate_name ?? "").trim();

/** ISO YYYY-MM-DD → ДД.ММ.ГГГГ; пустое или не-ISO оставляем как есть. */
export const isoToDisplayDate = (value: string): string => {
  const text = value.trim();
  if (!text) {
    return "";
  }
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!match) {
    return text;
  }
  return `${match[3]}.${match[2]}.${match[1]}`;
};

export function buildScheduleTemplateRows(
  plates: OfferPlateForSchedule[],
  batches: BatchDraft[],
): TemplateRow[] {
  const plateById = new Map(plates.map((plate) => [plate.id, plate]));
  const allocated = new Map<number, number>();
  const scheduled: TemplateRow[] = [];

  for (const draft of batches) {
    for (const item of draft.items) {
      if (item.qty < 1) {
        continue;
      }
      const plate = plateById.get(item.plate_id);
      const mark = plateMark(plate);
      if (!mark) {
        continue;
      }
      allocated.set(item.plate_id, (allocated.get(item.plate_id) ?? 0) + item.qty);
      scheduled.push({
        kind: "item",
        batchName: draft.name,
        deliverFrom: isoToDisplayDate(draft.deliver_from),
        deliverTo: isoToDisplayDate(draft.deliver_to),
        produceBy: isoToDisplayDate(draft.produce_by),
        plateName: mark,
        qty: item.qty,
      });
    }
  }

  const remainder: TemplateRow[] = [];
  for (const plate of plates) {
    const mark = plateMark(plate);
    if (!mark) {
      continue;
    }
    const leftover = plate.qty - (allocated.get(plate.id) ?? 0);
    if (leftover <= 0) {
      continue;
    }
    remainder.push({
      kind: "item",
      ...emptyItemDates,
      plateName: mark,
      qty: leftover,
    });
  }

  const rows: TemplateRow[] = [];
  if (scheduled.length > 0) {
    rows.push({ kind: "section", title: SECTION_IN_SCHEDULE }, ...scheduled);
  }
  if (remainder.length > 0) {
    rows.push({ kind: "section", title: SECTION_REMAINDER }, ...remainder);
  }
  return rows;
}
