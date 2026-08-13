import { describe, expect, it } from "vitest";
import {
  allocatedByPlate,
  computePositionSplits,
  draftsToPut,
  findQtyOverflows,
  importBatchesToDrafts,
  splitProgress,
  type BatchDraft,
  type OfferPlateForSchedule,
} from "@/features/delivery-schedule/lib/scheduleDraft";
import { validateScheduleEditor } from "@/features/delivery-schedule/components/DeliveryScheduleEditor";

const plates: OfferPlateForSchedule[] = [
  { id: 1, plate_name: "ПБ 59-12-8", qty: 10, position_number: 1 },
  { id: 2, plate_name: "ПБ 51-12-8", qty: 5, position_number: 2 },
];

const batch = (overrides: Partial<BatchDraft> & { items: BatchDraft["items"] }): BatchDraft => ({
  key: "k1",
  name: "Партия 1",
  deliver_from: "2026-08-10",
  deliver_to: "2026-08-12",
  produce_by: "2026-08-08",
  status: null,
  ready_date: null,
  hint: null,
  changed: false,
  ...overrides,
});

describe("scheduleDraft", () => {
  it("computes remaining per plate", () => {
    const batches = [
      batch({ items: [{ plate_id: 1, qty: 4 }] }),
      batch({ key: "k2", name: "П2", items: [{ plate_id: 1, qty: 3 }, { plate_id: 2, qty: 5 }] }),
    ];
    const splits = computePositionSplits(plates, batches);
    expect(splits[0]).toMatchObject({ allocated: 7, remaining: 3 });
    expect(splits[1]).toMatchObject({ allocated: 5, remaining: 0 });
    expect(splitProgress(splits)).toEqual({ allocatedPositions: 1, totalPositions: 2 });
  });

  it("detects qty overflow before save", () => {
    const batches = [batch({ items: [{ plate_id: 1, qty: 11 }] })];
    expect(findQtyOverflows(plates, batches)).toHaveLength(1);
    expect(validateScheduleEditor(plates, batches)).toMatch(/превышает qty/);
  });

  it("allows partial split (Σ ≤ qty)", () => {
    const batches = [batch({ items: [{ plate_id: 1, qty: 10 }] })];
    expect(allocatedByPlate(batches).get(1)).toBe(10);
    expect(validateScheduleEditor(plates, batches)).toBeNull();
    expect(draftsToPut(batches).batches[0].items).toEqual([{ plate_id: 1, qty: 10 }]);
  });

  it("requires dates and name", () => {
    expect(
      validateScheduleEditor(plates, [
        batch({ name: "", items: [{ plate_id: 1, qty: 1 }] }),
      ]),
    ).toMatch(/название/i);
    expect(
      validateScheduleEditor(plates, [
        batch({ deliver_from: "", items: [{ plate_id: 1, qty: 1 }] }),
      ]),
    ).toMatch(/даты/i);
  });

  it("maps import draft batches into editor drafts", () => {
    const drafts = importBatchesToDrafts([
      {
        name: "Импорт 1",
        deliver_from: "2026-09-01",
        deliver_to: "2026-09-03",
        produce_by: "2026-08-28",
        items: [{ plate_id: 1, plate_name: "ПБ 59-12-8", qty: 4 }],
      },
    ]);
    expect(drafts).toHaveLength(1);
    expect(drafts[0]).toMatchObject({
      name: "Импорт 1",
      deliver_from: "2026-09-01",
      items: [{ plate_id: 1, qty: 4 }],
      status: null,
      changed: false,
    });
    expect(drafts[0].key).toBeTruthy();
  });
});
