import { describe, expect, it } from "vitest";
import {
  draftFromSaved,
  draftFreeRow,
  draftRowWeightKg,
  draftTotalWeightKg,
  draftsToPayload,
  savedItemsVersion,
} from "@/features/logistics/lib/draftItems";
import type { ShipmentItem } from "@/features/logistics/types/logistics";

const plateItem = (overrides: Partial<ShipmentItem> = {}): ShipmentItem => ({
  id: 1,
  item_type: "plate",
  completed_plate_id: 10,
  kp_id: 100,
  mark: null,
  plate_name: "ПБ 60-12-8п",
  length_m: 6,
  width_m: 1.2,
  load_class: 800,
  qty: 2,
  unit_weight_kg: 1500,
  weight_kg: 3000,
  sort_order: 0,
  note: null,
  ...overrides,
});

const freeItem = (overrides: Partial<ShipmentItem> = {}): ShipmentItem => ({
  id: 2,
  item_type: "free",
  completed_plate_id: null,
  kp_id: 100,
  mark: "С60.30",
  plate_name: null,
  length_m: null,
  width_m: null,
  load_class: null,
  qty: 3,
  unit_weight_kg: 1060,
  weight_kg: 3180,
  sort_order: 1,
  note: "сваи",
  ...overrides,
});

describe("draftFromSaved", () => {
  it("maps plate rows and keeps server weight", () => {
    const draft = draftFromSaved(plateItem());
    expect(draft.item_type).toBe("plate");
    expect(draft.completed_plate_id).toBe(10);
    expect(draft.qty).toBe(2);
    expect(draft.weight_kg).toBe(3000);
    expect(draft.weight_manual).toBe(false);
    expect(draft.dims).toBe("6×1.2 / 800");
  });

  it("marks free row as manual when weight differs from unit×qty", () => {
    const draft = draftFromSaved(freeItem({ weight_kg: 4000 }));
    expect(draft.weight_manual).toBe(true);
    expect(draft.weight_kg).toBe(4000);
  });

  it("does not mark free row manual when weight matches auto", () => {
    const draft = draftFromSaved(freeItem({ weight_kg: 3180 }));
    expect(draft.weight_manual).toBe(false);
    expect(draft.weight_kg).toBeNull();
  });
});

describe("draftsToPayload", () => {
  it("serializes plate and free rows with sort_order", () => {
    const plate = draftFromSaved(plateItem());
    const free = draftFromSaved(freeItem({ weight_kg: 4000 }));
    const payload = draftsToPayload([plate, free]);

    expect(payload).toEqual([
      {
        item_type: "plate",
        completed_plate_id: 10,
        kp_id: 100,
        mark: undefined,
        qty: 2,
        weight_kg: undefined,
        sort_order: 0,
        note: undefined,
      },
      {
        item_type: "free",
        completed_plate_id: undefined,
        kp_id: 100,
        mark: "С60.30",
        qty: 3,
        weight_kg: 4000,
        sort_order: 1,
        note: "сваи",
      },
    ]);
  });

  it("omits manual weight when not set on free rows", () => {
    const free = draftFreeRow(55);
    free.mark = "С80";
    free.qty = 1;
    const payload = draftsToPayload([free]);
    expect(payload[0].weight_kg).toBeUndefined();
    expect(payload[0].kp_id).toBe(55);
  });
});

describe("draft weight totals", () => {
  it("uses unit×qty when weight_kg is null", () => {
    const item = draftFreeRow(1);
    item.unit_weight_kg = 1000;
    item.qty = 2;
    item.weight_kg = null;
    expect(draftRowWeightKg(item)).toBe(2000);
  });

  it("sums row weights for total", () => {
    const a = draftFromSaved(plateItem({ qty: 1, unit_weight_kg: 1500, weight_kg: 1500 }));
    const b = draftFromSaved(freeItem({ weight_kg: 4000 }));
    expect(draftTotalWeightKg([a, b])).toBe(5500);
  });
});

describe("savedItemsVersion", () => {
  it("changes when qty or weight changes", () => {
    const a = savedItemsVersion([plateItem()]);
    const b = savedItemsVersion([plateItem({ qty: 3 })]);
    expect(a).not.toEqual(b);
  });
});
