import { describe, expect, it } from "vitest";
import type { SgpPlateItem } from "@/features/production/api/sgpApi";
import {
  UNLINKED_GROUP_KEY,
  groupPlatesByKp,
  groupKeyForPlate,
} from "@/features/production/lib/sgpWarehouseGroups";

function makePlate(overrides: Partial<SgpPlateItem> & Pick<SgpPlateItem, "id">): SgpPlateItem {
  return {
    kp_id: 1,
    plate_name: "Плита",
    length_m: 4.5,
    width_m: 1.2,
    load_class: 600,
    qty: 1,
    completed_date: null,
    production_day: null,
    plan_id: null,
    nomenclature_id: null,
    customer_name: "Клиент",
    execution_terms: "03.08.2026",
    sgp_progress: { n: 1, m: 10 },
    ...overrides,
  };
}

describe("sgpWarehouseGroups", () => {
  it("groups plates by kp_id", () => {
    const groups = groupPlatesByKp([
      makePlate({ id: 1, kp_id: 2, qty: 3 }),
      makePlate({ id: 2, kp_id: 1, qty: 1 }),
      makePlate({ id: 3, kp_id: 1, qty: 2 }),
    ]);

    expect(groups).toHaveLength(2);
    expect(groups[0].kpId).toBe(1);
    expect(groups[0].positionCount).toBe(2);
    expect(groups[0].totalQty).toBe(3);
    expect(groups[1].kpId).toBe(2);
    expect(groups[1].totalQty).toBe(3);
  });

  it("places unlinked group first", () => {
    const groups = groupPlatesByKp([
      makePlate({ id: 1, kp_id: 5 }),
      makePlate({ id: 2, kp_id: null, sgp_progress: null }),
    ]);

    expect(groups[0].key).toBe(UNLINKED_GROUP_KEY);
    expect(groups[1].key).toBe(5);
  });

  it("resolves group key for unlinked plate", () => {
    expect(groupKeyForPlate(makePlate({ id: 1, kp_id: null }))).toBe(UNLINKED_GROUP_KEY);
    expect(groupKeyForPlate(makePlate({ id: 2, kp_id: 7 }))).toBe(7);
  });
});
