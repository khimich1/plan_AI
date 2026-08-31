import { describe, expect, it } from "vitest";
import type { GsmWaybill } from "@/features/gsm/types/gsm";
import {
  buildVehicleDayCells,
  layoutWeekSlots,
} from "@/features/gsm/lib/vehicleDayMap";

const wb = (overrides: Partial<GsmWaybill>): GsmWaybill => ({
  id: 1,
  vehicle_id: 1,
  date: "2026-08-03",
  driver_id: 1,
  status: "draft",
  source: "auto",
  odometer_start: null,
  odometer_end: null,
  fuel_start: null,
  fuel_issued: null,
  fuel_end: null,
  km: 0,
  route: [],
  warnings: [],
  ...overrides,
});

describe("buildVehicleDayCells", () => {
  it("marks fuel without PL as gap", () => {
    const cells = buildVehicleDayCells(
      "2026-08-05",
      "2026-08-05",
      [],
      [{ ts: "2026-08-05T10:00:00", service_type: "fuel" }],
    );
    expect(cells).toHaveLength(1);
    expect(cells[0]).toMatchObject({
      date: "2026-08-05",
      hasTx: true,
      hasPl: false,
      isGap: true,
      isRed: false,
      waybill: null,
    });
  });

  it("marks wash without PL as gap", () => {
    const cells = buildVehicleDayCells(
      "2026-08-05",
      "2026-08-05",
      [],
      [{ ts: "2026-08-05T12:00:00", service_type: "wash" }],
    );
    expect(cells[0].isGap).toBe(true);
    expect(cells[0].hasTx).toBe(true);
  });

  it("does not treat other as an anchor gap", () => {
    const cells = buildVehicleDayCells(
      "2026-08-05",
      "2026-08-05",
      [],
      [{ ts: "2026-08-05T12:00:00", service_type: "other" }],
    );
    expect(cells[0].hasTx).toBe(false);
    expect(cells[0].isGap).toBe(false);
  });

  it("does not mark PL-only day as gap", () => {
    const cells = buildVehicleDayCells(
      "2026-08-03",
      "2026-08-03",
      [wb({ date: "2026-08-03" })],
      [],
    );
    expect(cells[0]).toMatchObject({
      hasTx: false,
      hasPl: true,
      isGap: false,
      isRed: false,
    });
  });

  it("marks PL with manual_intervention as red, not gap even with tx", () => {
    const cells = buildVehicleDayCells(
      "2026-08-03",
      "2026-08-03",
      [wb({ date: "2026-08-03", warnings: ["manual_intervention"] })],
      [{ ts: "2026-08-03T09:00:00", service_type: "fuel" }],
    );
    expect(cells[0]).toMatchObject({
      hasTx: true,
      hasPl: true,
      isGap: false,
      isRed: true,
    });
  });

  it("picks the first waybill by id when several share a day", () => {
    const cells = buildVehicleDayCells(
      "2026-08-03",
      "2026-08-03",
      [
        wb({ id: 20, date: "2026-08-03" }),
        wb({ id: 5, date: "2026-08-03" }),
        wb({ id: 10, date: "2026-08-03" }),
      ],
      [],
    );
    expect(cells[0].waybill?.id).toBe(5);
  });

  it("collapses multiple fuel/wash txs into one hasTx marker", () => {
    const cells = buildVehicleDayCells(
      "2026-08-05",
      "2026-08-05",
      [],
      [
        { ts: "2026-08-05T08:00:00", service_type: "fuel" },
        { ts: "2026-08-05T18:00:00", service_type: "wash" },
      ],
    );
    expect(cells).toHaveLength(1);
    expect(cells[0].hasTx).toBe(true);
  });

  it("returns empty list when from > to", () => {
    expect(buildVehicleDayCells("2026-08-10", "2026-08-01", [], [])).toEqual([]);
  });

  it("walks inclusive day range", () => {
    const cells = buildVehicleDayCells("2026-08-01", "2026-08-03", [], []);
    expect(cells.map((c) => c.date)).toEqual([
      "2026-08-01",
      "2026-08-02",
      "2026-08-03",
    ]);
  });

  it("uses ts first 10 chars for transaction date", () => {
    const cells = buildVehicleDayCells(
      "2026-08-05",
      "2026-08-05",
      [],
      [{ ts: "2026-08-05 23:50:00", service_type: "fuel" }],
    );
    expect(cells[0].hasTx).toBe(true);
  });
});

describe("layoutWeekSlots", () => {
  it("pads August 2026 with empty slots before Sat 01 and after Mon 31", () => {
    const cells = buildVehicleDayCells("2026-08-01", "2026-08-31", [], []);
    const slots = layoutWeekSlots(cells);

    // Aug 1 2026 = Saturday → 5 leading empty slots (Mon–Fri)
    expect(slots.slice(0, 5).every((s) => s === null)).toBe(true);
    expect(slots[5]?.date).toBe("2026-08-01");

    // Aug 31 2026 = Monday → 6 trailing empty slots (Tue–Sun)
    const lastCellIdx = slots.findLastIndex((s) => s !== null);
    expect(slots[lastCellIdx]?.date).toBe("2026-08-31");
    expect(slots.slice(lastCellIdx + 1).every((s) => s === null)).toBe(true);
    expect(slots.length - lastCellIdx - 1).toBe(6);

    // No slot carries a date outside [from, to]
    for (const slot of slots) {
      if (slot === null) continue;
      expect(slot.date >= "2026-08-01").toBe(true);
      expect(slot.date <= "2026-08-31").toBe(true);
    }

    // Full weeks of 7
    expect(slots.length % 7).toBe(0);
  });
});
