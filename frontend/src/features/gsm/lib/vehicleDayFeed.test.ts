import { describe, expect, it } from "vitest";
import {
  buildVehicleDayFeed,
  monthBounds,
  shiftMonth,
} from "@/features/gsm/lib/vehicleDayFeed";
import type { GsmTransaction, GsmWaybill } from "@/features/gsm/types/gsm";

const wb = (overrides: Partial<GsmWaybill> = {}): GsmWaybill => ({
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

const tx = (overrides: Partial<GsmTransaction> = {}): GsmTransaction => ({
  ts: "2026-08-05T10:00:00",
  card_number: "1",
  vehicle_id: 1,
  service_type: "fuel",
  fuel_grade: "ДТ",
  qty_liters: 40,
  amount: 1000,
  station_id: null,
  address: "АЗС 1",
  ...overrides,
});

describe("monthBounds", () => {
  it("returns first and last day of a 31-day month", () => {
    expect(monthBounds("2026-08")).toEqual({ from: "2026-08-01", to: "2026-08-31" });
  });

  it("returns 28 days for February in a non-leap year", () => {
    expect(monthBounds("2026-02")).toEqual({ from: "2026-02-01", to: "2026-02-28" });
  });

  it("returns 29 days for February in a leap year", () => {
    expect(monthBounds("2024-02")).toEqual({ from: "2024-02-01", to: "2024-02-29" });
  });

  it("returns 30 days for April", () => {
    expect(monthBounds("2026-04")).toEqual({ from: "2026-04-01", to: "2026-04-30" });
  });
});

describe("shiftMonth", () => {
  it("shifts forward within a year", () => {
    expect(shiftMonth("2026-01", 1)).toBe("2026-02");
  });

  it("shifts backward across the year boundary", () => {
    expect(shiftMonth("2026-01", -1)).toBe("2025-12");
  });

  it("shifts forward across the year boundary", () => {
    expect(shiftMonth("2026-12", 1)).toBe("2027-01");
  });
});

describe("buildVehicleDayFeed", () => {
  it("marks a fuel tx without a waybill as a gap", () => {
    const feed = buildVehicleDayFeed("2026-08-01", "2026-08-31", [], [tx()]);
    expect(feed).toHaveLength(1);
    expect(feed[0].date).toBe("2026-08-05");
    expect(feed[0].isGap).toBe(true);
    expect(feed[0].txs).toHaveLength(1);
    expect(feed[0].waybills).toHaveLength(0);
  });

  it("marks a wash tx without a waybill as a gap", () => {
    const feed = buildVehicleDayFeed(
      "2026-08-01",
      "2026-08-31",
      [],
      [tx({ service_type: "wash", qty_liters: null, amount: 500 })],
    );
    expect(feed[0].isGap).toBe(true);
  });

  it("does not mark an 'other' tx without a waybill as a gap", () => {
    const feed = buildVehicleDayFeed(
      "2026-08-01",
      "2026-08-31",
      [],
      [tx({ service_type: "other", qty_liters: null })],
    );
    expect(feed[0].isGap).toBe(false);
  });

  it("does not mark a fuel tx as a gap when a waybill exists the same day", () => {
    const feed = buildVehicleDayFeed(
      "2026-08-01",
      "2026-08-31",
      [wb({ id: 5, date: "2026-08-05" })],
      [tx()],
    );
    expect(feed[0].isGap).toBe(false);
    expect(feed[0].txs).toHaveLength(1);
    expect(feed[0].waybills.map((w) => w.id)).toEqual([5]);
  });

  it("renders a waybill-only day as a regular (non-gap) section", () => {
    const feed = buildVehicleDayFeed(
      "2026-08-01",
      "2026-08-31",
      [wb({ id: 7, date: "2026-08-10" })],
      [],
    );
    expect(feed).toHaveLength(1);
    expect(feed[0].date).toBe("2026-08-10");
    expect(feed[0].isGap).toBe(false);
    expect(feed[0].txs).toHaveLength(0);
  });

  it("sorts txs by ts and waybills by id inside a day", () => {
    const feed = buildVehicleDayFeed(
      "2026-08-01",
      "2026-08-31",
      [wb({ id: 9, date: "2026-08-05" }), wb({ id: 3, date: "2026-08-05" })],
      [
        tx({ ts: "2026-08-05T18:30:00", amount: 2 }),
        tx({ ts: "2026-08-05T08:15:00", amount: 1 }),
      ],
    );
    expect(feed[0].txs.map((t) => t.ts)).toEqual([
      "2026-08-05T08:15:00",
      "2026-08-05T18:30:00",
    ]);
    expect(feed[0].waybills.map((w) => w.id)).toEqual([3, 9]);
  });

  it("splits events by day and orders sections by date ascending", () => {
    const feed = buildVehicleDayFeed(
      "2026-08-01",
      "2026-08-31",
      [wb({ id: 1, date: "2026-08-20" })],
      [tx({ ts: "2026-08-05 10:00:00" })],
    );
    expect(feed.map((d) => d.date)).toEqual(["2026-08-05", "2026-08-20"]);
  });

  it("skips days without events", () => {
    const feed = buildVehicleDayFeed("2026-08-01", "2026-08-03", [], []);
    expect(feed).toEqual([]);
  });

  it("ignores events outside the requested range", () => {
    const feed = buildVehicleDayFeed(
      "2026-08-01",
      "2026-08-31",
      [wb({ id: 1, date: "2026-07-31" })],
      [tx({ ts: "2026-09-01T10:00:00" })],
    );
    expect(feed).toEqual([]);
  });

  it("returns an empty feed when from > to", () => {
    const feed = buildVehicleDayFeed("2026-08-31", "2026-08-01", [wb()], [tx()]);
    expect(feed).toEqual([]);
  });
});
