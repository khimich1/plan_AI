import { describe, expect, it } from "vitest";
import {
  burnForKm,
  libraryRouteToLegs,
  previousWaybillChainStart,
  previewDownstream,
} from "@/features/gsm/lib/downstreamPreview";
import type { GsmVehicle, GsmWaybill } from "@/features/gsm/types/gsm";

const VEHICLE: Pick<GsmVehicle, "norm_summer" | "norm_winter" | "tank_volume_liters"> = {
  norm_summer: 10,
  norm_winter: 12,
  tank_volume_liters: 60,
};

const WAYBILLS: GsmWaybill[] = [
  {
    id: 1,
    vehicle_id: 1,
    date: "2025-04-03",
    driver_id: 7,
    status: "draft",
    source: "auto",
    odometer_start: 10000,
    odometer_end: 10200,
    fuel_start: 20,
    fuel_issued: 40,
    fuel_end: 40,
    km: 200,
    route: [{ from: "A", to: "B", km: 200 }],
    warnings: [],
  },
  {
    id: 2,
    vehicle_id: 1,
    date: "2025-04-04",
    driver_id: 7,
    status: "draft",
    source: "auto",
    odometer_start: 10200,
    odometer_end: 10400,
    fuel_start: 40,
    fuel_issued: 0,
    fuel_end: 20,
    km: 200,
    route: [{ from: "A", to: "C", km: 200 }],
    warnings: [],
  },
];

describe("downstreamPreview", () => {
  it("burnForKm matches backend rounding", () => {
    expect(burnForKm(190, 9.4)).toBe(17.86);
  });

  it("previews edited day and rechains following drafts", () => {
    const result = previewDownstream({
      edited: WAYBILLS[0],
      editedKm: 100,
      periodWaybills: WAYBILLS,
      vehicle: VEHICLE,
      winterStartMMDD: "11-01",
    });

    expect(result.error).toBeNull();
    const edited = result.days.find((d) => d.id === 1)!;
    expect(edited.km).toBe(100);
    expect(edited.fuel_end).toBe(50);
    expect(edited.odometer_end).toBe(10100);
    expect(edited.changed).toBe(true);

    const next = result.days.find((d) => d.id === 2)!;
    expect(next.fuel_start).toBe(50);
    expect(next.odometer_start).toBe(10100);
    expect(next.fuel_end).toBe(30);
    expect(next.odometer_end).toBe(10300);
    expect(next.changed).toBe(true);
  });

  it("leaves confirmed days untouched while continuing chain after them", () => {
    const confirmed: GsmWaybill = {
      ...WAYBILLS[1],
      status: "confirmed",
      fuel_start: 40,
      fuel_end: 20,
      odometer_start: 10200,
      odometer_end: 10400,
    };
    const result = previewDownstream({
      edited: WAYBILLS[0],
      editedKm: 100,
      periodWaybills: [WAYBILLS[0], confirmed],
      vehicle: VEHICLE,
      winterStartMMDD: "11-01",
    });
    const day2 = result.days.find((d) => d.id === 2)!;
    expect(day2.changed).toBe(false);
    expect(day2.fuel_end).toBe(20);
  });

  it("previousWaybillChainStart picks latest day before date", () => {
    expect(previousWaybillChainStart(WAYBILLS, "2025-04-04")).toEqual({
      fuel_start: 40,
      odometer_start: 10200,
    });
    expect(previousWaybillChainStart(WAYBILLS, "2025-04-03")).toEqual({
      fuel_start: null,
      odometer_start: null,
    });
  });

  it("libraryRouteToLegs maps library row to waybill leg", () => {
    expect(
      libraryRouteToLegs({
        id: 9,
        addr_a: "Завод",
        addr_b: "Объект",
        km: 42,
        typical_station_ids: [3],
      }),
    ).toEqual([
      {
        from: "Завод",
        to: "Объект",
        km: 42,
        route_id: 9,
        station_id: 3,
      },
    ]);
  });
});
