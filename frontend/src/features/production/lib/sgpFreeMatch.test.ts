import { describe, expect, it } from "vitest";
import {
  freeQtyForPlate,
  pickFreeReservations,
  plateMatchesFree,
} from "@/features/production/lib/sgpFreeMatch";

const plate = {
  id: 10,
  plate_name: "ПБ 60-12-8п",
  length_m: 6,
  width_m: 1.2,
  load_class: 800,
  qty: 5,
};

describe("sgpFreeMatch", () => {
  it("matches strict identity", () => {
    expect(
      plateMatchesFree(plate, {
        id: 1,
        plate_name: "ПБ 60-12-8п",
        length_m: 6,
        width_m: 1.2,
        load_class: 800,
        qty: 3,
        completed_date: null,
      }),
    ).toBe(true);
    expect(
      plateMatchesFree(plate, {
        id: 2,
        plate_name: "ПБ 60-12-8п",
        length_m: 5,
        width_m: 1.2,
        load_class: 800,
        qty: 3,
        completed_date: null,
      }),
    ).toBe(false);
  });

  it("picks FIFO reservations min(free, demand)", () => {
    const free = [
      {
        id: 1,
        plate_name: "ПБ 60-12-8п",
        length_m: 6,
        width_m: 1.2,
        load_class: 800,
        qty: 2,
        completed_date: "01.01.2026",
      },
      {
        id: 2,
        plate_name: "ПБ 60-12-8п",
        length_m: 6,
        width_m: 1.2,
        load_class: 800,
        qty: 4,
        completed_date: "02.01.2026",
      },
    ];
    expect(freeQtyForPlate(plate, free)).toBe(6);
    expect(pickFreeReservations(plate, free, 3)).toEqual([
      { sgp_id: 1, qty: 2 },
      { sgp_id: 2, qty: 1 },
    ]);
  });
});
