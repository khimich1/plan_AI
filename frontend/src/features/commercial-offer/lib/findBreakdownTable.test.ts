import { describe, expect, it } from "vitest";
import { findBreakdownTable, normalizePlateName } from "@/features/commercial-offer/lib/findBreakdownTable";
import type { BreakdownTable } from "@/features/commercial-offer/types/commercialOffer";

const tables: BreakdownTable[] = [
  {
    name: "Плиты ПБ 72,8-12-8п",
    rows: [["Базовая цена", "x", "1 руб"]],
  },
  {
    name: "Плиты ПБ 72,8-8-8п (нагрузка?)",
    rows: [["Продольный рез", "y", "2 руб"]],
  },
];

describe("normalizePlateName", () => {
  it("removes load warning suffix", () => {
    expect(normalizePlateName("Плиты ПБ 72,8-8-8п (нагрузка?)")).toBe("Плиты ПБ 72,8-8-8п");
  });
});

describe("findBreakdownTable", () => {
  it("finds table by exact name", () => {
    const found = findBreakdownTable(tables, "Плиты ПБ 72,8-12-8п");
    expect(found?.name).toBe("Плиты ПБ 72,8-12-8п");
  });

  it("matches order name to breakdown with warning suffix", () => {
    const found = findBreakdownTable(tables, "Плиты ПБ 72,8-8-8п");
    expect(found?.rows[0][0]).toBe("Продольный рез");
  });

  it("returns undefined when no match", () => {
    expect(findBreakdownTable(tables, "Плиты ПБ 99-12-8п")).toBeUndefined();
  });
});
