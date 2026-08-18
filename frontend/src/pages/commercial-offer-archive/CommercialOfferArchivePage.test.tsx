import { describe, expect, it } from "vitest";
import {
  concreteProductTypes,
  filterByProductType,
  sectionFromStatus,
} from "@/pages/commercial-offer-archive/CommercialOfferArchivePage";
import type { ArchiveOfferListItem } from "@/features/commercial-archive/types/archive";

const baseItem: ArchiveOfferListItem = {
  kp_id: 1,
  creation_date: "01.03.2026",
  customer_name: "Клиент",
  manager_name: "Менеджер",
  discount_percent: 0,
  subtotal: 1000,
  vat_amount: 220,
  total_amount: 1220,
  execution_terms: null,
  status: "в архиве",
  completion_percentage: null,
};

describe("sectionFromStatus", () => {
  it("maps На СГП to in_production", () => {
    expect(sectionFromStatus({ ...baseItem, status: "На СГП" })).toBe("in_production");
  });

  it("maps в работе to in_production", () => {
    expect(sectionFromStatus({ ...baseItem, status: "в работе" })).toBe("in_production");
  });

  it("maps выполнено to completed", () => {
    expect(sectionFromStatus({ ...baseItem, status: "выполнено" })).toBe("completed");
  });

  it("maps в архиве to archived", () => {
    expect(sectionFromStatus({ ...baseItem, status: "в архиве" })).toBe("archived");
  });
});

describe("filterByProductType MNA-602 — contains-type", () => {
  const mixedWithPlates: ArchiveOfferListItem = {
    ...baseItem,
    kp_id: 10,
    product_type: "mixed",
    product_types: ["plates", "piles"],
  };

  const monoPiles: ArchiveOfferListItem = {
    ...baseItem,
    kp_id: 11,
    product_type: "piles",
  };

  const monoPlates: ArchiveOfferListItem = {
    ...baseItem,
    kp_id: 12,
    product_type: "plates",
  };

  it("keeps mixed KP when product_types contains the filter", () => {
    const result = filterByProductType(
      [mixedWithPlates, monoPiles, monoPlates],
      "plates",
    );
    expect(result.map((item) => item.kp_id)).toEqual([10, 12]);
  });

  it("matches piles filter via product_types on mixed KP", () => {
    const result = filterByProductType([mixedWithPlates, monoPlates], "piles");
    expect(result.map((item) => item.kp_id)).toEqual([10]);
  });

  it("ignores literal mixed token in product_types", () => {
    expect(
      concreteProductTypes({
        ...baseItem,
        product_type: "mixed",
        product_types: ["mixed", "steps"],
      }),
    ).toEqual(["steps"]);
  });

  it("returns all items when filter is all", () => {
    const items = [mixedWithPlates, monoPiles];
    expect(filterByProductType(items, "all")).toEqual(items);
  });
});
