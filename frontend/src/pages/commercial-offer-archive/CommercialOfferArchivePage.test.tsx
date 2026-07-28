import { describe, expect, it } from "vitest";
import { sectionFromStatus } from "@/pages/commercial-offer-archive/CommercialOfferArchivePage";
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
