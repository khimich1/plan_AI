import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ArchiveOfferList } from "@/features/commercial-archive/components/ArchiveOfferList";
import type { ArchiveOfferListItem } from "@/features/commercial-archive/types/archive";

const baseItem = {
  kp_id: 42,
  creation_date: "01.03.2026",
  customer_name: "ООО Тест",
  manager_name: "Иван Иванов",
  discount_percent: 5,
  subtotal: 1000,
  vat_amount: 220,
  total_amount: 1220,
  execution_terms: null,
  status: "в архиве",
  completion_percentage: null,
} satisfies ArchiveOfferListItem;

afterEach(() => {
  cleanup();
});

describe("ArchiveOfferList MNA-602 — multi badges", () => {
  it("renders N badges for product_types present", () => {
    const item = {
      ...baseItem,
      product_type: "plates",
      product_types: ["plates", "piles"],
    } as ArchiveOfferListItem;

    render(
      <ArchiveOfferList section="archived" items={[item]} onSelect={vi.fn()} />,
    );

    expect(screen.getByText("Плиты")).toBeInTheDocument();
    expect(screen.getByText("Сваи")).toBeInTheDocument();
  });

  it("renders three badges when three types are present", () => {
    const item = {
      ...baseItem,
      product_types: ["plates", "piles", "steps"],
    } as ArchiveOfferListItem;

    render(
      <ArchiveOfferList section="archived" items={[item]} onSelect={vi.fn()} />,
    );

    expect(screen.getByText("Плиты")).toBeInTheDocument();
    expect(screen.getByText("Сваи")).toBeInTheDocument();
    expect(screen.getByText("Ступени")).toBeInTheDocument();
  });

  it("does not render a single mixed badge when product_types has concrete types", () => {
    const item = {
      ...baseItem,
      product_type: "plates",
      product_types: ["plates", "piles"],
    } as ArchiveOfferListItem;

    render(
      <ArchiveOfferList section="archived" items={[item]} onSelect={vi.fn()} />,
    );

    expect(screen.queryByText(/mixed/i)).not.toBeInTheDocument();
  });
});
