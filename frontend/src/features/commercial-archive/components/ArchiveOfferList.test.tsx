import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ArchiveOfferList } from "@/features/commercial-archive/components/ArchiveOfferList";
import type { ArchiveOfferListItem } from "@/features/commercial-archive/types/archive";
import type { PromiseHold } from "@/features/factory-capacity/api/promiseQuote";

const holdsByKp = new Map<number, PromiseHold>();

vi.mock("@/features/factory-capacity/api/promiseQuote", async () => {
  const actual = await vi.importActual<typeof import("@/features/factory-capacity/api/promiseQuote")>(
    "@/features/factory-capacity/api/promiseQuote",
  );
  return {
    ...actual,
    usePromiseHoldsMap: (kpIds: number[]) => {
      const map = new Map<number, PromiseHold>();
      for (const id of kpIds) {
        const hold = holdsByKp.get(id);
        if (hold) {
          map.set(id, hold);
        }
      }
      return map;
    },
  };
});

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
  holdsByKp.clear();
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

describe("ArchiveOfferList promise hold badge", () => {
  it("shows срок закреплён до сегодня with who pinned it", () => {
    holdsByKp.set(42, {
      id: 1,
      kp_id: 42,
      kind: "hold",
      status: "active",
      tracks_total: 2,
      promised_date: "2026-09-04",
      expires_at: "2026-09-03T23:59:59",
      created_by: "alice",
      created_at: "2026-09-03T12:00:00",
      allocations: [],
    });

    render(
      <ArchiveOfferList section="archived" items={[baseItem]} onSelect={vi.fn()} />,
    );

    const badge = screen.getByTestId("promise-hold-badge");
    expect(badge).toHaveTextContent("срок закреплён до сегодня");
    expect(badge).toHaveAttribute("title", "Закрепил: alice");
  });

  it("hides the hold badge when there is no active hold", () => {
    render(
      <ArchiveOfferList section="archived" items={[baseItem]} onSelect={vi.fn()} />,
    );
    expect(screen.queryByTestId("promise-hold-badge")).not.toBeInTheDocument();
  });
});
