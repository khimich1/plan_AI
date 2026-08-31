import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { TransactionsJournalView } from "@/features/gsm/components/TransactionsJournalView";

const mockTxQuery = vi.fn();

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useGsmTransactionsQuery: (params: unknown) => {
    mockTxQuery(params);
    return {
      isLoading: false,
      error: null,
      data: {
        rows: [
          {
            ts: "2026-08-03T10:00:00",
            card_number: "111",
            vehicle_id: 1,
            service_type: "fuel",
            fuel_grade: "АИ-95",
            qty_liters: 10,
            amount: 100,
            station_id: 1,
            address: "АЗС 1",
          },
          {
            ts: "2026-08-04T10:00:00",
            card_number: "orphan",
            vehicle_id: null,
            service_type: "fuel",
            fuel_grade: "АИ-95",
            qty_liters: 50,
            amount: 3000,
            station_id: null,
            address: "АЗС 2",
          },
        ],
        total_count: 2,
        sum_liters: 99,
        sum_amount: 12.5,
      },
    };
  },
  useGsmVehiclesQuery: () => ({
    isLoading: false,
    error: null,
    data: [
      {
        id: 1,
        name: "Palisade",
        plate_number: "О 521",
        tank_volume_liters: 70,
        norm_summer: 10,
        norm_winter: 11,
        primary_driver_id: 1,
        is_active: true,
      },
    ],
  }),
}));

describe("TransactionsJournalView", () => {
  afterEach(() => {
    cleanup();
    mockTxQuery.mockClear();
  });

  it("renders backend totals and highlights an unbound card", () => {
    render(<TransactionsJournalView />);
    expect(screen.getByTestId("tx-sum-liters")).toHaveTextContent("99");
    expect(screen.getByTestId("tx-sum-amount")).toHaveTextContent("12,50");
    expect(screen.getByTestId("unbound-card-row")).toHaveTextContent(/Не привязана/);
  });

  it("updates query params when filters change", () => {
    render(<TransactionsJournalView />);
    fireEvent.change(screen.getByLabelText("Фильтр машины"), { target: { value: "1" } });
    fireEvent.change(screen.getByLabelText("Фильтр типа"), { target: { value: "fuel" } });
    const last = mockTxQuery.mock.calls.at(-1)?.[0] as {
      vehicleId?: number;
      serviceType?: string;
    };
    expect(last.vehicleId).toBe(1);
    expect(last.serviceType).toBe("fuel");
  });
});
