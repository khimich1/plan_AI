import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { DeliveryScheduleDialog } from "@/features/delivery-schedule/components/DeliveryScheduleDialog";

afterEach(() => {
  cleanup();
});

vi.mock("@/features/delivery-schedule/hooks/useDeliveryScheduleQueries", () => ({
  useDeliveryScheduleQuery: () => ({
    isPending: false,
    isError: false,
    data: {
      kp_id: 1,
      invoice_number: null,
      contract_number: null,
      status: "draft",
      updated_at: "2026-03-01T00:00:00",
      traffic_light_degraded: false,
      batches: [
        {
          id: 1,
          sort_order: 0,
          name: "П1",
          deliver_from: "2026-04-01",
          deliver_to: "2026-04-10",
          produce_by: "2026-03-20",
          items: [{ plate_id: 1, plate_name: "ПБ", qty: 2, changed: false }],
          status: "red",
          ready_date: "2026-03-25",
          hint: "нужно +2 дорожек до 20.03.2026",
          changed: false,
        },
      ],
    },
    error: null,
  }),
  usePutDeliveryScheduleMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
  useDownloadDeliveryScheduleTemplateMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
  }),
  useDownloadDeliveryScheduleDocumentMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
  }),
  useImportDeliveryScheduleMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
  }),
}));

vi.mock("@/features/factory-capacity/hooks/useCapacitySnapshotQuery", () => ({
  useCapacitySnapshotQuery: () => ({
    data: {
      start_date: "2026-03-03",
      target_date: "2026-03-20",
      tracks_needed: 5,
      tracks_free_in_window: 1,
      delta: -4,
      status: "red",
      hint: "нужно +4 дорожек до 20.03.2026",
      days_info: {},
      holidays: [],
      extra_workdays: [],
      calendar_from_month: "2026-03",
      calendar_to_month: "2026-03",
    },
    isFetching: false,
    isError: false,
    error: null,
  }),
}));

const wrap = (ui: ReactNode) => {
  const client = new QueryClient({ defaultOptions: { queries: { retry: false } } });
  return render(<QueryClientProvider client={client}>{ui}</QueryClientProvider>);
};

describe("DeliveryScheduleDialog capacity gate", () => {
  it("keeps hint in modal, calendar only after Ёмкость click", () => {
    wrap(
      <DeliveryScheduleDialog
        open
        onClose={() => undefined}
        kpId={1}
        plates={[{ id: 1, plate_name: "ПБ", qty: 10 }]}
      />,
    );

    expect(screen.getByText(/нужно \+4 дорожек/)).toBeInTheDocument();
    expect(screen.queryByTestId("factory-capacity-panel")).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Сохранить/i })).toBeDisabled();

    fireEvent.click(screen.getByRole("button", { name: /^Ёмкость$/i }));
    expect(screen.getByTestId("factory-capacity-panel")).toBeInTheDocument();
    expect(screen.getByTestId("factory-mini-calendar")).toBeInTheDocument();
  });
});
