import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { MoveToProductionDialog } from "@/features/commercial-archive/components/MoveToProductionDialog";

afterEach(() => {
  cleanup();
});

vi.mock("@/features/commercial-archive/hooks/useArchiveQueries", () => ({
  useMoveToProductionMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
  useProductionEstimateQuery: () => ({
    isPending: false,
    data: { estimated_tracks: 2, estimated_days: 3, total_length_m: 100 },
  }),
}));

vi.mock("@/features/factory-capacity/hooks/useCapacitySnapshotQuery", () => ({
  useCapacitySnapshotQuery: () => ({
    data: {
      start_date: "2026-03-03",
      target_date: "2026-03-20",
      tracks_needed: 10,
      tracks_free_in_window: 2,
      delta: -8,
      status: "red",
      hint: "нужно +8 дорожек до 20.03.2026",
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

describe("MoveToProductionDialog capacity gate", () => {
  it("keeps manufacturing terms field for archive → production", () => {
    wrap(
      <MoveToProductionDialog
        open
        onClose={() => undefined}
        kpId={42}
        initialExecutionTerms="20.03.2026"
      />,
    );

    expect(screen.getByText("Срок выполнения")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Перевести в производство/i })).toBeInTheDocument();
  });

  it("keeps hint in modal, calendar only after Ёмкость click", () => {
    wrap(
      <MoveToProductionDialog
        open
        onClose={() => undefined}
        kpId={42}
        initialExecutionTerms="20.03.2026"
      />,
    );

    expect(screen.getByText(/нужно \+8 дорожек/)).toBeInTheDocument();
    expect(screen.queryByTestId("factory-capacity-panel")).not.toBeInTheDocument();
    const submit = screen.getByRole("button", { name: /Перевести в производство/i });
    expect(submit).toBeDisabled();

    fireEvent.click(screen.getByRole("button", { name: /^Ёмкость$/i }));
    expect(screen.getByTestId("factory-capacity-panel")).toBeInTheDocument();
    expect(screen.getByTestId("factory-mini-calendar")).toBeInTheDocument();
  });

  it("Esc closes capacity drawer without closing modal", () => {
    const onClose = vi.fn();
    wrap(
      <MoveToProductionDialog
        open
        onClose={onClose}
        kpId={42}
        initialExecutionTerms="20.03.2026"
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: /^Ёмкость$/i }));
    expect(screen.getByTestId("factory-capacity-panel")).toBeInTheDocument();

    fireEvent.keyDown(window, { key: "Escape" });

    expect(screen.queryByTestId("factory-capacity-panel")).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Перевести в производство/i })).toBeInTheDocument();
    expect(onClose).not.toHaveBeenCalled();
  });
});
