import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { MoveToProductionDialog } from "@/features/commercial-archive/components/MoveToProductionDialog";
import type { PromiseHold, PromiseQuote } from "@/features/factory-capacity/api/promiseQuote";

afterEach(() => {
  cleanup();
  quoteState.isPending = false;
  quoteState.isError = false;
  quoteState.error = null;
  quoteState.data = sampleQuote;
  holdState.data = null;
  holdState.isPending = false;
  createHoldMutate.mockReset();
  moveMutate.mockReset();
});

const sampleQuote: PromiseQuote = {
  tracks: 2,
  solo_days: 1,
  solo_date: "2026-09-04",
  solo_week_end_date: "2026-09-06",
  earliest_start_week: "2026-08-31",
  first_pour_date: "2026-09-03",
  first_pour_free: 3,
  window: {
    from_week: "2026-08-31",
    to_week: "2026-08-31",
    promised_date: "2026-09-04",
  },
  weeks: [
    {
      week_start: "2026-08-31",
      workdays: 2,
      capacity: 6,
      planned: 0,
      promised: 0,
      held: 1,
      free: 6,
    },
  ],
  knob: 3,
  occupancy: { "2026-09-03": 0 },
};

const quoteState = {
  isPending: false,
  isError: false,
  error: null as Error | null,
  data: sampleQuote as PromiseQuote | undefined,
};

const sampleHold: PromiseHold = {
  id: 7,
  kp_id: 42,
  kind: "hold",
  status: "active",
  tracks_total: 2,
  promised_date: "2026-09-04",
  expires_at: "2026-09-03T23:59:59",
  created_by: "alice",
  created_at: "2026-09-03T12:00:00",
  allocations: [{ week_start: "2026-08-31", tracks: 2 }],
};

const holdState = {
  data: null as PromiseHold | null,
  isPending: false,
};

const createHoldMutate = vi.fn();
const moveMutate = vi.fn();

vi.mock("@/features/commercial-archive/hooks/useArchiveQueries", () => ({
  useMoveToProductionMutation: () => ({
    mutateAsync: moveMutate,
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
}));

vi.mock("@/features/factory-capacity/api/promiseQuote", async () => {
  const actual = await vi.importActual<typeof import("@/features/factory-capacity/api/promiseQuote")>(
    "@/features/factory-capacity/api/promiseQuote",
  );
  return {
    ...actual,
    usePromiseQuoteQuery: () => quoteState,
    usePromiseHoldQuery: () => holdState,
    useCreatePromiseHoldMutation: () => ({
      mutateAsync: createHoldMutate,
      isPending: false,
      isError: false,
      error: null,
      reset: vi.fn(),
    }),
    usePromiseWeekOccupantsQuery: () => ({
      isPending: false,
      isError: false,
      error: null,
      data: { week_start: "2026-08-31", planned: 0, occupants: [] },
    }),
  };
});

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

  it("shows quote and window band instead of production estimate", () => {
    wrap(
      <MoveToProductionDialog
        open
        onClose={() => undefined}
        kpId={42}
        initialExecutionTerms="20.03.2026"
      />,
    );

    const block = screen.getByTestId("promise-quote-block");
    expect(block).toHaveTextContent("~2 дорожек");
    expect(block).toHaveTextContent(/Обещать к 4\.09/);
    expect(block).toHaveTextContent(/Начало:\s*3\.09\s*·\s*остаток 3 дор\./);
    expect(block).toHaveTextContent(/Если только его:\s*4\.09/);
    expect(block).toHaveTextContent(/Соло \+ до конца недели:\s*6\.09/);
    expect(screen.getByTestId("promise-window-band")).toBeInTheDocument();
    expect(screen.queryByTestId("promise-week-strip")).not.toBeInTheDocument();
    expect(screen.queryByText(/Оценка производства/)).not.toBeInTheDocument();
  });

  it("prefills execution terms from promised_date when the card has none", () => {
    wrap(<MoveToProductionDialog open onClose={() => undefined} kpId={42} />);
    expect(screen.getByRole("textbox")).toHaveValue("04.09.2026");
  });

  it("does not block submit on capacity-snapshot; Ёмкость shows period calendar", () => {
    wrap(
      <MoveToProductionDialog
        open
        onClose={() => undefined}
        kpId={42}
        initialExecutionTerms="20.03.2026"
      />,
    );

    expect(screen.queryByText(/нужно \+8 дорожек/)).not.toBeInTheDocument();
    expect(screen.queryByTestId("factory-capacity-panel")).not.toBeInTheDocument();
    expect(screen.queryByTestId("factory-mini-calendar")).not.toBeInTheDocument();
    const submit = screen.getByRole("button", { name: /Перевести в производство/i });
    expect(submit).toBeEnabled();

    fireEvent.click(screen.getByRole("button", { name: /^Ёмкость$/i }));
    expect(screen.getByText("Ёмкость завода")).toBeInTheDocument();
    expect(screen.getByTestId("promise-period-calendar")).toBeInTheDocument();
    expect(screen.getByTestId("promise-week-occupants")).toBeInTheDocument();
    expect(screen.getByTestId("promise-knob-settings")).toBeInTheDocument();
    expect(screen.queryByTestId("promise-week-strip")).not.toBeInTheDocument();
    expect(screen.queryByTestId("factory-mini-calendar")).not.toBeInTheDocument();
    expect(screen.queryByTestId("factory-capacity-panel")).not.toBeInTheDocument();
  });

  it("shows occupancy error instead of an empty quote", () => {
    quoteState.isError = true;
    quoteState.error = new Error("Недоступна занятость плана — котировка остановлена (fail-closed).");
    quoteState.data = undefined;

    wrap(
      <MoveToProductionDialog
        open
        onClose={() => undefined}
        kpId={42}
        initialExecutionTerms="20.03.2026"
      />,
    );

    expect(screen.getByText(/Недоступна занятость плана/)).toBeInTheDocument();
    expect(screen.queryByTestId("promise-quote-block")).not.toBeInTheDocument();
    expect(screen.queryByTestId("promise-week-strip")).not.toBeInTheDocument();
    expect(screen.queryByTestId("promise-window-band")).not.toBeInTheDocument();
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
    expect(screen.getByText("Ёмкость завода")).toBeInTheDocument();

    fireEvent.keyDown(window, { key: "Escape" });

    expect(screen.queryByText("Ёмкость завода")).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Перевести в производство/i })).toBeInTheDocument();
    expect(onClose).not.toHaveBeenCalled();
  });
});

describe("MoveToProductionDialog promise hold", () => {
  it("shows Закрепить срок when quote is loaded", () => {
    wrap(<MoveToProductionDialog open onClose={() => undefined} kpId={42} />);

    expect(screen.getByRole("button", { name: /Закрепить срок/i })).toBeEnabled();
    expect(screen.queryByTestId("promise-hold-locked")).not.toBeInTheDocument();
  });

  it("creates a hold via Закрепить срок", async () => {
    createHoldMutate.mockResolvedValue(sampleHold);

    wrap(<MoveToProductionDialog open onClose={() => undefined} kpId={42} />);
    fireEvent.click(screen.getByRole("button", { name: /Закрепить срок/i }));

    expect(createHoldMutate).toHaveBeenCalledWith(42);
  });

  it("shows срок закреплён до сегодня when hold is active", () => {
    holdState.data = sampleHold;

    wrap(<MoveToProductionDialog open onClose={() => undefined} kpId={42} />);

    expect(screen.getByTestId("promise-hold-locked")).toHaveTextContent("Срок закреплён до сегодня");
    expect(screen.getByTestId("promise-hold-locked")).toHaveAttribute("title", "Закрепил: alice");
    expect(screen.getByRole("button", { name: /Закрепить срок/i })).toBeDisabled();
  });

  it("moves to production from an active hold without re-entering the date", async () => {
    holdState.data = sampleHold;
    moveMutate.mockResolvedValue({ kp_id: 42 });

    wrap(<MoveToProductionDialog open onClose={() => undefined} kpId={42} />);

    expect(screen.getByRole("textbox")).toHaveValue("04.09.2026");
    const submit = screen.getByRole("button", { name: /Перевести в производство/i });
    expect(submit).toBeEnabled();
    fireEvent.click(submit);

    expect(moveMutate).toHaveBeenCalledWith({ kpId: 42, executionTerms: "04.09.2026" });
  });

  it("does not show week-strip cards in the dialog after the window band", () => {
    wrap(<MoveToProductionDialog open onClose={() => undefined} kpId={42} />);
    expect(screen.getByTestId("promise-window-band")).toBeInTheDocument();
    expect(screen.queryByTestId("promise-week-strip")).not.toBeInTheDocument();
  });

  it("clicking a calendar day selects the week and does not change execution terms", () => {
    wrap(
      <MoveToProductionDialog
        open
        onClose={() => undefined}
        kpId={42}
        initialExecutionTerms="20.03.2026"
      />,
    );

    const field = screen.getByRole("textbox");
    expect(field).toHaveValue("20.03.2026");
    fireEvent.click(screen.getByRole("button", { name: /^Ёмкость$/i }));
    fireEvent.click(screen.getByTestId("promise-cal-day-2026-09-04"));
    expect(field).toHaveValue("20.03.2026");
    expect(screen.getByTestId("promise-week-day-line")).toHaveTextContent("4.09: свободно 3");
    expect(screen.getByTestId("promise-week-free")).toHaveTextContent("Свободно: 6 из 6");
  });

  it("opens the later field month in Ёмкость even past quote weeks", () => {
    wrap(
      <MoveToProductionDialog
        open
        onClose={() => undefined}
        kpId={42}
        initialExecutionTerms="20.03.2027"
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: /^Ёмкость$/i }));
    expect(screen.getByTestId("promise-period-calendar")).toHaveTextContent("март 2027");
  });
});
