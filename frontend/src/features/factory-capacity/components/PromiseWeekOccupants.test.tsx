import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render, screen } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { PromiseWeekOccupants } from "@/features/factory-capacity/components/PromiseWeekOccupants";
import type { PromiseWeekOccupants as OccupantsPayload } from "@/features/factory-capacity/api/promiseQuote";

const occupantsState = {
  data: undefined as OccupantsPayload | undefined,
  isPending: false,
  isError: false,
  error: null as Error | null,
};

vi.mock("@/features/factory-capacity/api/promiseQuote", async () => {
  const actual = await vi.importActual<typeof import("@/features/factory-capacity/api/promiseQuote")>(
    "@/features/factory-capacity/api/promiseQuote",
  );
  return {
    ...actual,
    usePromiseWeekOccupantsQuery: () => occupantsState,
  };
});

afterEach(() => {
  cleanup();
  occupantsState.data = undefined;
  occupantsState.isPending = false;
  occupantsState.isError = false;
  occupantsState.error = null;
});

const wrap = (ui: ReactNode) => {
  const client = new QueryClient({ defaultOptions: { queries: { retry: false } } });
  return render(<QueryClientProvider client={client}>{ui}</QueryClientProvider>);
};

describe("PromiseWeekOccupants", () => {
  it("lists named holds and promises with planned count", () => {
    occupantsState.data = {
      week_start: "2026-09-07",
      planned: 4,
      occupants: [
        {
          kp_id: 7,
          customer_name: "АО Чужой",
          kind: "promise",
          tracks: 6,
          promised_date: "2026-09-18",
          is_current: false,
        },
        {
          kp_id: 42,
          customer_name: "ООО Тест",
          kind: "hold",
          tracks: 3,
          promised_date: "2026-09-18",
          is_current: true,
        },
      ],
    };

    wrap(
      <PromiseWeekOccupants
        kpId={42}
        weekStart="2026-09-07"
        weekFree={8}
        weekCapacity={15}
        selectedDay="2026-09-09"
        occupancy={{ "2026-09-09": 1 }}
        knob={3}
      />,
    );

    const list = screen.getByTestId("promise-week-occupants");
    expect(list).toHaveTextContent("КП №7");
    expect(list).toHaveTextContent("АО Чужой");
    expect(list).toHaveTextContent("обещано");
    expect(list).toHaveTextContent("6");
    expect(list).toHaveTextContent("КП №42");
    expect(list).toHaveTextContent("ООО Тест");
    expect(list).toHaveTextContent("холд");
    expect(list).toHaveTextContent("это КП");
    expect(list).toHaveTextContent("Уже в плане: 4 дорожек");
    expect(list).toHaveTextContent("Свободно: 8 из 15");
    expect(list).toHaveTextContent("9.09: 1/3, остаток 2");
    expect(list).toHaveTextContent("Холды не занимают свободно — до перевода место могут взять другие.");
    expect(list).not.toHaveTextContent("created_by");
    expect(list).not.toHaveTextContent("Закрепил");
  });

  it("shows empty copy when the journal is empty", () => {
    occupantsState.data = {
      week_start: "2026-09-07",
      planned: 0,
      occupants: [],
    };

    wrap(<PromiseWeekOccupants kpId={42} weekStart="2026-09-07" />);

    expect(screen.getByTestId("promise-week-occupants")).toHaveTextContent(
      "На этой неделе нет холдов и обещаний",
    );
    expect(screen.getByTestId("promise-week-occupants")).toHaveTextContent("Уже в плане: 0 дорожек");
  });

  it("formats overflow and empty clicked days", () => {
    occupantsState.data = {
      week_start: "2026-09-07",
      planned: 4,
      occupants: [],
    };

    wrap(
      <PromiseWeekOccupants
        kpId={42}
        weekStart="2026-09-07"
        weekFree={0}
        weekCapacity={15}
        selectedDay="2026-09-10"
        occupancy={{ "2026-09-10": 4 }}
        knob={3}
      />,
    );
    expect(screen.getByTestId("promise-week-day-line")).toHaveTextContent("10.09: 4/3 · перебор");

    cleanup();
    wrap(
      <PromiseWeekOccupants
        kpId={42}
        weekStart="2026-09-07"
        selectedDay="2026-09-11"
        occupancy={{}}
        knob={3}
      />,
    );
    expect(screen.getByTestId("promise-week-day-line")).toHaveTextContent("11.09: свободно 3");

    cleanup();
    wrap(
      <PromiseWeekOccupants
        kpId={42}
        weekStart="2026-09-07"
        selectedDay="2026-09-12"
        knob={3}
      />,
    );
    expect(screen.getByTestId("promise-week-day-line")).toHaveTextContent("12.09: нерабочий");
  });

  it("shows an alert when the occupants request fails", () => {
    occupantsState.isError = true;
    occupantsState.error = new Error("week_start должен быть понедельником.");

    wrap(<PromiseWeekOccupants kpId={42} weekStart="2026-09-02" />);

    expect(screen.getByText(/понедельник/)).toBeInTheDocument();
  });
});
