import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { GsmRegistriesView } from "@/features/gsm/components/GsmRegistriesView";
import { ApiError } from "@/shared/lib/apiError";
import type { GsmSettings } from "@/features/gsm/types/gsm";

const mockSeasonMutate = vi.fn();
const mockResetMutate = vi.fn();

const SUMMER_SETTINGS: GsmSettings = {
  winter_start: "11-01",
  hook_threshold_km: 13,
  season_mode: "summer",
  season_switched_at: null,
};

let settingsData: GsmSettings = SUMMER_SETTINGS;
let mutationState = { isPending: false, isError: false, error: null as unknown };
let resetMutationState = { isPending: false, isError: false, error: null as unknown };
let authRole: string | undefined = "accountant";

vi.mock("@/features/auth/model/AuthProvider", () => ({
  useAuth: () => ({ user: authRole ? { role: authRole } : null }),
}));

vi.mock("@/features/gsm/components/CardsRegistryView", () => ({
  CardsRegistryView: () => null,
}));
vi.mock("@/features/gsm/components/DriversRegistryView", () => ({
  DriversRegistryView: () => null,
}));
vi.mock("@/features/gsm/components/VehiclesCard", () => ({
  VehiclesCard: () => null,
}));

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useGsmStationsQuery: () => ({ isLoading: false, error: null, data: [] }),
  useGsmSettingsQuery: () => ({ isLoading: false, error: null, data: settingsData }),
  useUpdateGsmSeasonMutation: () => ({
    mutate: mockSeasonMutate,
    isPending: mutationState.isPending,
    isError: mutationState.isError,
    error: mutationState.error,
  }),
  useGsmResetToAnchorsMutation: () => ({
    mutate: mockResetMutate,
    isPending: resetMutationState.isPending,
    isError: resetMutationState.isError,
    error: resetMutationState.error,
  }),
}));

describe("GsmRegistriesView season switch", () => {
  beforeEach(() => {
    settingsData = SUMMER_SETTINGS;
    mutationState = { isPending: false, isError: false, error: null };
    resetMutationState = { isPending: false, isError: false, error: null };
    authRole = "accountant";
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
    vi.useRealTimers();
  });

  it("shows летний режим and offers switch to winter", () => {
    render(<GsmRegistriesView />);

    expect(screen.getByText(/Режим: летний/)).toBeInTheDocument();
    expect(screen.getByText(/порог крюка 13 км/)).toBeInTheDocument();
    expect(
      screen.getByRole("button", { name: "Перевести на зимний режим" }),
    ).toBeInTheDocument();
  });

  it("shows зимний режим with switch date and offers switch to summer", () => {
    settingsData = {
      ...SUMMER_SETTINGS,
      season_mode: "winter",
      season_switched_at: "2026-11-01",
    };
    render(<GsmRegistriesView />);

    expect(screen.getByText(/Режим: зимний \(с 2026-11-01\)/)).toBeInTheDocument();
    expect(
      screen.getByRole("button", { name: "Перевести на летний режим" }),
    ).toBeInTheDocument();
  });

  it("click switches season with today's ISO date", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-08-25T12:00:00"));

    render(<GsmRegistriesView />);
    fireEvent.click(screen.getByRole("button", { name: "Перевести на зимний режим" }));

    expect(mockSeasonMutate).toHaveBeenCalledWith({ mode: "winter", date: "2026-08-25" });
  });

  it("shows backend error text when season switch fails", () => {
    mutationState = {
      isPending: false,
      isError: true,
      error: new ApiError("season", 422, "season date must not be before last switch"),
    };
    render(<GsmRegistriesView />);

    expect(
      screen.getByText("Дата перевода сезона не может быть раньше предыдущего перевода."),
    ).toBeInTheDocument();
  });
});

describe("GsmRegistriesView admin reset", () => {
  beforeEach(() => {
    settingsData = SUMMER_SETTINGS;
    mutationState = { isPending: false, isError: false, error: null };
    resetMutationState = { isPending: false, isError: false, error: null };
    authRole = "admin";
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("shows reset button for admin", () => {
    render(<GsmRegistriesView />);

    expect(screen.getByRole("button", { name: "Сброс к якорям" })).toBeInTheDocument();
    expect(screen.getByText(/Dev-инструменты/)).toBeInTheDocument();
  });

  it("hides reset button for accountant", () => {
    authRole = "accountant";
    render(<GsmRegistriesView />);

    expect(screen.queryByRole("button", { name: "Сброс к якорям" })).not.toBeInTheDocument();
  });
});
