import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { afterEach, describe, expect, it, vi } from "vitest";
import { GsmPage } from "@/pages/gsm/GsmPage";

const { noopMutation } = vi.hoisted(() => ({
  noopMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
}));

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useGsmVehiclesQuery: () => ({
    isLoading: false,
    error: null,
    data: [
      {
        id: 1,
        name: "Geely Monjaro",
        plate_number: "A123BC77",
        tank_volume_liters: 60,
        norm_summer: 12.5,
        norm_winter: 14.1,
        primary_driver_id: null,
        is_active: true,
      },
    ],
  }),
  useGsmDriversQuery: () => ({
    isLoading: false,
    error: null,
    data: [
      {
        id: 1,
        full_name: "Иванов Иван",
        license_number: "1234567",
        license_issued_at: null,
        personnel_number: "42",
        snils: null,
        is_active: true,
      },
    ],
  }),
  useGsmCardsQuery: () => ({
    isLoading: false,
    error: null,
    data: [
      {
        id: 1,
        card_number: "7001",
        vehicle_id: 1,
        assigned_at: "2026-01-01",
        archived_at: null,
      },
    ],
  }),
  useGsmStationsQuery: () => ({
    isLoading: false,
    error: null,
    data: [{ id: 1, address: "МКАД 42 км", brand: "Лукойл", lat: null, lon: null, geocode_source: null }],
  }),
  useGsmSettingsQuery: () => ({
    isLoading: false,
    error: null,
    data: { winter_start: "11-01", hook_threshold_km: 13 },
  }),
  useGsmWaybillsQuery: () => ({
    isLoading: false,
    error: null,
    data: [],
    refetch: vi.fn(),
  }),
  useCreateVehicleMutation: noopMutation,
  usePatchVehicleMutation: noopMutation,
  useCreateDriverMutation: noopMutation,
  usePatchDriverMutation: noopMutation,
  useCreateCardMutation: noopMutation,
  usePatchCardMutation: noopMutation,
  useCreateStationMutation: noopMutation,
  usePatchStationMutation: noopMutation,
  usePutGsmSettingsMutation: noopMutation,
  useImportGsmTransactionsMutation: noopMutation,
  useGenerateGsmWaybillsMutation: noopMutation,
  useExportGsmWaybillsMutation: noopMutation,
  useGsmRoutesQuery: () => ({ isLoading: false, error: null, data: [] }),
  usePatchGsmWaybillMutation: noopMutation,
  useCreateGsmWaybillMutation: noopMutation,
}));

function renderGsmPage() {
  const qc = new QueryClient({
    defaultOptions: { queries: { retry: false } },
  });
  return render(
    <QueryClientProvider client={qc}>
      <GsmPage />
    </QueryClientProvider>,
  );
}

describe("GsmPage", () => {
  afterEach(() => {
    cleanup();
  });

  it("renders title and three tabs", () => {
    renderGsmPage();
    expect(screen.getByRole("heading", { name: "ГСМ" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Период" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Транзакции" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Справочники" })).toBeInTheDocument();
  });

  it("shows period review screen by default and switches to registries with loaded data", () => {
    renderGsmPage();

    expect(screen.getByRole("heading", { name: "Период × машина" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Сгенерировать" })).toBeInTheDocument();
    expect(screen.getByText(/Выберите машину и период/i)).toBeInTheDocument();

    fireEvent.click(screen.getByRole("tab", { name: "Справочники" }));

    expect(screen.getByText("Geely Monjaro")).toBeInTheDocument();
    expect(screen.getByText("Иванов Иван")).toBeInTheDocument();
    expect(screen.getByText("7001")).toBeInTheDocument();
    expect(screen.getByText("МКАД 42 км")).toBeInTheDocument();
    expect(screen.getByText(/порог крюка 13 км/i)).toBeInTheDocument();
  });

  it("opens import dialog from transactions tab", () => {
    renderGsmPage();
    fireEvent.click(screen.getByRole("tab", { name: "Транзакции" }));
    expect(screen.getByRole("button", { name: "Импорт транзакций" })).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Импорт транзакций" }));
    expect(screen.getByRole("dialog")).toBeInTheDocument();
    expect(screen.getByText(/Перетащите \.xls сюда/i)).toBeInTheDocument();
  });
});
