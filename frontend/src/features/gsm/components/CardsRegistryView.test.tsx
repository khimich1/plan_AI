import { cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { CardsRegistryView } from "@/features/gsm/components/CardsRegistryView";
import { ApiError } from "@/shared/lib/apiError";
import type { GsmCard, GsmVehicle } from "@/features/gsm/types/gsm";

const mockPatch = vi.fn();
const mockUseCards = vi.fn();
const mockUseVehicles = vi.fn();

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useGsmCardsQuery: (...args: unknown[]) => mockUseCards(...args),
  useGsmVehiclesQuery: (...args: unknown[]) => mockUseVehicles(...args),
  usePatchCardMutation: () => ({ mutateAsync: mockPatch, isPending: false }),
}));

const VEHICLES: GsmVehicle[] = [
  {
    id: 1,
    name: "Geely Monjaro",
    plate_number: "A123BC77",
    tank_volume_liters: 60,
    norm_summer: 12,
    norm_winter: 14,
    primary_driver_id: null,
    is_active: true,
  },
  {
    id: 2,
    name: "Geely Tugella",
    plate_number: "B456CD77",
    tank_volume_liters: 55,
    norm_summer: 9.4,
    norm_winter: 10.3,
    primary_driver_id: null,
    is_active: true,
  },
];

const CARDS: GsmCard[] = [
  {
    id: 10,
    card_number: "7001",
    vehicle_id: 1,
    assigned_at: "2026-01-01",
    archived_at: null,
  },
  {
    id: 11,
    card_number: "7002",
    vehicle_id: null,
    assigned_at: "2026-02-01",
    archived_at: null,
  },
];

describe("CardsRegistryView", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  const renderView = () => {
    mockUseCards.mockReturnValue({ data: CARDS, isLoading: false, error: null });
    mockUseVehicles.mockReturnValue({ data: VEHICLES, isLoading: false, error: null });
    mockPatch.mockResolvedValue({ ...CARDS[1], vehicle_id: 2 });
    return render(<CardsRegistryView />);
  };

  it("binds card to vehicle without page reload (mutation + success message)", async () => {
    renderView();

    const select = screen.getByLabelText("Машина для карты 7002") as HTMLSelectElement;
    fireEvent.change(select, { target: { value: "2" } });

    await waitFor(() => {
      expect(mockPatch).toHaveBeenCalledWith({ id: 11, payload: { vehicle_id: 2 } });
    });
    expect(await screen.findByText(/Карта 7002 привязана к машине/)).toBeInTheDocument();
  });

  it("archives card via mutation and shows confirmation", async () => {
    mockPatch.mockResolvedValue({ ...CARDS[0], archived_at: "2026-08-14T12:00:00" });
    renderView();

    const row = screen.getByText("7001").closest("tr") as HTMLElement;
    fireEvent.click(within(row).getByRole("button", { name: "Архив" }));

    await waitFor(() => {
      expect(mockPatch).toHaveBeenCalledWith({ id: 10, payload: { archive: true } });
    });
    expect(await screen.findByText(/Карта 7001 архивирована/)).toBeInTheDocument();
  });

  it("shows human-readable API validation error", async () => {
    renderView();
    mockPatch.mockRejectedValue(
      new ApiError("dup", 422, "card_number «7001» already exists", "gsm_card_duplicate"),
    );

    const select = screen.getByLabelText("Машина для карты 7002");
    fireEvent.change(select, { target: { value: "2" } });

    expect(await screen.findByText(/Карта «7001» уже существует/)).toBeInTheDocument();
  });
});
