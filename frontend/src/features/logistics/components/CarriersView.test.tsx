import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { CarriersView } from "@/features/logistics/components/CarriersView";
import type { Carrier } from "@/features/logistics/types/logistics";

const mockUseCarriersQuery = vi.fn();
const mockMerge = vi.fn();

vi.mock("@/features/logistics/hooks/useLogisticsQueries", () => ({
  useCarriersQuery: (...args: unknown[]) => mockUseCarriersQuery(...args),
  useMergeCarrierMutation: () => ({ mutateAsync: mockMerge, isPending: false }),
}));

const CARRIERS: Carrier[] = [
  { id: 1, name: "ООО АвтоЛайн", shipments_count: 3, active: 1, source_sheet: "Перевозчики" },
  { id: 2, name: "ООО ТрансЛогистик", shipments_count: 12, active: 1, source_sheet: "Транспортные Компании" },
  { id: 3, name: "ИП Старый (дубль)", shipments_count: 0, active: 0, source_sheet: "Перевозчики" },
];

describe("CarriersView merge flow", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  const renderView = () => {
    mockUseCarriersQuery.mockReturnValue({
      data: CARRIERS,
      isLoading: false,
      isError: false,
      error: null,
    });
    mockMerge.mockResolvedValue({ moved_shipments: 3 });
    return render(<CarriersView />);
  };

  it("renders carriers table with counts and active flag", () => {
    renderView();
    expect(screen.getByText("ООО АвтоЛайн")).toBeInTheDocument();
    expect(screen.getByText("Транспортные Компании")).toBeInTheDocument();
    expect(screen.getByText("12")).toBeInTheDocument();
    expect(screen.getAllByRole("button", { name: "Слить с…" })).toHaveLength(2);
  });

  it("merges duplicate into target carrier with confirmation", async () => {
    renderView();

    const sourceRow = screen.getByText("ООО АвтоЛайн").closest("tr") as HTMLElement;
    fireEvent.click(within(sourceRow).getByRole("button", { name: "Слить с…" }));

    const dialog = screen.getByRole("dialog");
    expect(within(dialog).getByText(/Все рейсы дубля \(3\) будут перенесены/)).toBeInTheDocument();

    const targetInput = within(dialog).getByPlaceholderText("Начните вводить название целевого");
    fireEvent.change(targetInput, { target: { value: "ООО ТрансЛогистик" } });

    expect(
      within(dialog).getByText(/Будет перенесено рейсов: 3 → «ООО ТрансЛогистик»/),
    ).toBeInTheDocument();

    fireEvent.click(within(dialog).getByRole("button", { name: "Подтвердить слияние" }));

    expect(mockMerge).toHaveBeenCalledWith({ id: 1, intoId: 2 });
    expect(await screen.findByText(/перенесено рейсов — 3/)).toBeInTheDocument();
  });

  it("blocks merging carrier into itself", () => {
    renderView();

    const sourceRow = screen.getByText("ООО АвтоЛайн").closest("tr") as HTMLElement;
    fireEvent.click(within(sourceRow).getByRole("button", { name: "Слить с…" }));

    const dialog = screen.getByRole("dialog");
    const targetInput = within(dialog).getByPlaceholderText("Начните вводить название целевого");
    fireEvent.change(targetInput, { target: { value: "ООО АвтоЛайн" } });

    const confirm = within(dialog).getByRole("button", { name: "Подтвердить слияние" });
    expect(confirm).toBeDisabled();
    expect(mockMerge).not.toHaveBeenCalled();
  });
});
