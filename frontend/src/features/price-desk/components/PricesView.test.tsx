import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { PricesView } from "@/features/price-desk/components/PricesView";
import type { PriceDeskPreview, PriceDeskStatus } from "@/features/price-desk/types/priceDesk";
import type { DuplicateTask, GuidTasksResponse } from "@/features/nomenclature/types/nomenclature";
import { ApiError } from "@/shared/lib/apiError";

const mockPreview = vi.fn();
const mockApply = vi.fn();
const mockResolvePrice = vi.fn();
const mockResolveDuplicate = vi.fn();

const tasksState: GuidTasksResponse = {
  to_create_1c: [],
  to_price: [],
  duplicates: [],
};

vi.mock("@/features/price-desk/hooks/usePriceDeskQueries", () => ({
  usePriceDeskStatusQuery: () => ({
    data: STATUS,
    isError: false,
    error: null,
  }),
  usePriceDeskPreviewMutation: () => ({
    mutateAsync: mockPreview,
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
  usePriceDeskApplyMutation: () => ({
    mutateAsync: mockApply,
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
}));

vi.mock("@/features/nomenclature/hooks/useNomenclatureQueries", () => ({
  useImport1cMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
  useGuidTasksQuery: () => ({
    data: tasksState,
    isError: false,
    isLoading: false,
    error: null,
  }),
  useGuidDuplicatesQuery: () => ({
    data: { to_create_1c: [], to_price: [], duplicates: tasksState.duplicates },
    isError: false,
    isLoading: false,
    error: null,
  }),
  useResolvePriceMutation: () => ({
    mutateAsync: mockResolvePrice,
    isPending: false,
  }),
  useResolveDuplicateMutation: () => ({
    mutateAsync: mockResolveDuplicate,
    isPending: false,
  }),
}));

vi.mock("@/features/nomenclature/components/Import1cDialog", () => ({
  Import1cDialog: ({ open }: { open: boolean }) =>
    open ? <div data-testid="import-1c-dialog">import-1c-dialog</div> : null,
}));

const STATUS: PriceDeskStatus = {
  groups: [
    { product_kind: "plates", price_list_date: "2026-08-17", imported_at: "2026-09-15", row_count: 277 },
    { product_kind: "fbs", price_list_date: null, imported_at: null, row_count: 0 },
    { product_kind: "march", price_list_date: null, imported_at: null, row_count: 0 },
    { product_kind: "step", price_list_date: null, imported_at: null, row_count: 0 },
    { product_kind: "bridge_pile", price_list_date: null, imported_at: null, row_count: 0 },
    { product_kind: "pile", price_list_date: null, imported_at: null, row_count: 0 },
  ],
};

const PREVIEW: PriceDeskPreview = {
  product_kind: "fbs",
  price_list_date: "2026-09-07",
  file_sha256: "abc123",
  parsed_rows: 8,
  changed: 1,
  new: 2,
  missing: 1,
  unchanged: 4,
  examples: {
    changed: [{ key: "ФБС 9.3.6-Т / B25", old_price: 1700, new_price: 1788.33 }],
    new: [{ key: "ФБС 12.4.6-Т / B25", old_price: null, new_price: 2951.52 }],
    missing: [{ key: "ФБС 24.6.6-Т / B25", old_price: 1000, new_price: null }],
    unchanged: [],
  },
};

const resetTasks = () => {
  tasksState.to_create_1c = [];
  tasksState.to_price = [];
  tasksState.duplicates = [];
};

describe("PricesView", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
    resetTasks();
  });

  it("shows group dates and hides write until preview", () => {
    render(<PricesView />);
    expect(screen.getByText("Плиты ПБ")).toBeInTheDocument();
    expect(screen.getByText("2026-08-17")).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Записать в программу" })).not.toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Загрузка прайса завода" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Выгрузка из 1С" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Задачи по изделиям" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Дубли GUID" })).toBeInTheDocument();
  });

  it("opens 1C import dialog from the prices tab", () => {
    render(<PricesView />);
    fireEvent.click(screen.getByRole("button", { name: "Загрузить выгрузку 1С" }));
    expect(screen.getByTestId("import-1c-dialog")).toBeInTheDocument();
  });

  it("previews on file select, shows diff, then apply sends sha256", async () => {
    mockPreview.mockResolvedValue(PREVIEW);
    mockApply.mockResolvedValue(PREVIEW);
    render(<PricesView />);

    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    const file = new File(["xlsx"], "Прайс на ФБС от 07.09.2026.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    fireEvent.change(input, { target: { files: [file] } });

    await waitFor(() => {
      expect(mockPreview).toHaveBeenCalledWith(file);
    });

    expect(await screen.findByText(/изменено 1, новых 2, пропало 1/)).toBeInTheDocument();
    expect(screen.getByText(/ФБС 9.3.6-Т \/ B25/)).toBeInTheDocument();
    expect(mockApply).not.toHaveBeenCalled();

    fireEvent.click(screen.getByRole("button", { name: "Записать в программу" }));
    await waitFor(() => {
      expect(mockApply).toHaveBeenCalledWith({ file, fileSha256: "abc123" });
    });
  });

  it("shows empty tasks copy when queues are empty", () => {
    render(<PricesView />);
    expect(screen.getByText("задач нет")).toBeInTheDocument();
    expect(screen.getByText("дублей нет")).toBeInTheDocument();
  });

  it("submits inline price and closes the 💰 task", async () => {
    tasksState.to_price = [
      { guid: "aaa-111", product_kind: "pile", mark: "С30.30-3", name: "Сваи С30.30-3" },
    ];
    mockResolvePrice.mockImplementation(async () => {
      tasksState.to_price = [];
      return { guid: "aaa-111", mark: "С30.30-3", product_kind: "pile", price: 1500, state: "resolved" };
    });
    const { rerender } = render(<PricesView />);
    fireEvent.change(screen.getByLabelText("Цена для С30.30-3"), { target: { value: "1500" } });
    fireEvent.click(screen.getByRole("button", { name: "Записать цену" }));
    await waitFor(() => {
      expect(mockResolvePrice).toHaveBeenCalledWith({ guid: "aaa-111", price: 1500 });
    });
    rerender(<PricesView />);
    expect(screen.queryByLabelText("Цена для С30.30-3")).not.toBeInTheDocument();
    expect(screen.getByText("задач нет")).toBeInTheDocument();
  });

  it("resolves a duplicate and shows 409 as a live error", async () => {
    const dup: DuplicateTask = {
      scope: "nonplate",
      key: "С40.30-6",
      product_kind: "pile",
      candidates: [
        { guid: "bbb-222", name: "Сваи С40.30-6 А", price: 1000 },
        { guid: "ccc-333", name: "Сваи С40.30-6 Б", price: 1000 },
      ],
    };
    tasksState.duplicates = [dup];
    mockResolveDuplicate.mockRejectedValue(new ApiError("кандидат исчез", 409));
    render(<PricesView />);
    fireEvent.click(screen.getByRole("radio", { name: /Сваи С40.30-6 Б/ }));
    fireEvent.click(screen.getByRole("button", { name: "Запомнить выбор" }));
    expect(await screen.findByText("кандидат исчез, задача открыта")).toBeInTheDocument();
  });
});
