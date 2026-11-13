import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { MemoryRouter } from "react-router";
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

const duplicateItems: DuplicateTask[] = [];

const listFlags = {
  tasksLoading: false,
  tasksError: false,
  duplicatesLoading: false,
  duplicatesError: false,
};

vi.mock("@/features/price-desk/hooks/usePileCostQueries", () => ({
  usePileCostQuery: () => ({
    data: undefined,
    isLoading: false,
    isFetching: false,
    isError: false,
    error: null,
  }),
  useSavePileCostInputsMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
  usePileCostPreviewMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
  usePileCostApplyMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
}));

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
    isError: listFlags.tasksError,
    isLoading: listFlags.tasksLoading,
    error: listFlags.tasksError ? new Error("не удалось загрузить задачи") : null,
  }),
  useGuidDuplicatesQuery: () => ({
    data: { to_create_1c: [], to_price: [], duplicates: duplicateItems },
    isError: listFlags.duplicatesError,
    isLoading: listFlags.duplicatesLoading,
    error: listFlags.duplicatesError ? new Error("не удалось загрузить дубли") : null,
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
  duplicateItems.splice(0, duplicateItems.length);
  listFlags.tasksLoading = false;
  listFlags.tasksError = false;
  listFlags.duplicatesLoading = false;
  listFlags.duplicatesError = false;
};

const pricesTree = (path = "/prices") => (
  <MemoryRouter initialEntries={[path]}>
    <PricesView />
  </MemoryRouter>
);

const renderPrices = (path = "/prices") => render(pricesTree(path));

describe("PricesView", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
    resetTasks();
  });

  it("shows group dates and hides write until preview", () => {
    renderPrices("/prices?view=update");
    expect(screen.getByRole("heading", { name: "Обновить цены" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Сейчас в программе" })).toBeInTheDocument();
    expect(screen.getByText("Плиты ПБ")).toBeInTheDocument();
    expect(screen.getByText("2026-08-17")).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Записать в программу" })).not.toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Загрузка прайса завода" })).toBeInTheDocument();
    expect(screen.queryByRole("region", { name: "Цельные сваи" })).not.toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Выгрузка из 1С" })).not.toBeInTheDocument();
    expect(screen.getByRole("link", { name: "Ждут решения" })).toHaveAttribute(
      "href",
      "/prices?view=queue",
    );
  });

  it("opens 1C import dialog from the prices tab", () => {
    renderPrices();
    expect(screen.getByRole("heading", { name: "Ждут решения" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Выгрузка из 1С" })).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Загрузить выгрузку 1С" }));
    expect(screen.getByTestId("import-1c-dialog")).toBeInTheDocument();
  });

  it("previews on file select, shows diff, then apply sends sha256", async () => {
    mockPreview.mockResolvedValue(PREVIEW);
    mockApply.mockResolvedValue(PREVIEW);
    renderPrices("/prices?view=update");

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
    expect(screen.queryByRole("region", { name: "Цельные сваи" })).not.toBeInTheDocument();
  });

  it("shows empty tasks copy when queues are empty", () => {
    renderPrices();
    const queue = screen.getByRole("heading", { name: "Ждут решения" });
    const cost = screen.getByRole("link", { name: "Себестоимость" });
    const update = screen.getByRole("link", { name: "Обновить цены" });
    const empty = screen.getByText("задач нет");
    const upload = screen.getByRole("heading", { name: "Выгрузка из 1С" });
    expect(queue.compareDocumentPosition(cost) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(cost.compareDocumentPosition(empty) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(empty.compareDocumentPosition(upload) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(cost).toHaveAttribute("href", "/prices?view=cost");
    expect(update).toHaveAttribute("href", "/prices?view=update");
    expect(screen.queryByText("дублей нет")).not.toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Завести в 1С" })).not.toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Дубли GUID" })).not.toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Задачи по изделиям" })).not.toBeInTheDocument();
    expect(
      screen.getByText(/Сначала очередь «Ждут решения», а себестоимость и прайс завода/),
    ).toBeInTheDocument();
  });

  it("renders КП numbers on a create-1c demand row", () => {
    tasksState.to_create_1c = [
      {
        product_kind: "pile",
        mark: "С70.35-9у",
        hint: "заведите карточку «у» в 1С и загрузите отчёт",
        field: "guid_1c_u",
        kp_ids: [12, 18],
      },
    ];
    renderPrices();
    expect(screen.getByText(/КП №12, №18/)).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Завести в 1С" })).toBeInTheDocument();
    expect(screen.queryByText("задач нет")).not.toBeInTheDocument();
    expect(screen.getByRole("link", { name: "Себестоимость" })).toBeInTheDocument();
    expect(screen.getByRole("link", { name: "Обновить цены" })).toBeInTheDocument();
  });

  it("keeps the hint and offers the explain panel on a create-1c row", () => {
    tasksState.to_create_1c = [
      {
        product_kind: "pile",
        mark: "С70.35-9у",
        hint: "заведите карточку «у» в 1С и загрузите отчёт",
        field: "guid_1c_u",
        kp_ids: [12],
      },
    ];
    renderPrices();
    expect(screen.getByText(/заведите карточку «у» в 1С/)).toBeInTheDocument();
    expect(
      screen.getByRole("button", { name: /Почему нет номера/ }),
    ).toHaveAttribute("aria-expanded", "false");
  });

  it("submits inline price and closes the 💰 task", async () => {
    tasksState.to_price = [
      { guid: "aaa-111", product_kind: "pile", mark: "С30.30-3", name: "Сваи С30.30-3" },
    ];
    mockResolvePrice.mockImplementation(async () => {
      tasksState.to_price = [];
      return { guid: "aaa-111", mark: "С30.30-3", product_kind: "pile", price: 1500, state: "resolved" };
    });
    const { rerender } = renderPrices();
    const priceRow = screen.getByLabelText("Цена для С30.30-3").closest("li");
    expect(priceRow).not.toHaveTextContent(/КП №/);
    expect(screen.queryByText("задач нет")).not.toBeInTheDocument();
    expect(screen.getByRole("link", { name: "Обновить цены" })).toBeInTheDocument();
    fireEvent.change(screen.getByLabelText("Цена для С30.30-3"), { target: { value: "1500" } });
    fireEvent.click(screen.getByRole("button", { name: "Записать цену" }));
    await waitFor(() => {
      expect(mockResolvePrice).toHaveBeenCalledWith({ guid: "aaa-111", price: 1500 });
    });
    rerender(pricesTree());
    expect(screen.queryByLabelText("Цена для С30.30-3")).not.toBeInTheDocument();
    expect(screen.getByText("задач нет")).toBeInTheDocument();
    expect(screen.getByRole("link", { name: "Себестоимость" })).toBeInTheDocument();
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
      kp_ids: [12],
    };
    duplicateItems.push(dup);
    mockResolveDuplicate.mockRejectedValue(new ApiError("кандидат исчез", 409));
    renderPrices();
    expect(screen.getByRole("heading", { name: "Дубли GUID" })).toBeInTheDocument();
    expect(screen.getByText(/КП №12/)).toBeInTheDocument();
    expect(screen.queryByText("задач нет")).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: /Загрузить выгрузку 1С, С40.30-6/ })).not.toBeInTheDocument();
    fireEvent.click(screen.getByRole("radio", { name: /Сваи С40.30-6 Б/ }));
    fireEvent.click(screen.getByRole("button", { name: "Запомнить выбор" }));
    expect(await screen.findByText("кандидат исчез, задача открыта")).toBeInTheDocument();
  });

  it("opens the queue for a bare path, view=queue and an unknown view", () => {
    for (const path of ["/prices", "/prices?view=queue", "/prices?view=other"]) {
      const view = renderPrices(path);
      expect(screen.getByRole("heading", { name: "Ждут решения" })).toBeInTheDocument();
      expect(screen.queryByRole("heading", { name: "Сейчас в программе" })).not.toBeInTheDocument();
      expect(screen.queryByRole("region", { name: "Цельные сваи" })).not.toBeInTheDocument();
      expect(screen.getByRole("link", { name: "Себестоимость" })).toBeInTheDocument();
      view.unmount();
    }
  });

  it("opens cost without the factory file and returns to the queue", () => {
    renderPrices("/prices?view=cost");
    expect(screen.getByRole("heading", { name: "Себестоимость" })).toBeInTheDocument();
    expect(screen.getByRole("region", { name: "Цельные сваи" })).toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Загрузка прайса завода" })).not.toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Выгрузка из 1С" })).not.toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Ждут решения" })).not.toBeInTheDocument();
    expect(screen.queryByText("задач нет")).not.toBeInTheDocument();
    fireEvent.click(screen.getByRole("link", { name: "Ждут решения" }));
    expect(screen.getByRole("heading", { name: "Ждут решения" })).toBeInTheDocument();
    expect(screen.queryByRole("region", { name: "Цельные сваи" })).not.toBeInTheDocument();
  });

  it("opens the factory price from the queue and hides pile cost", () => {
    renderPrices();
    fireEvent.click(screen.getByRole("link", { name: "Обновить цены" }));
    expect(screen.getByRole("heading", { name: "Обновить цены" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Сейчас в программе" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "Загрузка прайса завода" })).toBeInTheDocument();
    expect(screen.queryByRole("region", { name: "Цельные сваи" })).not.toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Выгрузка из 1С" })).not.toBeInTheDocument();
    fireEvent.click(screen.getByRole("link", { name: "Ждут решения" }));
    expect(screen.getByRole("heading", { name: "Ждут решения" })).toBeInTheDocument();
  });

  it("does not call the queue empty while a list is loading or failed", () => {
    listFlags.duplicatesLoading = true;
    const loading = renderPrices();
    expect(screen.queryByText("задач нет")).not.toBeInTheDocument();
    expect(screen.getByRole("link", { name: "Обновить цены" })).toBeInTheDocument();
    loading.unmount();

    listFlags.duplicatesLoading = false;
    listFlags.tasksError = true;
    const failedTasks = renderPrices();
    expect(screen.getByText("не удалось загрузить задачи")).toBeInTheDocument();
    expect(screen.queryByText("задач нет")).not.toBeInTheDocument();
    failedTasks.unmount();

    listFlags.tasksError = false;
    listFlags.duplicatesError = true;
    renderPrices();
    expect(screen.getByRole("heading", { name: "Дубли GUID" })).toBeInTheDocument();
    expect(screen.getByText("не удалось загрузить дубли")).toBeInTheDocument();
    expect(screen.queryByText("дублей нет")).not.toBeInTheDocument();
    expect(screen.queryByText("задач нет")).not.toBeInTheDocument();
  });

  it("ignores duplicates parked on the tasks payload", () => {
    tasksState.duplicates = [
      {
        scope: "nonplate",
        key: "С40.30-6",
        product_kind: "pile",
        candidates: [{ guid: "bbb-222", name: "Сваи С40.30-6 А", price: 1000 }],
        kp_ids: [12],
      },
    ];
    renderPrices();
    expect(screen.getByText("задач нет")).toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Дубли GUID" })).not.toBeInTheDocument();
    expect(screen.queryByText(/КП №12/)).not.toBeInTheDocument();
  });

  it("keeps create, price and duplicate rows in separate blocks", () => {
    tasksState.to_create_1c = [
      {
        product_kind: "pile",
        mark: "С70.35-9у",
        hint: "нет карточки",
        kp_ids: [12],
      },
    ];
    tasksState.to_price = [
      { guid: "aaa-111", product_kind: "pile", mark: "С30.30-3", name: "Сваи С30.30-3" },
    ];
    duplicateItems.push({
      scope: "nonplate",
      key: "С40.30-6",
      product_kind: "pile",
      candidates: [
        { guid: "bbb-222", name: "Сваи С40.30-6 А", price: 1000 },
        { guid: "ccc-333", name: "Сваи С40.30-6 Б", price: 1100 },
      ],
      kp_ids: [9],
    });
    renderPrices();
    const create = screen.getByRole("heading", { name: "Завести в 1С" });
    const price = screen.getByRole("heading", { name: "Ввести цену" });
    const duplicates = screen.getByRole("heading", { name: "Дубли GUID" });
    const upload = screen.getByRole("heading", { name: "Выгрузка из 1С" });
    expect(create.compareDocumentPosition(price) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(price.compareDocumentPosition(duplicates) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(duplicates.compareDocumentPosition(upload) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(screen.getByText(/КП №12/)).toBeInTheDocument();
    expect(screen.getByText(/КП №9/)).toBeInTheDocument();
    expect(screen.getByLabelText("Цена для С30.30-3").closest("li")).not.toHaveTextContent(/КП №/);
    expect(screen.queryByText("задач нет")).not.toBeInTheDocument();
  });

  it("opens one import dialog from a non-plate row and leaves the plate row", () => {
    tasksState.to_create_1c = [
      {
        product_kind: "pile",
        mark: "С70.35-9у",
        hint: "нет карточки",
        kp_ids: [12],
      },
      {
        product_kind: "composite_pile",
        mark: "СНМ 30.30",
        hint: "составная свая",
        kp_ids: [5],
      },
      {
        product_kind: "plate",
        mark: "ПБ 56-12",
        hint: "плита без карточки",
        kp_ids: [4],
      },
      {
        product_kind: "plates",
        mark: "ПБ 70-12",
        hint: "плиты без карточки",
        kp_ids: [7],
      },
    ];
    renderPrices();
    expect(screen.queryByRole("button", { name: "Загрузить выгрузку 1С, ПБ 56-12" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Загрузить выгрузку 1С, ПБ 70-12" })).not.toBeInTheDocument();
    expect(screen.getAllByRole("button", { name: /Почему нет номера/ })).toHaveLength(4);
    fireEvent.click(screen.getByRole("button", { name: "Загрузить выгрузку 1С, С70.35-9у" }));
    expect(screen.getAllByTestId("import-1c-dialog")).toHaveLength(1);
    expect(screen.getByText(/ПБ 56-12/)).toBeInTheDocument();
    expect(screen.getByText(/ПБ 70-12/)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Загрузить выгрузку 1С, СНМ 30.30" }));
    expect(screen.getAllByTestId("import-1c-dialog")).toHaveLength(1);
    fireEvent.click(screen.getByRole("button", { name: "Загрузить выгрузку 1С" }));
    expect(screen.getAllByTestId("import-1c-dialog")).toHaveLength(1);
  });

  it("keeps the quiet import button when the only create row is a plate", () => {
    tasksState.to_create_1c = [
      {
        product_kind: "plate",
        mark: "ПБ 56-12",
        hint: "плита без карточки",
        kp_ids: [4],
      },
    ];
    renderPrices();
    expect(screen.getByRole("button", { name: "Загрузить выгрузку 1С" })).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Загрузить выгрузку 1С, ПБ 56-12" })).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Почему нет номера/ })).toBeInTheDocument();
  });

  it("does not put a file button on a price row", () => {
    tasksState.to_price = [
      { guid: "aaa-111", product_kind: "pile", mark: "С30.30-3", name: "Сваи С30.30-3" },
    ];
    renderPrices();
    expect(screen.getAllByRole("button", { name: /Загрузить выгрузку 1С/ })).toHaveLength(1);
    expect(screen.getByRole("button", { name: "Загрузить выгрузку 1С" })).toBeInTheDocument();
  });
});
