import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { SgpPlateItem } from "@/features/production/api/sgpApi";
import { SgpWarehouseView } from "@/features/production/components/SgpWarehouseView";

const mockUseSgpPlatesQuery = vi.fn();
const mockUnlinkMutateAsync = vi.fn();
const mockRelinkMutateAsync = vi.fn();

vi.mock("@/features/production/hooks/useSgpQueries", () => ({
  useSgpPlatesQuery: (...args: unknown[]) => mockUseSgpPlatesQuery(...args),
  useSgpUnlinkMutation: () => ({
    mutateAsync: mockUnlinkMutateAsync,
    isPending: false,
  }),
  useSgpRelinkMutation: () => ({
    mutateAsync: mockRelinkMutateAsync,
    isPending: false,
  }),
}));

function makePlate(overrides: Partial<SgpPlateItem> & Pick<SgpPlateItem, "id">): SgpPlateItem {
  return {
    kp_id: 1,
    plate_name: "Плиты ПБ 45-12-6п",
    length_m: 4.5,
    width_m: 1.2,
    load_class: 600,
    qty: 1,
    completed_date: "2026-07-27",
    production_day: 1,
    plan_id: "plan-1",
    nomenclature_id: null,
    customer_name: "Клиент",
    execution_terms: "03.08.2026",
    sgp_progress: { n: 1, m: 10 },
    ...overrides,
  };
}

function mockPlates(items: SgpPlateItem[]) {
  mockUseSgpPlatesQuery.mockReturnValue({
    data: { items, count: items.length, filter: "all" },
    isLoading: false,
    isError: false,
    error: null,
  });
}

describe("SgpWarehouseView unlink/relink UX", () => {
  let scrollIntoViewMock: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    scrollIntoViewMock = vi.fn();
    Element.prototype.scrollIntoView = scrollIntoViewMock;

    mockUnlinkMutateAsync.mockResolvedValue({
      ok: true,
      message: "Отвязано",
      sgp_id: 1,
      qty: 1,
      kp_id: null,
      target_kp_id: null,
    });
    mockRelinkMutateAsync.mockResolvedValue({
      ok: true,
      message: "Перепривязано",
      sgp_id: 2,
      qty: 1,
      kp_id: 2,
      target_kp_id: 2,
    });
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("opens unlink confirm inline under the clicked row", () => {
    mockPlates([
      makePlate({ id: 1, plate_name: "Плита A", qty: 1 }),
      makePlate({ id: 2, plate_name: "Плита B", qty: 1 }),
    ]);

    render(<SgpWarehouseView />);

    const rows = screen.getAllByRole("row");
    const firstDataRow = rows[1];
    fireEvent.click(within(firstDataRow).getByRole("button", { name: "Отвязать" }));

    expect(screen.getByText("Отвязать 1 шт от КП #1?")).toBeInTheDocument();
    expect(screen.queryByRole("spinbutton")).not.toBeInTheDocument();

    const tbody = screen.getAllByRole("rowgroup")[1];
    const tbodyRows = within(tbody).getAllByRole("row");
    expect(tbodyRows).toHaveLength(3);
    expect(within(tbodyRows[1]).getByText("Отвязать 1 шт от КП #1?")).toBeInTheDocument();
    expect(within(tbodyRows[2]).getByText("Плита B")).toBeInTheDocument();
  });

  it("shows quantity input for multi-qty unlink inline under the row", () => {
    mockPlates([makePlate({ id: 10, qty: 4, plate_name: "Плита bulk" })]);

    render(<SgpWarehouseView />);
    fireEvent.click(screen.getByRole("button", { name: "Отвязать" }));

    expect(screen.getByText("Отвязать от КП #1: Плита bulk")).toBeInTheDocument();
    expect(screen.getByRole("spinbutton")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Подтвердить" })).toBeInTheDocument();
  });

  it("closes inline unlink panel when row action toggles to Отмена", () => {
    mockPlates([makePlate({ id: 1, qty: 1 })]);

    render(<SgpWarehouseView />);

    fireEvent.click(screen.getByRole("button", { name: "Отвязать" }));
    expect(screen.getByText("Отвязать 1 шт от КП #1?")).toBeInTheDocument();

    fireEvent.click(screen.getAllByRole("button", { name: "Отмена" })[0]);
    expect(screen.queryByText("Отвязать 1 шт от КП #1?")).not.toBeInTheDocument();
  });

  it("submits qty=1 for single-qty unlink via Да, отвязать", async () => {
    mockPlates([makePlate({ id: 7, qty: 1 })]);

    render(<SgpWarehouseView />);

    fireEvent.click(screen.getByRole("button", { name: "Отвязать" }));
    fireEvent.click(screen.getByRole("button", { name: "Да, отвязать" }));

    expect(mockUnlinkMutateAsync).toHaveBeenCalledWith({ sgpId: 7, qty: 1 });
    expect(await screen.findByText("Отвязано")).toBeInTheDocument();
  });

  it("scrolls active row into view when opening unlink", () => {
    mockPlates([makePlate({ id: 1, qty: 1 })]);

    render(<SgpWarehouseView />);
    fireEvent.click(screen.getByRole("button", { name: "Отвязать" }));

    expect(scrollIntoViewMock).toHaveBeenCalledWith({ block: "nearest", behavior: "smooth" });
  });

  it("opens relink form in a modal dialog", () => {
    mockPlates([
      makePlate({
        id: 3,
        kp_id: null,
        plate_name: "Свободная плита",
        sgp_progress: null,
      }),
    ]);

    render(<SgpWarehouseView />);
    fireEvent.click(screen.getByRole("button", { name: "Перепривязать" }));

    const dialog = screen.getByRole("dialog");
    expect(
      within(dialog).getByRole("heading", { name: "Перепривязать: Свободная плита" }),
    ).toBeInTheDocument();
    expect(within(dialog).getByPlaceholderText("например 42")).toBeInTheDocument();
  });

  it("closes unlink panel when relink modal opens", () => {
    mockPlates([
      makePlate({ id: 1, qty: 1, plate_name: "Linked" }),
      makePlate({
        id: 2,
        kp_id: null,
        plate_name: "Free",
        sgp_progress: null,
      }),
    ]);

    render(<SgpWarehouseView />);

    fireEvent.click(screen.getAllByRole("button", { name: "Отвязать" })[0]);
    expect(screen.getByText("Отвязать 1 шт от КП #1?")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Перепривязать" }));
    expect(screen.queryByText("Отвязать 1 шт от КП #1?")).not.toBeInTheDocument();
    expect(screen.getByRole("dialog")).toBeInTheDocument();
  });
});
