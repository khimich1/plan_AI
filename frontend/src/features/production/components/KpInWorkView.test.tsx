import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { KpInWorkView } from "@/features/production/components/KpInWorkView";
import type { KpCandidateItem } from "@/features/production/types/production";

const mockUseKpCandidatesQuery = vi.fn();

vi.mock("@/features/production/hooks/useProductionQueries", () => ({
  useKpCandidatesQuery: (...args: unknown[]) => mockUseKpCandidatesQuery(...args),
}));

function makeItem(overrides: Partial<KpCandidateItem> = {}): KpCandidateItem {
  return {
    kp_id: 47,
    customer_name: "Завод",
    creation_date: "01.03.2026",
    execution_terms: "01.09.2026",
    total_plates: 10,
    completed_plates: 0,
    completion_pct: 0,
    in_plan_pct: 60,
    total_length_m: 60,
    remaining_qty: 4,
    in_plan_qty: 6,
    on_sgp_qty: 0,
    plates: [
      {
        id: 1,
        plate_name: "ПБ 60-12-8п",
        length_m: 6,
        width_m: 1.2,
        load_class: 800,
        qty: 4,
        bucket: "awaiting_plan",
      },
      {
        id: 2,
        plate_name: "ПБ 51-12-8п",
        length_m: 5.1,
        width_m: 1.2,
        load_class: 800,
        qty: 6,
        bucket: "in_plan",
      },
    ],
    ...overrides,
  };
}

function mockQuery(items: KpCandidateItem[]) {
  mockUseKpCandidatesQuery.mockReturnValue({
    data: { items, count: items.length },
    isLoading: false,
    isError: false,
    error: null,
  });
}

describe("KpInWorkView", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("requests in_work scope", () => {
    mockQuery([]);
    render(<KpInWorkView />);
    expect(mockUseKpCandidatesQuery).toHaveBeenCalledWith(true, "in_work");
  });

  it("shows empty state without money", () => {
    mockQuery([]);
    render(<KpInWorkView />);
    expect(
      screen.getByText("Все плиты на СГП — смотрите склад готовой продукции"),
    ).toBeInTheDocument();
    expect(screen.queryByText(/₽|руб|total_amount/i)).not.toBeInTheDocument();
  });

  it("renders row fields and expands plates with in-plan label", () => {
    mockQuery([makeItem()]);
    render(<KpInWorkView />);

    expect(screen.getByText(/КП №47/)).toBeInTheDocument();
    expect(screen.getByText("Завод")).toBeInTheDocument();
    expect(screen.getByText("осталось 10 шт")).toBeInTheDocument();
    expect(screen.getByText("в плане 6")).toBeInTheDocument();
    expect(screen.getByText("на СГП 0")).toBeInTheDocument();
    expect(screen.queryByText("ПБ 60-12-8п")).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Развернуть КП №47" }));
    expect(screen.getByText("ПБ 60-12-8п")).toBeInTheDocument();
    expect(screen.getByText("в плане, ждёт отливки")).toBeInTheDocument();
    expect(screen.getByText("ждёт плана")).toBeInTheDocument();
    expect(screen.queryByText(/1[,\s]?200|скидк/i)).not.toBeInTheDocument();
  });

  it("keeps only one row expanded", () => {
    mockQuery([
      makeItem({ kp_id: 1, customer_name: "A" }),
      makeItem({
        kp_id: 2,
        customer_name: "B",
        plates: [
          {
            id: 9,
            plate_name: "Только вторая",
            length_m: 6,
            width_m: 1.2,
            load_class: 800,
            qty: 1,
            bucket: "in_plan",
          },
        ],
      }),
    ]);
    render(<KpInWorkView />);

    fireEvent.click(screen.getByRole("button", { name: "Развернуть КП №1" }));
    expect(screen.getByText("ПБ 60-12-8п")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Развернуть КП №2" }));
    expect(screen.queryByText("ПБ 60-12-8п")).not.toBeInTheDocument();
    expect(screen.getByText("Только вторая")).toBeInTheDocument();
  });

  it("collapses the same row on second click", () => {
    mockQuery([makeItem()]);
    render(<KpInWorkView />);
    fireEvent.click(screen.getByRole("button", { name: "Развернуть КП №47" }));
    expect(screen.getByText("ПБ 60-12-8п")).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Свернуть КП №47" }));
    expect(screen.queryByText("ПБ 60-12-8п")).not.toBeInTheDocument();
  });
});
