import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { UrgentPosition } from "@/features/production/types/production";
import { UrgentPositionsBlock } from "./UrgentPositionsBlock";

const basePosition = (overrides: Partial<UrgentPosition> = {}): UrgentPosition => ({
  plate_id: 123,
  kp_id: 115,
  plate_name: "ПБ 57-7,2 ×8п",
  qty_remaining: 2,
  deadline: "2026-08-15",
  deadline_source: "delivery_batch",
  deadline_details: [
    {
      type: "delivery_batch",
      batch_name: "3 этаж",
      deadline: "2026-08-15",
      qty: 1,
    },
    {
      type: "execution_terms",
      deadline: "2026-08-26",
      qty: 2,
    },
  ],
  conflict: null,
  ...overrides,
});

describe("UrgentPositionsBlock", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("lists plate_name, deadline, qty_remaining, kp_id", () => {
    render(
      <UrgentPositionsBlock
        positions={[basePosition()]}
        selectedPlatesByKp={{ 115: [123] }}
        onTogglePosition={vi.fn()}
      />,
    );

    expect(screen.getByText("Срочные по срокам")).toBeInTheDocument();
    expect(screen.getByText("ПБ 57-7,2 ×8п")).toBeInTheDocument();
    expect(screen.getByText("15.08")).toBeInTheDocument();
    expect(screen.getByText("2")).toBeInTheDocument();
    expect(screen.getByText("115")).toBeInTheDocument();
  });

  it("expands deadline_details with batches and execution_terms", () => {
    render(
      <UrgentPositionsBlock
        positions={[basePosition()]}
        selectedPlatesByKp={{ 115: [123] }}
        onTogglePosition={vi.fn()}
      />,
    );

    expect(screen.queryByText(/Партия «3 этаж»/)).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Развернуть детали дедлайна" }));

    expect(screen.getByText(/Партия «3 этаж»/)).toBeInTheDocument();
    expect(screen.getByText(/Срок КП:/)).toBeInTheDocument();
  });

  it("shows conflict warning with tooltip title", () => {
    render(
      <UrgentPositionsBlock
        positions={[basePosition({ conflict: "schedule_earlier" })]}
        selectedPlatesByKp={{ 115: [123] }}
        onTogglePosition={vi.fn()}
      />,
    );

    const warn = screen.getByLabelText("Конфликт сроков");
    expect(warn).toHaveAttribute(
      "title",
      "График поставки раньше срока КП более чем на 7 дней",
    );
  });

  it("checkboxes reflect selection; toggle calls onTogglePosition", () => {
    const onToggle = vi.fn();
    const pos = basePosition();
    render(
      <UrgentPositionsBlock
        positions={[pos]}
        selectedPlatesByKp={{ 115: [123] }}
        onTogglePosition={onToggle}
      />,
    );

    const checkbox = screen.getByRole("checkbox", { name: "Выбрать ПБ 57-7,2 ×8п" });
    expect(checkbox).toBeChecked();

    fireEvent.click(checkbox);
    expect(onToggle).toHaveBeenCalledWith(pos);
  });

  it("shows loading and error states", () => {
    const { rerender } = render(
      <UrgentPositionsBlock
        positions={[]}
        selectedPlatesByKp={{}}
        loading
        onTogglePosition={vi.fn()}
      />,
    );
    expect(screen.getByText(/Анализ срочных позиций/)).toBeInTheDocument();

    rerender(
      <UrgentPositionsBlock
        positions={[]}
        selectedPlatesByKp={{}}
        errorMessage="Нет плит «в производстве»"
        onTogglePosition={vi.fn()}
      />,
    );
    expect(screen.getByText("Нет плит «в производстве»")).toBeInTheDocument();
  });

  it("unchecked position is not checked", () => {
    render(
      <UrgentPositionsBlock
        positions={[basePosition()]}
        selectedPlatesByKp={{}}
        onTogglePosition={vi.fn()}
      />,
    );
    const row = screen.getByText("ПБ 57-7,2 ×8п").closest("tr");
    expect(row).not.toBeNull();
    expect(within(row!).getByRole("checkbox")).not.toBeChecked();
  });
});
