import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import type { ComponentProps } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { MonthCalendarGrid } from "@/features/production/components/MonthCalendarGrid";
import type { DayInfo } from "@/features/production/types/production";

afterEach(() => {
  cleanup();
});

const month = new Date(2026, 7, 1); // August 2026

function renderGrid(
  overrides: Partial<ComponentProps<typeof MonthCalendarGrid>> = {},
) {
  const daysInfo: Record<string, DayInfo> = {
    "2026-08-12": { occupied: 2, max: 5, completed: false, day_number: 1 },
  };
  const onDayActivate = vi.fn();
  const onSaveDayCapacity = vi.fn();
  const onModeChange = vi.fn();

  const result = render(
    <MonthCalendarGrid
      daysInfo={daysInfo}
      holidays={new Set()}
      extraWorkdays={new Set()}
      month={month}
      onMonthChange={vi.fn()}
      selectedDate={null}
      onDayActivate={onDayActivate}
      dayCapacity={{ "2026-08-12": 4, "2026-08-13": 5 }}
      onSaveDayCapacity={onSaveDayCapacity}
      onModeChange={onModeChange}
      mode="planning"
      {...overrides}
    />,
  );

  return { ...result, onDayActivate, onSaveDayCapacity, onModeChange };
}

describe("MonthCalendarGrid capacity mode", () => {
  it("toggles between planning and capacity without calling day activate", () => {
    const { onDayActivate, onModeChange } = renderGrid();

    const capacityTab = screen.getByRole("tab", { name: "Ёмкость" });
    fireEvent.click(capacityTab);

    expect(onModeChange).toHaveBeenCalledWith("capacity");
    expect(onDayActivate).not.toHaveBeenCalled();
  });

  it("shows max_tracks in capacity mode with capacity cell style", () => {
    renderGrid({ mode: "capacity" });

    const editBtn = screen.getByRole("button", { name: "Изменить ёмкость 2026-08-12" });
    expect(editBtn).toHaveTextContent("4");

    const cell = editBtn.closest(".prod-calendar__day");
    expect(cell).toHaveClass("prod-calendar__day--capacity");
  });

  it("opens inline edit and saves via callback", async () => {
    const { onSaveDayCapacity, onDayActivate } = renderGrid({ mode: "capacity" });

    fireEvent.click(screen.getByRole("button", { name: "Изменить ёмкость 2026-08-12" }));

    const input = screen.getByRole("spinbutton", { name: "Ёмкость 2026-08-12" });
    fireEvent.change(input, { target: { value: "3" } });
    fireEvent.click(screen.getByRole("button", { name: "Сохранить ёмкость" }));

    expect(onSaveDayCapacity).toHaveBeenCalledWith("2026-08-12", 3);
    expect(onDayActivate).not.toHaveBeenCalled();
  });

  it("clamps capacity input above hard cap 5", () => {
    const { onSaveDayCapacity } = renderGrid({ mode: "capacity" });

    fireEvent.click(screen.getByRole("button", { name: "Изменить ёмкость 2026-08-12" }));

    const input = screen.getByRole("spinbutton", { name: "Ёмкость 2026-08-12" });
    fireEvent.change(input, { target: { value: "9" } });
    expect(input).toHaveValue(5);
    fireEvent.click(screen.getByRole("button", { name: "Сохранить ёмкость" }));

    expect(onSaveDayCapacity).toHaveBeenCalledWith("2026-08-12", 5);
  });

  it("allows capacity 0", () => {
    const { onSaveDayCapacity } = renderGrid({ mode: "capacity" });

    fireEvent.click(screen.getByRole("button", { name: "Изменить ёмкость 2026-08-12" }));

    const input = screen.getByRole("spinbutton", { name: "Ёмкость 2026-08-12" });
    fireEvent.change(input, { target: { value: "0" } });
    fireEvent.click(screen.getByRole("button", { name: "Сохранить ёмкость" }));

    expect(onSaveDayCapacity).toHaveBeenCalledWith("2026-08-12", 0);
  });

  it("disables increase stepper at hard cap", () => {
    renderGrid({
      mode: "capacity",
      dayCapacity: { "2026-08-12": 5 },
    });

    fireEvent.click(screen.getByRole("button", { name: "Изменить ёмкость 2026-08-12" }));
    expect(screen.getByRole("button", { name: "Увеличить ёмкость" })).toBeDisabled();
  });

  it("does not trigger brush selection when clicking a capacity cell", () => {
    const { onDayActivate } = renderGrid({ mode: "capacity" });

    const editBtn = screen.getByRole("button", { name: "Изменить ёмкость 2026-08-12" });
    const cell = editBtn.closest(".prod-calendar__day-wrap");
    expect(cell).not.toBeNull();
    fireEvent.click(within(cell!).getByText("12"));
    fireEvent.click(editBtn);

    expect(onDayActivate).not.toHaveBeenCalled();
  });

  it("still fires brush activate in planning mode", () => {
    const { onDayActivate } = renderGrid({ mode: "planning" });

    fireEvent.click(screen.getByText("2/5").closest("button")!);

    expect(onDayActivate).toHaveBeenCalledWith("2026-08-12", { shiftKey: false });
  });

  it("keeps planning occupancy display in planning mode", () => {
    renderGrid({ mode: "planning" });

    expect(screen.getByText("2/5")).toBeInTheDocument();
    expect(
      screen.queryByRole("button", { name: "Изменить ёмкость 2026-08-12" }),
    ).not.toBeInTheDocument();
  });
});
