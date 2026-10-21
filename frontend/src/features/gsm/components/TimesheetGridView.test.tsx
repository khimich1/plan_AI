import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { TimesheetGridView } from "@/features/gsm/components/TimesheetGridView";
import type { TimesheetGridCell } from "@/features/gsm/types/gsm";

const cells = (): TimesheetGridCell[] => {
  const rows: TimesheetGridCell[] = [];
  for (let day = 1; day <= 31; day += 1) {
    const iso = `2026-07-${String(day).padStart(2, "0")}`;
    rows.push({
      driver_id: 29,
      full_name: "Скрябин Иван",
      personnel_number: "29",
      date: iso,
      code: day <= 9 ? "ОТ" : "Я",
    });
  }
  return rows;
};

describe("TimesheetGridView", () => {
  afterEach(() => cleanup());

  it("highlights Скрябин ОТ on 1–9 July", () => {
    render(<TimesheetGridView period="2026-07" cells={cells()} />);
    expect(screen.getByText("Скрябин Иван")).toBeInTheDocument();
    const first = screen.getByTestId("timesheet-cell-29-2026-07-01");
    const ninth = screen.getByTestId("timesheet-cell-29-2026-07-09");
    const tenth = screen.getByTestId("timesheet-cell-29-2026-07-10");
    expect(first).toHaveTextContent("ОТ");
    expect(first).toHaveAttribute("data-absent", "true");
    expect(ninth).toHaveAttribute("data-absent", "true");
    expect(tenth).toHaveTextContent("Я");
    expect(tenth).toHaveAttribute("data-absent", "false");
  });
});
