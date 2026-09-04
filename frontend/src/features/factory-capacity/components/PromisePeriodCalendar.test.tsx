import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { PromisePeriodCalendar } from "@/features/factory-capacity/components/PromisePeriodCalendar";

afterEach(() => {
  cleanup();
});

describe("PromisePeriodCalendar", () => {
  it("clicking two days of the same ISO week emits one week_start", () => {
    const onSelectWeek = vi.fn();
    render(
      <PromisePeriodCalendar
        month="2026-09-01"
        minMonth="2026-08-01"
        maxMonth="2026-10-01"
        onMonthChange={() => undefined}
        selectedWeekStart={null}
        onSelectWeek={onSelectWeek}
        promisedDate="2026-09-11"
      />,
    );

    fireEvent.click(screen.getByTestId("promise-cal-day-2026-09-10"));
    fireEvent.click(screen.getByTestId("promise-cal-day-2026-09-11"));

    expect(onSelectWeek).toHaveBeenCalledTimes(2);
    expect(onSelectWeek).toHaveBeenNthCalledWith(1, "2026-09-07");
    expect(onSelectWeek).toHaveBeenNthCalledWith(2, "2026-09-07");
  });

  it("emits week and day on click without treating the cell as a due date", () => {
    const onSelectWeek = vi.fn();
    const onSelectDay = vi.fn();
    render(
      <PromisePeriodCalendar
        month="2026-09-01"
        minMonth="2026-08-01"
        maxMonth="2026-10-01"
        onMonthChange={() => undefined}
        selectedWeekStart="2026-09-07"
        onSelectWeek={onSelectWeek}
        onSelectDay={onSelectDay}
      />,
    );

    expect(screen.getByTestId("promise-period-calendar")).toBeInTheDocument();
    fireEvent.click(screen.getByTestId("promise-cal-day-2026-09-09"));
    expect(onSelectWeek).toHaveBeenCalledWith("2026-09-07");
    expect(onSelectDay).toHaveBeenCalledWith("2026-09-09");
  });

  it("marks promised_date and disables month nav outside min/max", () => {
    const onMonthChange = vi.fn();
    render(
      <PromisePeriodCalendar
        month="2026-09-01"
        minMonth="2026-09-01"
        maxMonth="2026-09-01"
        onMonthChange={onMonthChange}
        selectedWeekStart={null}
        onSelectWeek={() => undefined}
        promisedDate="2026-09-18"
        firstPourDate="2026-09-14"
        pourFrom="2026-09-14"
        pourToSunday="2026-09-20"
      />,
    );

    expect(screen.getByTestId("promise-cal-promised-marker")).toBeInTheDocument();
    expect(screen.getByTestId("promise-cal-start-marker")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Предыдущий месяц" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Следующий месяц" })).toBeDisabled();
    fireEvent.click(screen.getByRole("button", { name: "Следующий месяц" }));
    expect(onMonthChange).not.toHaveBeenCalled();
  });

  it("paints yellow from first pour day and shows knob fractions", () => {
    render(
      <PromisePeriodCalendar
        month="2026-09-01"
        minMonth="2026-08-01"
        maxMonth="2026-10-01"
        onMonthChange={() => undefined}
        selectedWeekStart={null}
        onSelectWeek={() => undefined}
        firstPourDate="2026-09-09"
        pourFrom="2026-09-09"
        pourToSunday="2026-09-13"
        promisedDate="2026-09-11"
        occupancy={{
          "2026-09-07": 3,
          "2026-09-08": 3,
          "2026-09-09": 1,
          "2026-09-15": 4,
        }}
        knob={3}
      />,
    );

    expect(screen.getByTestId("promise-cal-day-2026-09-07")).not.toHaveAttribute("data-pour");
    expect(screen.getByTestId("promise-cal-day-2026-09-08")).not.toHaveAttribute("data-pour");
    expect(screen.getByTestId("promise-cal-day-2026-09-09")).toHaveAttribute("data-pour", "true");
    expect(screen.getByTestId("promise-cal-day-2026-09-13")).toHaveAttribute("data-pour", "true");
    expect(screen.getByTestId("promise-cal-frac-2026-09-07")).toHaveTextContent("3/3");
    expect(screen.getByTestId("promise-cal-frac-2026-09-09")).toHaveTextContent("1/3");
    expect(screen.getByTestId("promise-cal-frac-2026-09-15")).toHaveTextContent("4/3");
    expect(screen.getByTestId("promise-cal-day-2026-09-15")).toHaveAttribute("data-overflow", "true");
    expect(screen.queryByTestId("promise-cal-frac-2026-09-10")).not.toBeInTheDocument();
    expect(screen.getByTestId("promise-cal-start-marker")).toHaveAttribute(
      "title",
      "начало отливки 2026-09-09",
    );
    expect(screen.getByTestId("promise-cal-promised-marker")).toHaveAttribute(
      "title",
      "дата клиенту 2026-09-11",
    );
  });
});
