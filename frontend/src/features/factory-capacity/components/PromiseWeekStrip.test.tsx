import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render, screen } from "@testing-library/react";
import { PromiseWeekStrip } from "@/features/factory-capacity/components/PromiseWeekStrip";
import type { PromiseQuoteWeek, PromiseQuoteWindow } from "@/features/factory-capacity/api/promiseQuote";

afterEach(() => {
  cleanup();
});

const week: PromiseQuoteWeek = {
  week_start: "2026-08-31",
  workdays: 5,
  capacity: 15,
  planned: 3,
  promised: 2,
  held: 1,
  free: 10,
};

const window: PromiseQuoteWindow = {
  from_week: "2026-08-31",
  to_week: "2026-08-31",
  promised_date: "2026-09-04",
};

describe("PromiseWeekStrip", () => {
  it("shows planned / promised / held / free for each week", () => {
    render(<PromiseWeekStrip weeks={[week]} quoteWindow={window} />);

    const strip = screen.getByTestId("promise-week-strip");
    expect(strip).toHaveTextContent("31.08");
    expect(strip).toHaveTextContent("план");
    expect(strip).toHaveTextContent("3");
    expect(strip).toHaveTextContent("обещано");
    expect(strip).toHaveTextContent("2");
    expect(strip).toHaveTextContent("холды");
    expect(strip).toHaveTextContent("1");
    expect(strip).toHaveTextContent("свободно");
    expect(strip).toHaveTextContent("10");
  });

  it("marks weeks inside the quote window", () => {
    render(<PromiseWeekStrip weeks={[week]} quoteWindow={window} />);
    expect(screen.getByRole("listitem")).toHaveAttribute("aria-current", "true");
  });

  it("shows empty state when there are no weeks", () => {
    render(<PromiseWeekStrip weeks={[]} />);
    expect(screen.getByTestId("promise-week-strip")).toHaveTextContent("Нет данных по неделям");
  });
});
