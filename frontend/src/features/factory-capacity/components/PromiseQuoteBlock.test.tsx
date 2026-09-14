import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render, screen } from "@testing-library/react";
import { PromiseQuoteBlock } from "@/features/factory-capacity/components/PromiseQuoteBlock";
import type { PromiseQuote } from "@/features/factory-capacity/api/promiseQuote";

afterEach(() => {
  cleanup();
});

const quote: PromiseQuote = {
  tracks: 5,
  solo_days: 2,
  solo_date: "2026-09-10",
  solo_week_end_date: "2026-09-11",
  earliest_start_week: "2026-09-07",
  first_pour_date: "2026-09-09",
  first_pour_free: 2,
  window: {
    from_week: "2026-09-07",
    to_week: "2026-09-07",
    promised_date: "2026-09-11",
  },
  weeks: [],
  knob: 3,
};

describe("PromiseQuoteBlock", () => {
  it("shows start as first pour day with remainder, not week Monday", () => {
    render(<PromiseQuoteBlock quote={quote} />);

    const block = screen.getByTestId("promise-quote-block");
    expect(block).toHaveTextContent("~5 дорожек");
    expect(block).toHaveTextContent(/Обещать к 11\.09/);
    expect(block).toHaveTextContent(/Начало:\s*9\.09\s*·\s*остаток 2 дор\./);
    expect(block).toHaveTextContent(/Если только его:\s*10\.09/);
    expect(block).toHaveTextContent(/Соло \+ до конца недели:\s*11\.09/);
    expect(block).not.toHaveTextContent("7.09");
  });

  it("uses em dash when first pour and solo dates are missing", () => {
    render(
      <PromiseQuoteBlock
        quote={{
          ...quote,
          solo_date: null,
          solo_week_end_date: null,
          earliest_start_week: null,
          first_pour_date: null,
          first_pour_free: 0,
          window: null,
        }}
      />,
    );

    const block = screen.getByTestId("promise-quote-block");
    expect(block).toHaveTextContent("Обещать к —");
    expect(block).toHaveTextContent("Начало: —");
    expect(block).not.toHaveTextContent("остаток");
    expect(block).toHaveTextContent("Если только его: —");
    expect(block).toHaveTextContent("Соло + до конца недели: —");
  });
});
