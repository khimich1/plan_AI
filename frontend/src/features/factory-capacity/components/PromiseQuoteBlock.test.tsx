import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render, screen } from "@testing-library/react";
import { PromiseQuoteBlock } from "@/features/factory-capacity/components/PromiseQuoteBlock";
import type { PromiseQuote } from "@/features/factory-capacity/api/promiseQuote";

afterEach(() => {
  cleanup();
});

const quote: PromiseQuote = {
  tracks: 2,
  solo_days: 1,
  solo_date: "2026-09-04",
  solo_week_end_date: "2026-09-06",
  earliest_start_week: "2026-08-31",
  window: {
    from_week: "2026-08-31",
    to_week: "2026-08-31",
    promised_date: "2026-09-04",
  },
  weeks: [],
  knob: 3,
};

describe("PromiseQuoteBlock", () => {
  it("shows tracks and four quote dates", () => {
    render(<PromiseQuoteBlock quote={quote} />);

    const block = screen.getByTestId("promise-quote-block");
    expect(block).toHaveTextContent("~2 дорожек");
    expect(block).toHaveTextContent(/Обещать к 4\.09/);
    expect(block).toHaveTextContent(/Начало:\s*31\.08/);
    expect(block).toHaveTextContent(/Если только его:\s*4\.09/);
    expect(block).toHaveTextContent(/Соло \+ до конца недели:\s*6\.09/);
  });

  it("uses em dash when window and solo dates are missing", () => {
    render(
      <PromiseQuoteBlock
        quote={{
          ...quote,
          solo_date: null,
          solo_week_end_date: null,
          earliest_start_week: null,
          window: null,
        }}
      />,
    );

    const block = screen.getByTestId("promise-quote-block");
    expect(block).toHaveTextContent("Обещать к —");
    expect(block).toHaveTextContent("Начало: —");
    expect(block).toHaveTextContent("Если только его: —");
    expect(block).toHaveTextContent("Соло + до конца недели: —");
  });
});
