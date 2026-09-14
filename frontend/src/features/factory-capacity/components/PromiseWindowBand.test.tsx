import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render, screen } from "@testing-library/react";
import { PromiseWindowBand } from "@/features/factory-capacity/components/PromiseWindowBand";

afterEach(() => {
  cleanup();
});

describe("PromiseWindowBand", () => {
  it("shows the quote window range and a marker on promised_date", () => {
    render(
      <PromiseWindowBand
        window={{
          from_week: "2026-09-07",
          to_week: "2026-09-07",
          promised_date: "2026-09-11",
        }}
        firstPourDate="2026-09-09"
        pourToSunday="2026-09-13"
      />,
    );

    const band = screen.getByTestId("promise-window-band");
    expect(band).toHaveTextContent("9.09");
    expect(band).toHaveTextContent("13.09");
    expect(band).not.toHaveTextContent("7.09");
    expect(band).toHaveTextContent("дата клиенту 11.09");
    expect(screen.getByTestId("promise-window-band-marker")).toHaveAttribute(
      "title",
      "дата клиенту 11.09",
    );
  });

  it("renders nothing when there is no window", () => {
    const { container } = render(<PromiseWindowBand window={null} />);
    expect(screen.queryByTestId("promise-window-band")).not.toBeInTheDocument();
    expect(container).toBeEmptyDOMElement();
  });
});
