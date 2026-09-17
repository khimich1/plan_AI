import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { PriceCatalogDrawer } from "@/features/commercial-offer/components/PriceCatalogDrawer";
import type { PriceCatalogItem } from "@/features/commercial-offer/types/commercialOffer";

afterEach(() => {
  cleanup();
});

const items: PriceCatalogItem[] = [
  { mark: "С110.30-9", concrete_grade: "B25", price: 27585.43 },
  { mark: "С120.35-12", concrete_grade: "B25", price: 44634.03 },
];

describe("PriceCatalogDrawer", () => {
  it("pins a missing mark even when it is not in items", () => {
    render(
      <PriceCatalogDrawer
        open
        onClose={vi.fn()}
        q="C110"
        onQueryChange={vi.fn()}
        items={items}
        missingMarks={["C110.30-6"]}
      />,
    );

    expect(screen.getByText("Прайс")).toBeInTheDocument();
    expect(screen.getByText("В этом КП нет в прайсе")).toBeInTheDocument();
    expect(screen.getByText("C110.30-6")).toBeInTheDocument();
    expect(screen.getByText("С110.30-9")).toBeInTheDocument();
    expect(screen.getByText("27 585,43")).toBeInTheDocument();
  });

  it("calls onQueryChange when the search input changes", () => {
    const onQueryChange = vi.fn();
    render(
      <PriceCatalogDrawer
        open
        onClose={vi.fn()}
        q="C110"
        onQueryChange={onQueryChange}
        items={items}
        missingMarks={["C110.30-6"]}
      />,
    );

    fireEvent.change(screen.getByLabelText("Поиск по прайсу"), { target: { value: "C120" } });
    expect(onQueryChange).toHaveBeenCalledWith("C120");
  });

  it("closes on Escape and overlay click", () => {
    const onClose = vi.fn();
    render(
      <PriceCatalogDrawer
        open
        onClose={onClose}
        q=""
        onQueryChange={vi.fn()}
        items={items}
        missingMarks={[]}
      />,
    );

    fireEvent.keyDown(window, { key: "Escape" });
    expect(onClose).toHaveBeenCalled();

    onClose.mockClear();
    fireEvent.click(screen.getByRole("dialog"));
    expect(onClose).toHaveBeenCalled();
  });

  it("renders nothing when closed", () => {
    const { container } = render(
      <PriceCatalogDrawer
        open={false}
        onClose={vi.fn()}
        q=""
        onQueryChange={vi.fn()}
        items={items}
        missingMarks={["C110.30-6"]}
      />,
    );
    expect(container).toBeEmptyDOMElement();
  });
});
