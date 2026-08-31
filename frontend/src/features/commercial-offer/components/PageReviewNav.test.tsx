import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { PageReviewNav } from "@/features/commercial-offer/components/PageReviewNav";
import type { PageSource } from "@/features/commercial-offer/lib/multiPageSource";

afterEach(() => {
  cleanup();
});

const page = (id: string, status: PageSource["status"]): PageSource => ({
  id,
  file: new File(["x"], `${id}.png`, { type: "image/png" }),
  name: `${id}.png`,
  previewUrl: `blob:${id}`,
  status,
  batchReviewText: id,
});

describe("PageReviewNav", () => {
  it("shows progress label and navigates among ready pages", () => {
    const onSelect = vi.fn();
    const onPrev = vi.fn();
    const onNext = vi.fn();

    render(
      <PageReviewNav
        pages={[page("a", "ready"), page("b", "ready"), page("c", "running")]}
        activeId="a"
        progressLabel="Распознано 2/3"
        onSelect={onSelect}
        onPrev={onPrev}
        onNext={onNext}
      />,
    );

    expect(screen.getByText("Распознано 2/3")).toBeInTheDocument();
    fireEvent.click(screen.getByTitle("Следующая страница"));
    expect(onNext).toHaveBeenCalled();
    fireEvent.click(screen.getByRole("button", { name: "b.png, готово" }));
    expect(onSelect).toHaveBeenCalledWith("b");
  });
});
