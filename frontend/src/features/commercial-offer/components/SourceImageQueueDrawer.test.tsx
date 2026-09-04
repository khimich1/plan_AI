import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { SourceImageQueueDrawer } from "@/features/commercial-offer/components/SourceImageQueueDrawer";
import type { SourceImageQueueItem } from "@/features/commercial-offer/lib/sourceImageQueue";

afterEach(() => {
  cleanup();
});

const items2: SourceImageQueueItem[] = [
  { id: "a", url: "blob:a", name: "a.png" },
  { id: "b", url: "blob:b", name: "b.png" },
];

describe("SourceImageQueueDrawer", () => {
  it("renders nothing when closed", () => {
    const { container } = render(
      <SourceImageQueueDrawer open={false} onClose={vi.fn()} items={items2} />,
    );
    expect(container).toBeEmptyDOMElement();
  });

  it("renders nothing when items are empty even if open", () => {
    const { container } = render(
      <SourceImageQueueDrawer open onClose={vi.fn()} items={[]} />,
    );
    expect(container).toBeEmptyDOMElement();
  });

  it("opens as dialog with title, image, name, and open-in-tab link", () => {
    render(<SourceImageQueueDrawer open onClose={vi.fn()} items={items2} />);

    expect(screen.getByRole("dialog")).toBeInTheDocument();
    expect(screen.getByText("Исходные фото")).toBeInTheDocument();
    expect(screen.getByRole("img", { name: "a.png" })).toHaveAttribute("src", "blob:a");
    expect(screen.getByText("a.png")).toBeInTheDocument();
    expect(screen.getByRole("link", { name: /открыть в новой вкладке/i })).toHaveAttribute(
      "href",
      "blob:a",
    );
    expect(screen.getByText("1 / 2")).toBeInTheDocument();
  });

  it("N=2: next/prev change image and counter (clamp, no wrap)", () => {
    render(<SourceImageQueueDrawer open onClose={vi.fn()} items={items2} />);

    fireEvent.click(screen.getByRole("button", { name: "Следующее фото" }));
    expect(screen.getByRole("img", { name: "b.png" })).toHaveAttribute("src", "blob:b");
    expect(screen.getByText("2 / 2")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Следующее фото" }));
    expect(screen.getByText("2 / 2")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Предыдущее фото" }));
    expect(screen.getByRole("img", { name: "a.png" })).toHaveAttribute("src", "blob:a");
    expect(screen.getByText("1 / 2")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Предыдущее фото" }));
    expect(screen.getByText("1 / 2")).toBeInTheDocument();
  });

  it("N=1: pager arrows are absent or disabled", () => {
    render(
      <SourceImageQueueDrawer
        open
        onClose={vi.fn()}
        items={[{ id: "solo", url: "blob:solo", name: "solo.png" }]}
      />,
    );

    expect(screen.getByText("1 / 1")).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Следующее фото" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Предыдущее фото" })).not.toBeInTheDocument();
  });

  it("calls onClose when close button is clicked", () => {
    const onClose = vi.fn();
    render(<SourceImageQueueDrawer open onClose={onClose} items={items2} />);

    fireEvent.click(screen.getByRole("button", { name: "Закрыть" }));
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it("fits image to drawer width by default and zooms with + / − / По ширине", () => {
    render(<SourceImageQueueDrawer open onClose={vi.fn()} items={items2} />);

    const image = screen.getByRole("img", { name: "a.png" });
    expect(image).toHaveStyle({ width: "100%" });
    expect(screen.getByText("100%")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Увеличить" }));
    expect(image).toHaveStyle({ width: "125%" });
    expect(screen.getByText("125%")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Уменьшить" }));
    expect(image).toHaveStyle({ width: "100%" });

    fireEvent.click(screen.getByRole("button", { name: "Увеличить" }));
    fireEvent.click(screen.getByRole("button", { name: "По ширине" }));
    expect(image).toHaveStyle({ width: "100%" });
    expect(screen.getByText("100%")).toBeInTheDocument();
  });

  it("resets zoom when paging to the next photo", () => {
    render(<SourceImageQueueDrawer open onClose={vi.fn()} items={items2} />);

    fireEvent.click(screen.getByRole("button", { name: "Увеличить" }));
    expect(screen.getByRole("img", { name: "a.png" })).toHaveStyle({ width: "125%" });

    fireEvent.click(screen.getByRole("button", { name: "Следующее фото" }));
    expect(screen.getByRole("img", { name: "b.png" })).toHaveStyle({ width: "100%" });
    expect(screen.getByText("100%")).toBeInTheDocument();
  });

  it("starts at the first photo after close and reopen", () => {
    const onClose = vi.fn();
    const { rerender } = render(
      <SourceImageQueueDrawer open onClose={onClose} items={items2} />,
    );

    fireEvent.click(screen.getByRole("button", { name: "Следующее фото" }));
    expect(screen.getByText("2 / 2")).toBeInTheDocument();

    rerender(<SourceImageQueueDrawer open={false} onClose={onClose} items={items2} />);
    rerender(<SourceImageQueueDrawer open onClose={onClose} items={items2} />);

    expect(screen.getByRole("img", { name: "a.png" })).toBeInTheDocument();
    expect(screen.getByText("1 / 2")).toBeInTheDocument();
  });

  it("Ctrl+wheel zooms; wheel without Ctrl does not", () => {
    render(<SourceImageQueueDrawer open onClose={vi.fn()} items={items2} />);

    const image = screen.getByRole("img", { name: "a.png" });
    const scroller = image.parentElement;
    expect(scroller).toBeTruthy();

    fireEvent.wheel(scroller!, { ctrlKey: false, deltaY: -100 });
    expect(image).toHaveStyle({ width: "100%" });

    fireEvent.wheel(scroller!, { ctrlKey: true, deltaY: -1 });
    expect(image).toHaveStyle({ width: "125%" });

    fireEvent.wheel(scroller!, { ctrlKey: true, deltaY: 1 });
    expect(image).toHaveStyle({ width: "100%" });
  });
});
