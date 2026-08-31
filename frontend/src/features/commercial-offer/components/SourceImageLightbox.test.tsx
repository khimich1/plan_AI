import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { SourceImageLightbox } from "@/features/commercial-offer/components/SourceImageLightbox";

afterEach(() => {
  cleanup();
});

const pages = [
  { id: "a", name: "a.png", previewUrl: "blob:a" },
  { id: "b", name: "b.png", previewUrl: "blob:b" },
  { id: "c", name: "c.png", previewUrl: "blob:c" },
];

describe("SourceImageLightbox", () => {
  it("renders nothing when closed", () => {
    const { container } = render(
      <SourceImageLightbox pages={pages} openId={null} onClose={vi.fn()} onNavigate={vi.fn()} />,
    );
    expect(container).toBeEmptyDOMElement();
  });

  it("shows the open page image when openId is set", () => {
    render(
      <SourceImageLightbox pages={pages} openId="b" onClose={vi.fn()} onNavigate={vi.fn()} />,
    );

    expect(screen.getByRole("dialog")).toBeInTheDocument();
    expect(screen.getByRole("img", { name: "b.png" })).toHaveAttribute("src", "blob:b");
  });

  it("closes on Escape", () => {
    const onClose = vi.fn();
    render(
      <SourceImageLightbox pages={pages} openId="a" onClose={onClose} onNavigate={vi.fn()} />,
    );

    fireEvent.keyDown(window, { key: "Escape" });
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it("closes on backdrop click", () => {
    const onClose = vi.fn();
    render(
      <SourceImageLightbox pages={pages} openId="a" onClose={onClose} onNavigate={vi.fn()} />,
    );

    fireEvent.click(screen.getByRole("dialog"));
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it("closes via the close button", () => {
    const onClose = vi.fn();
    render(
      <SourceImageLightbox pages={pages} openId="a" onClose={onClose} onNavigate={vi.fn()} />,
    );

    fireEvent.click(screen.getByRole("button", { name: /Закрыть/i }));
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it("navigates to next and previous pages with buttons and arrow keys", () => {
    const onNavigate = vi.fn();
    const { rerender } = render(
      <SourceImageLightbox pages={pages} openId="a" onClose={vi.fn()} onNavigate={onNavigate} />,
    );

    fireEvent.click(screen.getByRole("button", { name: /Следующ/i }));
    expect(onNavigate).toHaveBeenCalledWith("b");

    fireEvent.keyDown(window, { key: "ArrowRight" });
    expect(onNavigate).toHaveBeenLastCalledWith("b");

    rerender(
      <SourceImageLightbox pages={pages} openId="b" onClose={vi.fn()} onNavigate={onNavigate} />,
    );

    fireEvent.click(screen.getByRole("button", { name: /Предыдущ/i }));
    expect(onNavigate).toHaveBeenCalledWith("a");

    fireEvent.keyDown(window, { key: "ArrowLeft" });
    expect(onNavigate).toHaveBeenLastCalledWith("a");
  });
});
