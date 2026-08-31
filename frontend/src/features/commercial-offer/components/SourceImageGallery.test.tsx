import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { SourceImageGallery } from "@/features/commercial-offer/components/SourceImageGallery";
import type { PageStatus } from "@/features/commercial-offer/lib/multiPageSource";

afterEach(() => {
  cleanup();
});

const item = (
  id: string,
  status: PageStatus,
  name = `${id}.png`,
): {
  id: string;
  name: string;
  previewUrl: string;
  status: PageStatus;
  errorMessage?: string;
} => ({
  id,
  name,
  previewUrl: `blob:${id}`,
  status,
});

describe("SourceImageGallery", () => {
  it("renders two or more previews", () => {
    render(
      <SourceImageGallery
        pages={[item("a", "pending"), item("b", "pending")]}
        activeId="a"
        onSelect={vi.fn()}
        onRemove={vi.fn()}
      />,
    );

    expect(screen.getByRole("img", { name: "a.png" })).toBeInTheDocument();
    expect(screen.getByRole("img", { name: "b.png" })).toBeInTheDocument();
  });

  it("calls onSelect when a preview is clicked", () => {
    const onSelect = vi.fn();
    render(
      <SourceImageGallery
        pages={[item("a", "pending"), item("b", "pending")]}
        activeId="a"
        onSelect={onSelect}
        onRemove={vi.fn()}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: "b.png, ожидает" }));
    expect(onSelect).toHaveBeenCalledWith("b");
  });

  it("calls onRemove only when the page is removable", () => {
    const onRemove = vi.fn();
    render(
      <SourceImageGallery
        pages={[item("ready", "ready"), item("running", "running")]}
        activeId="ready"
        onSelect={vi.fn()}
        onRemove={onRemove}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: /Удалить ready\.png/i }));
    expect(onRemove).toHaveBeenCalledWith("ready");

    expect(screen.queryByRole("button", { name: /Удалить running\.png/i })).not.toBeInTheDocument();
  });

  it("shows remove for pending and error, hides for confirmed", () => {
    render(
      <SourceImageGallery
        pages={[item("p", "pending"), item("e", "error"), item("c", "confirmed")]}
        activeId="p"
        onSelect={vi.fn()}
        onRemove={vi.fn()}
      />,
    );

    expect(screen.getByRole("button", { name: /Удалить p\.png/i })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Удалить e\.png/i })).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: /Удалить c\.png/i })).not.toBeInTheDocument();
  });

  // S14: thumbnail reflects status in aria-label and visible chrome
  it("includes status in aria-label and shows status chrome", () => {
    render(
      <SourceImageGallery
        pages={[
          item("run", "running"),
          { ...item("err", "error"), errorMessage: "OCR failed" },
          item("ok", "confirmed"),
          item("wait", "pending"),
        ]}
        activeId="err"
        onSelect={vi.fn()}
        onRemove={vi.fn()}
      />,
    );

    expect(screen.getByRole("button", { name: /err\.png.*ошибка/i })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /run\.png.*распозна/i })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /ok\.png.*подтвержд/i })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /wait\.png.*ожидает/i })).toBeInTheDocument();
    expect(screen.getByLabelText(/статус: ошибка/i)).toBeInTheDocument();
    expect(screen.getByLabelText(/статус: распознаётся/i)).toBeInTheDocument();
  });

  // S13: OCR error message is visible when provided for an error page
  it("renders errorMessage alert for the active error page", () => {
    render(
      <SourceImageGallery
        pages={[{ ...item("bad", "error"), errorMessage: "Не удалось распознать" }]}
        activeId="bad"
        onSelect={vi.fn()}
        onRemove={vi.fn()}
        showErrorHint
      />,
    );

    expect(screen.getByText("Не удалось распознать")).toBeInTheDocument();
    expect(
      screen.getByText(/Удалите страницу с ошибкой и добавьте фото в конец/i),
    ).toBeInTheDocument();
  });

  // S19: lightbox before OCR
  it("opens lightbox on thumbnail click when enableLightbox is true", () => {
    const onSelect = vi.fn();
    render(
      <SourceImageGallery
        pages={[item("a", "pending"), item("b", "pending")]}
        activeId="a"
        onSelect={onSelect}
        onRemove={vi.fn()}
        enableLightbox
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: "b.png, ожидает" }));
    expect(onSelect).toHaveBeenCalledWith("b");
    const dialog = screen.getByRole("dialog");
    expect(dialog).toBeInTheDocument();
    expect(dialog.querySelector('img[alt="b.png"]')).toHaveAttribute("src", "blob:b");
  });

  it("does not open lightbox when enableLightbox is false", () => {
    render(
      <SourceImageGallery
        pages={[item("a", "pending"), item("b", "pending")]}
        activeId="a"
        onSelect={vi.fn()}
        onRemove={vi.fn()}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: "b.png, ожидает" }));
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
  });

  it("closes lightbox on Escape without breaking remove", () => {
    const onRemove = vi.fn();
    render(
      <SourceImageGallery
        pages={[item("a", "pending"), item("b", "pending")]}
        activeId="a"
        onSelect={vi.fn()}
        onRemove={onRemove}
        enableLightbox
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: "a.png, ожидает" }));
    expect(screen.getByRole("dialog")).toBeInTheDocument();

    fireEvent.keyDown(window, { key: "Escape" });
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: /Удалить b\.png/i }));
    expect(onRemove).toHaveBeenCalledWith("b");
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
  });
});
