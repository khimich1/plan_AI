import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { LineRowActions } from "@/features/commercial-offer/components/LineRowActions";

afterEach(() => {
  cleanup();
});

describe("LineRowActions", () => {
  it("renders edit and delete icon buttons with aria-labels", () => {
    render(
      <LineRowActions
        lineId="ln1"
        defaultQty={2}
        defaultSourceText="ПБ 78-12-8п 2"
        onSave={vi.fn()}
        onDelete={vi.fn()}
      />,
    );
    expect(screen.getByRole("button", { name: "Изменить строку ln1" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Удалить строку ln1" })).toBeInTheDocument();
    expect(screen.queryByText("Удалить")).not.toBeInTheDocument();
  });

  it("opens qty and source fields on pencil and saves qty-only", () => {
    const onSave = vi.fn();
    render(
      <LineRowActions
        lineId="ln1"
        defaultQty={2}
        defaultSourceText="ПБ 78-12-8п 2"
        onSave={onSave}
        onDelete={vi.fn()}
      />,
    );
    fireEvent.click(screen.getByRole("button", { name: "Изменить строку ln1" }));
    const qty = screen.getByRole("spinbutton", { name: "Количество строки ln1" });
    fireEvent.change(qty, { target: { value: "90" } });
    fireEvent.click(screen.getByRole("button", { name: "Сохранить" }));
    expect(onSave).toHaveBeenCalledExactlyOnceWith({ qty: 90 });
  });

  it("saves source_text when the list field changes", () => {
    const onSave = vi.fn();
    render(
      <LineRowActions
        lineId="ln1"
        defaultQty={2}
        defaultSourceText="ПБ 78-12-8п 2"
        onSave={onSave}
        onDelete={vi.fn()}
      />,
    );
    fireEvent.click(screen.getByRole("button", { name: "Изменить строку ln1" }));
    fireEvent.change(screen.getByRole("textbox", { name: "Текст строки ln1" }), {
      target: { value: "ПБ 60-12-8п 3" },
    });
    fireEvent.click(screen.getByRole("button", { name: "Сохранить" }));
    expect(onSave).toHaveBeenCalledExactlyOnceWith({ sourceText: "ПБ 60-12-8п 3" });
  });

  it("Escape cancels editing without saving", () => {
    const onSave = vi.fn();
    render(
      <LineRowActions
        lineId="ln1"
        defaultQty={2}
        defaultSourceText="ПБ 78-12-8п 2"
        onSave={onSave}
        onDelete={vi.fn()}
      />,
    );
    fireEvent.click(screen.getByRole("button", { name: "Изменить строку ln1" }));
    fireEvent.change(screen.getByRole("spinbutton", { name: "Количество строки ln1" }), {
      target: { value: "9" },
    });
    fireEvent.keyDown(screen.getByRole("spinbutton", { name: "Количество строки ln1" }), {
      key: "Escape",
    });
    expect(onSave).not.toHaveBeenCalled();
    expect(screen.getByRole("button", { name: "Изменить строку ln1" })).toBeInTheDocument();
  });

  it("calls onDelete from the trash button", () => {
    const onDelete = vi.fn();
    render(
      <LineRowActions
        lineId="ln1"
        defaultQty={2}
        defaultSourceText="ПБ 78-12-8п 2"
        onSave={vi.fn()}
        onDelete={onDelete}
      />,
    );
    fireEvent.click(screen.getByRole("button", { name: "Удалить строку ln1" }));
    expect(onDelete).toHaveBeenCalledOnce();
  });

  it("does not disable the pencil", () => {
    render(
      <LineRowActions
        lineId="ln1"
        defaultQty={2}
        defaultSourceText="ПБ 78-12-8п 2"
        onSave={vi.fn()}
        onDelete={vi.fn()}
      />,
    );
    expect(screen.getByRole("button", { name: "Изменить строку ln1" })).not.toBeDisabled();
  });

  it("shows a row error message", () => {
    render(
      <LineRowActions
        lineId="ln1"
        defaultQty={2}
        defaultSourceText="ПБ 78-12-8п 2"
        onSave={vi.fn()}
        onDelete={vi.fn()}
        saveError="Не удалось распознать строку."
      />,
    );
    expect(screen.getByText("Не удалось распознать строку.")).toBeInTheDocument();
  });
});
