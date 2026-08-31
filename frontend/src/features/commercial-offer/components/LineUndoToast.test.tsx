import { fireEvent, render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { LineUndoToast } from "@/features/commercial-offer/components/LineUndoToast";

describe("LineUndoToast", () => {
  it("shows the operation message and calls onUndo", () => {
    const onUndo = vi.fn();
    render(<LineUndoToast message="Строка удалена" onUndo={onUndo} />);
    expect(screen.getByText("Строка удалена")).toBeInTheDocument();
    expect(screen.queryByText(/добавлен/i)).not.toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Отменить" }));
    expect(onUndo).toHaveBeenCalledOnce();
  });
});
