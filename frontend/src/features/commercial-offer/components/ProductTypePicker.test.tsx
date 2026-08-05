import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ProductTypePicker } from "@/features/commercial-offer/components/ProductTypePicker";

afterEach(() => {
  cleanup();
});

describe("ProductTypePicker", () => {
  it("renders plates, piles, steps, marches, bridge piles, and fbs options", () => {
    render(<ProductTypePicker onSelect={vi.fn()} />);

    expect(screen.getByRole("button", { name: /Плиты/i })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /цельные железобетонные сваи/i })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Ступени/i })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /лестничные марши/i })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /мостовые железобетонные сваи/i })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /фундаментные блоки ФБС/i })).toBeInTheDocument();
  });

  it("calls onSelect with steps when the steps card is clicked", () => {
    const onSelect = vi.fn();

    render(<ProductTypePicker onSelect={onSelect} />);
    const stepsButtons = screen.getAllByRole("button", { name: /Ступени/i });
    fireEvent.click(stepsButtons[stepsButtons.length - 1]!);

    expect(onSelect).toHaveBeenCalledWith("steps");
  });

  it("calls onSelect with marches when the marches card is clicked", () => {
    const onSelect = vi.fn();

    render(<ProductTypePicker onSelect={onSelect} />);
    const marchButtons = screen.getAllByRole("button", { name: /лестничные марши/i });
    fireEvent.click(marchButtons[marchButtons.length - 1]!);

    expect(onSelect).toHaveBeenCalledWith("marches");
  });

  it("calls onSelect with bridge_piles when the bridge piles card is clicked", () => {
    const onSelect = vi.fn();

    render(<ProductTypePicker onSelect={onSelect} />);
    fireEvent.click(screen.getByRole("button", { name: /мостовые железобетонные сваи/i }));

    expect(onSelect).toHaveBeenCalledWith("bridge_piles");
  });

  it("calls onSelect with fbs when the FBS card is clicked", () => {
    const onSelect = vi.fn();

    render(<ProductTypePicker onSelect={onSelect} />);
    fireEvent.click(screen.getByRole("button", { name: /фундаментные блоки ФБС/i }));

    expect(onSelect).toHaveBeenCalledWith("fbs");
  });
});
