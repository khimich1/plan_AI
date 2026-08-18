import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { CapacityDeficit } from "@/features/production/types/production";
import { CapacityDeficitAlert } from "./CapacityDeficitAlert";

const baseDeficit = (overrides: Partial<CapacityDeficit> = {}): CapacityDeficit => ({
  tracks_needed: 10,
  tracks_available: 5,
  tracks_missing: 2,
  deficit_until: "2026-08-20",
  options: [
    { action: "bump_fill", date: "2026-08-14", add_tracks: 2, free: 5 },
    { action: "propose_day", date: "2026-08-12", add_tracks: 3, free: 3 },
  ],
  ...overrides,
});

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

describe("CapacityDeficitAlert", () => {
  it("renders nothing when deficit is null", () => {
    const { container } = render(
      <CapacityDeficitAlert deficit={null} onApplyOption={vi.fn()} />,
    );
    expect(container).toBeEmptyDOMElement();
  });

  it("renders nothing when tracks_missing is 0", () => {
    const { container } = render(
      <CapacityDeficitAlert
        deficit={baseDeficit({ tracks_missing: 0 })}
        onApplyOption={vi.fn()}
      />,
    );
    expect(container).toBeEmptyDOMElement();
  });

  it("shows deficit numbers and option buttons", () => {
    render(<CapacityDeficitAlert deficit={baseDeficit()} onApplyOption={vi.fn()} />);
    expect(screen.getByText(/Нужно дорожек: 10/)).toBeInTheDocument();
    expect(screen.getByText(/Не хватает: 2/)).toBeInTheDocument();
    expect(
      screen.getByRole("button", { name: /дозаполнить выбранный день/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByRole("button", { name: /добавить день/i }),
    ).toBeInTheDocument();
  });

  it("applies bump_fill option without confirm", () => {
    const onApplyOption = vi.fn();
    render(
      <CapacityDeficitAlert deficit={baseDeficit()} onApplyOption={onApplyOption} />,
    );
    fireEvent.click(
      screen.getByRole("button", { name: /дозаполнить выбранный день/i }),
    );
    expect(onApplyOption).toHaveBeenCalledWith({
      action: "bump_fill",
      date: "2026-08-14",
      add_tracks: 2,
      free: 5,
    });
  });

  it("applies propose_day only after confirm", () => {
    const onApplyOption = vi.fn();
    vi.spyOn(window, "confirm").mockReturnValue(true);
    render(
      <CapacityDeficitAlert deficit={baseDeficit()} onApplyOption={onApplyOption} />,
    );
    fireEvent.click(screen.getByRole("button", { name: /добавить день/i }));
    expect(window.confirm).toHaveBeenCalled();
    expect(onApplyOption).toHaveBeenCalledWith({
      action: "propose_day",
      date: "2026-08-12",
      add_tracks: 3,
      free: 3,
    });
  });

  it("does not apply propose_day when confirm cancelled", () => {
    const onApplyOption = vi.fn();
    vi.spyOn(window, "confirm").mockReturnValue(false);
    render(
      <CapacityDeficitAlert deficit={baseDeficit()} onApplyOption={onApplyOption} />,
    );
    fireEvent.click(screen.getByRole("button", { name: /добавить день/i }));
    expect(onApplyOption).not.toHaveBeenCalled();
  });
});
