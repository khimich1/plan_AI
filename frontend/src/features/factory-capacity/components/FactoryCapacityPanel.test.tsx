import { describe, expect, it } from "vitest";
import { render, screen } from "@testing-library/react";
import {
  FactoryCapacityPanel,
  isCapacityRed,
} from "@/features/factory-capacity/components/FactoryCapacityPanel";
import type { CapacitySnapshot } from "@/features/factory-capacity/types/capacity";

const baseSnap = (overrides: Partial<CapacitySnapshot> = {}): CapacitySnapshot => ({
  start_date: "2026-03-03",
  target_date: "2026-03-20",
  tracks_needed: 4,
  tracks_free_in_window: 20,
  delta: 16,
  status: "green",
  hint: null,
  days_info: {
    "2026-03-03": { occupied: 1, max: 5 },
    "2026-03-04": { occupied: 5, max: 5 },
  },
  holidays: [],
  extra_workdays: [],
  calendar_from_month: "2026-03",
  calendar_to_month: "2026-03",
  ...overrides,
});

describe("FactoryCapacityPanel", () => {
  it("shows needed / free / delta", () => {
    render(<FactoryCapacityPanel snapshot={baseSnap()} />);
    const panel = screen.getByTestId("factory-capacity-panel");
    expect(panel).toHaveTextContent("нужно");
    expect(panel).toHaveTextContent("свободно");
    expect(panel).toHaveTextContent("Δ");
    expect(panel.textContent).toMatch(/нужно\s*4/);
    expect(panel.textContent).toMatch(/свободно\s*20/);
    expect(panel.textContent).toMatch(/Δ\s*16/);
  });

  it("shows hint alert on red", () => {
    render(
      <FactoryCapacityPanel
        snapshot={baseSnap({
          status: "red",
          hint: "нужно +3 дорожек до 20.03.2026",
          delta: -3,
        })}
      />,
    );
    expect(screen.getByText(/нужно \+3 дорожек/)).toBeInTheDocument();
    expect(screen.getByText(/Увеличьте срок/)).toBeInTheDocument();
  });

  it("renders mini calendar", () => {
    const { container } = render(<FactoryCapacityPanel snapshot={baseSnap()} />);
    expect(container.querySelectorAll('[data-testid="factory-mini-calendar"]')).toHaveLength(1);
  });

  it("isCapacityRed only for red", () => {
    expect(isCapacityRed(baseSnap({ status: "red" }))).toBe(true);
    expect(isCapacityRed(baseSnap({ status: "yellow" }))).toBe(false);
    expect(isCapacityRed(null)).toBe(false);
  });
});
