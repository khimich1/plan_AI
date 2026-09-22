import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { PlanHealthBadge } from "@/features/production/components/PlanHealthBadge";
import type { PlanIntegrityReport } from "@/features/production/types/production";

afterEach(() => {
  cleanup();
});

const clean: PlanIntegrityReport = {
  orphans: 0,
  surplus: 0,
  no_grade: 0,
  items: [],
};

const dirty: PlanIntegrityReport = {
  orphans: 1,
  surplus: 2,
  no_grade: 0,
  items: [
    { kind: "orphan", message: "призрак: item без kp_plate_id" },
    { kind: "surplus", message: "Лишние плиты: 2" },
  ],
};

describe("PlanHealthBadge", () => {
  it("shows clean label when counts are zero", () => {
    render(<PlanHealthBadge report={clean} />);
    expect(screen.getByText("чист")).toBeTruthy();
  });

  it("shows discrepancy count and expands the list", () => {
    render(<PlanHealthBadge report={dirty} />);
    const summary = screen.getByText("3 расхождений");
    expect(summary).toBeTruthy();
    expect(screen.queryByText("призрак: item без kp_plate_id")).toBeNull();
    fireEvent.click(summary);
    expect(screen.getByText("призрак: item без kp_plate_id")).toBeTruthy();
    expect(screen.getByText("Лишние плиты: 2")).toBeTruthy();
  });
});
