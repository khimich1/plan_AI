import { describe, expect, it } from "vitest";
import { planNameFromDates } from "@/features/production/lib/planNameFromDates";

describe("planNameFromDates", () => {
  it("formats a single day as «План DD.MM»", () => {
    expect(planNameFromDates(["2026-07-23"])).toBe("План 23.07");
  });

  it("formats a range as «План DD–DD.MM» from first to last sorted date", () => {
    expect(planNameFromDates(["2026-07-23", "2026-07-25"])).toBe(
      "План 23–25.07",
    );
  });

  it("sorts dates before formatting", () => {
    expect(planNameFromDates(["2026-07-25", "2026-07-23"])).toBe(
      "План 23–25.07",
    );
  });

  it("uses end-month for cross-month ranges", () => {
    expect(planNameFromDates(["2026-07-30", "2026-08-02"])).toBe(
      "План 30–02.08",
    );
  });

  it("returns empty string for empty input", () => {
    expect(planNameFromDates([])).toBe("");
  });
});
