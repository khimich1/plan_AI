import { describe, expect, it } from "vitest";
import {
  formatFrostPair,
  isOffTablePair,
  readConcreteSpec,
  tablePairsForGrade,
} from "@/features/commercial-offer/lib/concreteSpec";

describe("tablePairsForGrade", () => {
  it("offers granite and ordinary pairs for B25", () => {
    expect(tablePairsForGrade("B25")).toEqual([
      { aggregate: "granite", frost_resistance: "F200", waterproofness: "W8", label: "F200 · W8" },
      { aggregate: "ordinary", frost_resistance: "F200", waterproofness: "W6", label: "F200 · W6" },
    ]);
  });

  it("offers a single pair when ordinary waterproofness matches granite", () => {
    expect(tablePairsForGrade("B15").map((pair) => pair.label)).toEqual(["F100 · W4"]);
  });

  it("offers a single granite pair for B30 and plate marks", () => {
    expect(tablePairsForGrade("B30_granite").map((pair) => pair.label)).toEqual(["F300 · W10"]);
    expect(tablePairsForGrade("М500").map((pair) => pair.label)).toEqual(["F300 · W12"]);
    expect(tablePairsForGrade("M400").map((pair) => pair.label)).toEqual(["F300 · W10"]);
  });

  it("returns no pairs for an unknown grade", () => {
    expect(tablePairsForGrade("")).toEqual([]);
    expect(tablePairsForGrade("B99")).toEqual([]);
  });
});

describe("formatFrostPair", () => {
  it("prints the saved pair and a dash for an empty snapshot", () => {
    expect(formatFrostPair("F200", "W8")).toBe("F200 · W8");
    expect(formatFrostPair(null, null)).toBe("—");
    expect(formatFrostPair("", "")).toBe("—");
  });
});

describe("isOffTablePair", () => {
  it("marks a manual pair that is not a ready pair of the current grade", () => {
    expect(isOffTablePair("B25", "F150", "W4")).toBe(true);
    expect(isOffTablePair("B25", "F200", "W8")).toBe(false);
    expect(isOffTablePair("B25", null, null)).toBe(false);
  });
});

describe("readConcreteSpec", () => {
  it("keeps a manual empty aggregate and drops a missing snapshot to nulls", () => {
    expect(
      readConcreteSpec({
        frost_resistance: "F150",
        waterproofness: "W4",
        concrete_aggregate: "",
        concrete_spec_source: "manual",
      }),
    ).toEqual({
      frost_resistance: "F150",
      waterproofness: "W4",
      concrete_aggregate: "",
      concrete_spec_source: "manual",
    });
    expect(readConcreteSpec({})).toEqual({
      frost_resistance: null,
      waterproofness: null,
      concrete_aggregate: null,
      concrete_spec_source: null,
    });
  });
});
