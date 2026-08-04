import { describe, expect, it } from "vitest";
import type { DayInfo } from "@/features/production/types/production";
import {
  datesBetweenInclusive,
  isDayBrushSelectable,
  paintDays,
} from "@/features/production/lib/calendarRange";

const day = (
  occupied: number,
  max = 5,
  completed = false,
): DayInfo => ({
  occupied,
  max,
  completed,
  day_number: 1,
});

describe("datesBetweenInclusive", () => {
  it("returns sorted inclusive ISO range (a < b)", () => {
    expect(datesBetweenInclusive("2026-07-20", "2026-07-24")).toEqual([
      "2026-07-20",
      "2026-07-21",
      "2026-07-22",
      "2026-07-23",
      "2026-07-24",
    ]);
  });

  it("returns sorted inclusive ISO range when a > b", () => {
    expect(datesBetweenInclusive("2026-07-24", "2026-07-20")).toEqual([
      "2026-07-20",
      "2026-07-21",
      "2026-07-22",
      "2026-07-23",
      "2026-07-24",
    ]);
  });

  it("returns single day when a === b", () => {
    expect(datesBetweenInclusive("2026-07-20", "2026-07-20")).toEqual([
      "2026-07-20",
    ]);
  });

  it("crosses month boundary", () => {
    expect(datesBetweenInclusive("2026-07-30", "2026-08-02")).toEqual([
      "2026-07-30",
      "2026-07-31",
      "2026-08-01",
      "2026-08-02",
    ]);
  });
});

describe("isDayBrushSelectable", () => {
  const emptyHolidays = new Set<string>();
  const emptyExtra = new Set<string>();

  it("returns false for weekend (Sat/Sun) without extra workday", () => {
    // 2026-07-25 = Saturday
    expect(
      isDayBrushSelectable(day(0), {
        iso: "2026-07-25",
        holidays: emptyHolidays,
        extraWorkdays: emptyExtra,
      }),
    ).toBe(false);
  });

  it("returns true for weekend marked as extra workday", () => {
    expect(
      isDayBrushSelectable(day(0), {
        iso: "2026-07-25",
        holidays: emptyHolidays,
        extraWorkdays: new Set(["2026-07-25"]),
      }),
    ).toBe(true);
  });

  it("returns false for holiday", () => {
    expect(
      isDayBrushSelectable(day(0), {
        iso: "2026-07-23",
        holidays: new Set(["2026-07-23"]),
        extraWorkdays: emptyExtra,
      }),
    ).toBe(false);
  });

  it("returns false for completed", () => {
    expect(
      isDayBrushSelectable(day(0, 5, true), {
        iso: "2026-07-23",
        holidays: emptyHolidays,
        extraWorkdays: emptyExtra,
      }),
    ).toBe(false);
  });

  it("returns false for full (freeSlots === 0)", () => {
    expect(
      isDayBrushSelectable(day(5, 5), {
        iso: "2026-07-23",
        holidays: emptyHolidays,
        extraWorkdays: emptyExtra,
      }),
    ).toBe(false);
  });

  it("returns true for empty workday", () => {
    expect(
      isDayBrushSelectable(day(0), {
        iso: "2026-07-23",
        holidays: emptyHolidays,
        extraWorkdays: emptyExtra,
      }),
    ).toBe(true);
  });

  it("returns true for partial workday with free slots", () => {
    expect(
      isDayBrushSelectable(day(2, 5), {
        iso: "2026-07-23",
        holidays: emptyHolidays,
        extraWorkdays: emptyExtra,
      }),
    ).toBe(true);
  });

  it("treats missing dayInfo as empty with defaultMax", () => {
    expect(
      isDayBrushSelectable(undefined, {
        iso: "2026-07-23",
        holidays: emptyHolidays,
        extraWorkdays: emptyExtra,
        defaultMax: 5,
      }),
    ).toBe(true);
  });
});

describe("paintDays", () => {
  const holidays = new Set<string>();
  const extraWorkdays = new Set<string>();

  const daysInfo: Record<string, DayInfo> = {
    "2026-07-20": day(0), // Mon empty
    "2026-07-21": day(0), // Tue empty
    "2026-07-22": day(2), // Wed partial free=3
    "2026-07-23": day(5), // Thu full
    "2026-07-24": day(0, 5, true), // Fri completed
  };

  it("clamps tracks to freeSlots", () => {
    const result = paintDays({
      dates: ["2026-07-22"],
      brushTracks: 5,
      daysInfo,
      basketKind: null,
      holidays,
      extraWorkdays,
    });
    expect(result.error).toBeNull();
    expect(result.added).toEqual([{ date: "2026-07-22", tracks: 3 }]);
  });

  it("paints empty days with brushTracks", () => {
    const result = paintDays({
      dates: ["2026-07-20", "2026-07-21"],
      brushTracks: 3,
      daysInfo,
      basketKind: null,
      holidays,
      extraWorkdays,
    });
    expect(result.error).toBeNull();
    expect(result.added).toEqual([
      { date: "2026-07-20", tracks: 3 },
      { date: "2026-07-21", tracks: 3 },
    ]);
  });

  it("skips full/completed/weekend and sets error", () => {
    // 2026-07-25 = Saturday
    const result = paintDays({
      dates: ["2026-07-20", "2026-07-23", "2026-07-24", "2026-07-25"],
      brushTracks: 2,
      daysInfo,
      basketKind: null,
      holidays,
      extraWorkdays,
    });
    expect(result.added).toEqual([{ date: "2026-07-20", tracks: 2 }]);
    expect(result.error).not.toBeNull();
  });

  it("skips incompatible kind mid-range and sets error (keeps compatible)", () => {
    const result = paintDays({
      dates: ["2026-07-20", "2026-07-22"],
      brushTracks: 2,
      daysInfo,
      basketKind: null,
      holidays,
      extraWorkdays,
    });
    // First empty sets kind; partial is skipped
    expect(result.added).toEqual([{ date: "2026-07-20", tracks: 2 }]);
    expect(result.error).toMatch(/смешивать/i);
  });

  it("rejects partial when basketKind is empty", () => {
    const result = paintDays({
      dates: ["2026-07-22"],
      brushTracks: 2,
      daysInfo,
      basketKind: "empty",
      holidays,
      extraWorkdays,
    });
    expect(result.added).toEqual([]);
    expect(result.error).toMatch(/смешивать/i);
  });

  it("allows partial when basketKind is partial", () => {
    const result = paintDays({
      dates: ["2026-07-22"],
      brushTracks: 2,
      daysInfo,
      basketKind: "partial",
      holidays,
      extraWorkdays,
    });
    expect(result.error).toBeNull();
    expect(result.added).toEqual([{ date: "2026-07-22", tracks: 2 }]);
  });

  it("uses defaultMax for missing daysInfo (empty)", () => {
    const result = paintDays({
      dates: ["2099-01-05"], // Monday
      brushTracks: 4,
      daysInfo: {},
      basketKind: null,
      holidays,
      extraWorkdays,
      defaultMax: 5,
    });
    expect(result.error).toBeNull();
    expect(result.added).toEqual([{ date: "2099-01-05", tracks: 4 }]);
  });
});
