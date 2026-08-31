import { describe, expect, it } from "vitest";
import type { DayInfo, FillTargetItem } from "@/features/production/types/production";
import {
  canAddDayToBasket,
  getBasketKind,
  getDayKind,
} from "@/features/production/lib/basketDayKind";

const day = (occupied: number, max = 5): Pick<DayInfo, "occupied" | "max"> => ({
  occupied,
  max,
});

describe("getDayKind", () => {
  it("returns empty when occupied === 0", () => {
    expect(getDayKind(day(0))).toBe("empty");
    expect(getDayKind(day(0, 8))).toBe("empty");
  });

  it("returns partial when 0 < occupied < max", () => {
    expect(getDayKind(day(1, 5))).toBe("partial");
    expect(getDayKind(day(4, 5))).toBe("partial");
  });

  it("returns full when occupied >= max (freeSlots === 0)", () => {
    expect(getDayKind(day(5, 5))).toBe("full");
    expect(getDayKind(day(6, 5))).toBe("full");
  });
});

describe("getBasketKind", () => {
  const daysInfo: Record<string, DayInfo> = {
    "2026-07-23": {
      occupied: 0,
      max: 5,
      completed: false,
      day_number: 1,
    },
    "2026-07-24": {
      occupied: 2,
      max: 5,
      completed: false,
      day_number: 2,
    },
    "2026-07-25": {
      occupied: 5,
      max: 5,
      completed: false,
      day_number: 3,
    },
  };

  it("returns null for empty basket", () => {
    expect(getBasketKind([], daysInfo)).toBeNull();
  });

  it("returns empty when basket has empty days", () => {
    const items: FillTargetItem[] = [{ date: "2026-07-23", tracks: 3 }];
    expect(getBasketKind(items, daysInfo)).toBe("empty");
  });

  it("returns partial when basket has partial days", () => {
    const items: FillTargetItem[] = [{ date: "2026-07-24", tracks: 2 }];
    expect(getBasketKind(items, daysInfo)).toBe("partial");
  });

  it("uses first item kind (basket should be homogeneous)", () => {
    const items: FillTargetItem[] = [
      { date: "2026-07-23", tracks: 3 },
      { date: "2026-07-24", tracks: 2 },
    ];
    expect(getBasketKind(items, daysInfo)).toBe("empty");
  });

  it("treats missing daysInfo entry as empty (free day outside plan range)", () => {
    const items: FillTargetItem[] = [{ date: "2099-01-01", tracks: 1 }];
    expect(getBasketKind(items, daysInfo)).toBe("empty");
  });
});

describe("canAddDayToBasket", () => {
  it("allows any non-full day into empty basket", () => {
    expect(canAddDayToBasket(null, "empty")).toBe(true);
    expect(canAddDayToBasket(null, "partial")).toBe(true);
  });

  it("rejects full days always", () => {
    expect(canAddDayToBasket(null, "full")).toBe(false);
    expect(canAddDayToBasket("empty", "full")).toBe(false);
    expect(canAddDayToBasket("partial", "full")).toBe(false);
  });

  it("allows same kind", () => {
    expect(canAddDayToBasket("empty", "empty")).toBe(true);
    expect(canAddDayToBasket("partial", "partial")).toBe(true);
  });

  it("rejects mixing empty and partial", () => {
    expect(canAddDayToBasket("empty", "partial")).toBe(false);
    expect(canAddDayToBasket("partial", "empty")).toBe(false);
  });
});
