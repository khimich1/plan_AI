import { describe, expect, it } from "vitest";

import { liveWidePlateLines } from "@/features/commercial-offer/lib/liveWidePlateLines";

describe("liveWidePlateLines", () => {
  it("treats 44-15-10п 5 as wide with qty 5", () => {
    const lines = liveWidePlateLines("44-15-10п 5");
    expect(lines).toHaveLength(1);
    expect(lines[0]?.qty).toBe(5);
    expect(lines[0]?.line).toBe("44-15-10п 5");
    expect(lines[0]?.id).toBe("live-wide-0");
  });

  it("after replacing text with 34-15-10п 15 reports wide qty 15", () => {
    expect(liveWidePlateLines("44-15-10п 5")[0]?.qty).toBe(5);
    const next = liveWidePlateLines("34-15-10п 15");
    expect(next).toHaveLength(1);
    expect(next[0]?.qty).toBe(15);
    expect(next[0]?.line).toBe("34-15-10п 15");
  });

  it("does not treat 34-12-10п 15 as wide", () => {
    expect(liveWidePlateLines("34-12-10п 15")).toEqual([]);
  });

  it("treats ПБ 34-15-10п 15 as wide", () => {
    const lines = liveWidePlateLines("ПБ 34-15-10п 15");
    expect(lines).toHaveLength(1);
    expect(lines[0]?.qty).toBe(15);
    expect(lines[0]?.line).toBe("ПБ 34-15-10п 15");
  });
});
