import { describe, expect, it } from "vitest";
import type { KpCandidateItem } from "@/features/production/types/production";
import {
  allPlatesLengthM,
  clampTracksPerDay,
  estimateFromLengthM,
  estimateKpSelection,
  selectedLengthM,
} from "@/features/production/lib/productionEstimate";

const makeKp = (plates: KpCandidateItem["plates"]): KpCandidateItem => ({
  kp_id: 4,
  customer_name: "TKM",
  creation_date: "",
  execution_terms: "",
  total_plates: plates.length,
  completed_plates: 0,
  completion_pct: 0,
  in_plan_pct: 0,
  total_length_m: plates.reduce((s, p) => s + p.length_m * p.qty, 0),
  plates,
});

describe("estimateFromLengthM", () => {
  it("38.3 m and 1 track/day → 1 track, 1 day (ceil, no false extra day)", () => {
    const r = estimateFromLengthM(38.3, 1);
    expect(r.estimated_tracks).toBe(1);
    expect(r.estimated_days).toBe(1);
    expect(r.total_length_m).toBe(38.3);
  });

  it("200 m and 5 tracks/day → 2 tracks, 1 day (no false extra day)", () => {
    const r = estimateFromLengthM(200, 5);
    expect(r.estimated_tracks).toBe(2);
    expect(r.estimated_days).toBe(1);
  });

  it("clamps invalid tracksPerDay to 1", () => {
    expect(estimateFromLengthM(200, 0).estimated_tracks).toBe(2);
    expect(estimateFromLengthM(200, 0).estimated_days).toBe(2);
    expect(estimateFromLengthM(200, -3).estimated_days).toBe(2);
  });
});

describe("clampTracksPerDay", () => {
  it("returns at least 1", () => {
    expect(clampTracksPerDay(0)).toBe(1);
    expect(clampTracksPerDay(5)).toBe(5);
  });
});

describe("selectedLengthM", () => {
  it("sums only selected plates with partial qty", () => {
    const kp = makeKp([
      {
        id: 1,
        plate_name: "A",
        length_m: 2.4,
        width_m: 1.2,
        load_class: 1000,
        qty: 5,
      },
      {
        id: 2,
        plate_name: "B",
        length_m: 6.0,
        width_m: 1.2,
        load_class: 1000,
        qty: 3,
      },
    ]);
    expect(selectedLengthM(kp, [1], { 1: 2 })).toBeCloseTo(4.8);
    expect(allPlatesLengthM(kp)).toBeCloseTo(2.4 * 5 + 6.0 * 3);
  });
});

describe("estimateKpSelection", () => {
  it("returns null when nothing selected", () => {
    const kp = makeKp([]);
    expect(estimateKpSelection(kp, [], {}, 1)).toBeNull();
  });
});
