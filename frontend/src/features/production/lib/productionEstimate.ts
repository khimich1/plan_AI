import type { KpCandidateItem } from "@/features/production/types/production";

export const TRACK_LENGTH_M = 101;

export type ProductionEstimate = {
  total_length_m: number;
  estimated_tracks: number;
  estimated_days: number;
};

export function clampTracksPerDay(tracksPerDay: number): number {
  if (!Number.isFinite(tracksPerDay) || tracksPerDay < 1) {
    return 1;
  }
  return Math.floor(tracksPerDay);
}

export function estimateFromLengthM(
  totalLengthM: number,
  tracksPerDay: number,
): ProductionEstimate {
  const length = Math.max(0, totalLengthM);
  const perDay = clampTracksPerDay(tracksPerDay);
  const estimated_tracks = Math.max(1, Math.ceil(length / TRACK_LENGTH_M));
  const estimated_days = Math.max(1, Math.ceil(estimated_tracks / perDay));
  return {
    total_length_m: length,
    estimated_tracks,
    estimated_days,
  };
}

export function allPlatesLengthM(kp: KpCandidateItem): number {
  return kp.plates.reduce((sum, plate) => sum + plate.length_m * plate.qty, 0);
}

export function selectedLengthM(
  kp: KpCandidateItem,
  selectedIds: number[],
  qtyByPlate: Record<number, number>,
): number {
  return selectedIds.reduce((sum, id) => {
    const plate = kp.plates.find((p) => p.id === id);
    if (!plate) {
      return sum;
    }
    const qty = qtyByPlate[id] ?? plate.qty;
    return sum + plate.length_m * qty;
  }, 0);
}

export function estimateKpSelection(
  kp: KpCandidateItem,
  selectedIds: number[],
  qtyByPlate: Record<number, number>,
  tracksPerDay: number,
): ProductionEstimate | null {
  if (selectedIds.length === 0) {
    return null;
  }
  const length = selectedLengthM(kp, selectedIds, qtyByPlate);
  return estimateFromLengthM(length, tracksPerDay);
}
