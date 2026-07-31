/** Match free SGP plates to KP demand plates (strict identity). */

import type { SgpFreePlateItem } from "@/features/production/api/sgpApi";
import type { KpCandidatePlateItem } from "@/features/production/types/production";

const nearlyEqual = (a: number | null | undefined, b: number | null | undefined) => {
  if (a == null || b == null) return a == null && b == null;
  return Math.abs(Number(a) - Number(b)) < 0.005;
};

export const plateMatchesFree = (
  plate: KpCandidatePlateItem,
  free: SgpFreePlateItem,
): boolean =>
  plate.plate_name === free.plate_name &&
  nearlyEqual(plate.length_m, free.length_m) &&
  nearlyEqual(plate.width_m, free.width_m) &&
  Number(plate.load_class ?? 0) === Number(free.load_class ?? 0);

export const freeQtyForPlate = (
  plate: KpCandidatePlateItem,
  freeItems: SgpFreePlateItem[],
): number =>
  freeItems
    .filter((f) => plateMatchesFree(plate, f))
    .reduce((sum, f) => sum + f.qty, 0);

export const pickFreeReservations = (
  plate: KpCandidatePlateItem,
  freeItems: SgpFreePlateItem[],
  qty: number,
): { sgp_id: number; qty: number }[] => {
  let remaining = qty;
  const out: { sgp_id: number; qty: number }[] = [];
  for (const free of freeItems) {
    if (remaining <= 0) break;
    if (!plateMatchesFree(plate, free)) continue;
    const take = Math.min(remaining, free.qty);
    if (take <= 0) continue;
    out.push({ sgp_id: free.id, qty: take });
    remaining -= take;
  }
  return out;
};
