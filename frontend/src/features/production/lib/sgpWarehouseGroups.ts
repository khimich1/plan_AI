import type { SgpPlateItem, SgpProgress } from "@/features/production/api/sgpApi";

export const UNLINKED_GROUP_KEY = "unlinked" as const;

export type SgpGroupKey = number | typeof UNLINKED_GROUP_KEY;

export type SgpPlateGroup = {
  key: SgpGroupKey;
  kpId: number | null;
  plates: SgpPlateItem[];
  customerName: string | null;
  executionTerms: string | null;
  sgpProgress: SgpProgress | null;
  positionCount: number;
  totalQty: number;
};

export function groupKeyForPlate(plate: SgpPlateItem): SgpGroupKey {
  return plate.kp_id ?? UNLINKED_GROUP_KEY;
}

export function groupPlatesByKp(items: SgpPlateItem[]): SgpPlateGroup[] {
  const byKey = new Map<SgpGroupKey, SgpPlateItem[]>();

  for (const plate of items) {
    const key = groupKeyForPlate(plate);
    const bucket = byKey.get(key);
    if (bucket) {
      bucket.push(plate);
    } else {
      byKey.set(key, [plate]);
    }
  }

  const groups: SgpPlateGroup[] = [];

  for (const [key, plates] of byKey) {
    const first = plates[0];
    groups.push({
      key,
      kpId: key === UNLINKED_GROUP_KEY ? null : key,
      plates,
      customerName: first.customer_name,
      executionTerms: first.execution_terms,
      sgpProgress: first.sgp_progress,
      positionCount: plates.length,
      totalQty: plates.reduce((sum, p) => sum + p.qty, 0),
    });
  }

  groups.sort((a, b) => {
    if (a.key === UNLINKED_GROUP_KEY) return -1;
    if (b.key === UNLINKED_GROUP_KEY) return 1;
    return (a.kpId ?? 0) - (b.kpId ?? 0);
  });

  return groups;
}

export function allGroupKeys(groups: SgpPlateGroup[]): SgpGroupKey[] {
  return groups.map((g) => g.key);
}

export function groupLabel(group: SgpPlateGroup): string {
  return group.kpId != null ? `#${group.kpId}` : "Без КП";
}
