import { useQuery } from "@tanstack/react-query";
import { capacityApi } from "@/features/factory-capacity/api/capacityApi";
import type { CapacitySnapshot } from "@/features/factory-capacity/types/capacity";

export const capacityKeys = {
  all: ["factory-capacity"] as const,
  snapshot: (kpId: number, target: string) =>
    ["factory-capacity", "snapshot", kpId, target] as const,
};

/** GET capacity-snapshot; не дергается при kpId=null или пустом target. */
export const useCapacitySnapshotQuery = (
  kpId: number | null,
  targetIso: string | null,
  options?: { enabled?: boolean },
) =>
  useQuery<CapacitySnapshot>({
    queryKey: capacityKeys.snapshot(kpId ?? -1, targetIso ?? ""),
    queryFn: () => capacityApi.getSnapshot(kpId as number, targetIso as string),
    enabled:
      kpId !== null &&
      Boolean(targetIso) &&
      (options?.enabled ?? true),
    staleTime: 10_000,
  });
