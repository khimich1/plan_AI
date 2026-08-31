import { httpClient } from "@/shared/api/httpClient";
import type { CapacitySnapshot } from "@/features/factory-capacity/types/capacity";

const BASE = "/api/v1/commercial/archive";

export const capacityApi = {
  getSnapshot: (kpId: number, target: string) => {
    const params = new URLSearchParams({ target });
    return httpClient.get<CapacitySnapshot>(
      `${BASE}/${kpId}/capacity-snapshot?${params.toString()}`,
    );
  },
};
