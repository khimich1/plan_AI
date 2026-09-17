import { httpClient } from "@/shared/api/httpClient";
import type { PriceDeskPreview, PriceDeskStatus } from "@/features/price-desk/types/priceDesk";

const BASE = "/api/v1/prices";

export const priceDeskApi = {
  status: () => httpClient.get<PriceDeskStatus>(`${BASE}/status`),

  preview: (file: File) => {
    const formData = new FormData();
    formData.append("file", file);
    return httpClient.post<PriceDeskPreview>(`${BASE}/preview`, formData);
  },

  apply: (file: File, fileSha256: string) => {
    const formData = new FormData();
    formData.append("file", file);
    formData.append("file_sha256", fileSha256);
    return httpClient.post<PriceDeskPreview>(`${BASE}/apply`, formData);
  },
};
