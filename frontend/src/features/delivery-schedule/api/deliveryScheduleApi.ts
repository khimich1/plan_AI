import { httpClient } from "@/shared/api/httpClient";
import type {
  DeliveryScheduleDocumentFmt,
  DeliverySchedulePut,
  DeliveryScheduleView,
  ImportDraftResponse,
} from "@/features/delivery-schedule/types/deliverySchedule";

const basePath = (kpId: number) => `/api/v1/commercial/archive/${kpId}/delivery-schedule`;

const JSON_HEADERS = { "Content-Type": "application/json" };

export const deliveryScheduleApi = {
  get: (kpId: number) => httpClient.get<DeliveryScheduleView>(basePath(kpId)),

  put: (kpId: number, payload: DeliverySchedulePut) =>
    httpClient.put<DeliveryScheduleView>(
      basePath(kpId),
      JSON.stringify(payload),
      JSON_HEADERS,
    ),

  importFile: (kpId: number, file: File) => {
    const formData = new FormData();
    formData.append("file", file);
    return httpClient.post<ImportDraftResponse>(`${basePath(kpId)}/import`, formData);
  },

  downloadTemplate: (kpId: number) =>
    httpClient.download(`${basePath(kpId)}/template`, "delivery_schedule_template.xlsx"),

  downloadDocument: (kpId: number, fmt: DeliveryScheduleDocumentFmt) => {
    const fallbackName =
      fmt === "pdf" ? `delivery_schedule_${kpId}.pdf` : `delivery_schedule_${kpId}.xlsx`;
    return httpClient.download(`${basePath(kpId)}/document?fmt=${fmt}`, fallbackName);
  },
};
