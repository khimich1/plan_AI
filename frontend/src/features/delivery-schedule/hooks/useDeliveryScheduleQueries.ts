import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { deliveryScheduleApi } from "@/features/delivery-schedule/api/deliveryScheduleApi";
import type {
  DeliveryScheduleDocumentFmt,
  DeliverySchedulePut,
  DeliveryScheduleView,
} from "@/features/delivery-schedule/types/deliverySchedule";
import { ApiError } from "@/shared/lib/apiError";
import { saveBlobAs } from "@/shared/lib/downloadFile";

export const deliveryScheduleKeys = {
  all: ["delivery-schedule"] as const,
  detail: (kpId: number) => ["delivery-schedule", "detail", kpId] as const,
};

/** GET: 404 → `null` (графика ещё нет), без ретраев. */
export const useDeliveryScheduleQuery = (kpId: number | null) =>
  useQuery<DeliveryScheduleView | null>({
    queryKey: deliveryScheduleKeys.detail(kpId ?? -1),
    queryFn: async () => {
      try {
        return await deliveryScheduleApi.get(kpId as number);
      } catch (error) {
        if (error instanceof ApiError && error.status === 404) {
          return null;
        }
        throw error;
      }
    },
    enabled: kpId !== null,
    retry: false,
  });

export const usePutDeliveryScheduleMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ kpId, payload }: { kpId: number; payload: DeliverySchedulePut }) =>
      deliveryScheduleApi.put(kpId, payload),
    onSuccess: (view) => {
      queryClient.setQueryData(deliveryScheduleKeys.detail(view.kp_id), view);
      queryClient.invalidateQueries({ queryKey: deliveryScheduleKeys.detail(view.kp_id) });
    },
  });
};

export const useImportDeliveryScheduleMutation = () =>
  useMutation({
    mutationFn: ({ kpId, file }: { kpId: number; file: File }) =>
      deliveryScheduleApi.importFile(kpId, file),
  });

/** GET /template — только вне открытого редактора. Кнопка в модалке собирает файл на клиенте. */
export const useDownloadDeliveryScheduleTemplateMutation = () =>
  useMutation({
    mutationKey: ["delivery-schedule", "template"],
    mutationFn: async (kpId: number) => {
      const result = await deliveryScheduleApi.downloadTemplate(kpId);
      saveBlobAs(result.blob, result.filename);
      return result;
    },
  });

export const useDownloadDeliveryScheduleDocumentMutation = () =>
  useMutation({
    mutationKey: ["delivery-schedule", "document"],
    mutationFn: async ({
      kpId,
      fmt,
    }: {
      kpId: number;
      fmt: DeliveryScheduleDocumentFmt;
    }) => {
      const result = await deliveryScheduleApi.downloadDocument(kpId, fmt);
      saveBlobAs(result.blob, result.filename);
      return result;
    },
  });
