import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { logisticsApi } from "@/features/logistics/api/logisticsApi";
import type {
  CreateShipmentPayload,
  ShipmentDetails,
  ShipmentFilters,
  ShipmentItemInput,
  ShipmentRegistryRow,
  UpdateShipmentPayload,
  VehicleClass,
} from "@/features/logistics/types/logistics";
import { archiveKeys } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { productionKeys } from "@/features/production/hooks/useProductionQueries";
import { sgpKeys } from "@/features/production/hooks/useSgpQueries";
import { saveBlobAs } from "@/shared/lib/downloadFile";

export const logisticsKeys = {
  all: ["logistics"] as const,
  shipments: (filters: ShipmentFilters) => ["logistics", "shipments", filters] as const,
  shipment: (id: number) => ["logistics", "shipment", id] as const,
  carriers: (q: string, activeOnly: boolean) =>
    ["logistics", "carriers", q, activeOnly] as const,
  pileCatalog: (q: string) => ["logistics", "pile-catalog", q] as const,
};

/** Списание СГП/прогресс КП меняют смежные разделы — инвалидируем их вместе с логистикой. */
const invalidateRelated = (qc: ReturnType<typeof useQueryClient>) => {
  void qc.invalidateQueries({ queryKey: logisticsKeys.all });
  void qc.invalidateQueries({ queryKey: archiveKeys.all });
  void qc.invalidateQueries({ queryKey: productionKeys.all });
  void qc.invalidateQueries({ queryKey: sgpKeys.all });
};

export const useShipmentsQuery = (filters: ShipmentFilters) =>
  useQuery<ShipmentRegistryRow[]>({
    queryKey: logisticsKeys.shipments(filters),
    queryFn: () => logisticsApi.listShipments(filters),
    staleTime: 10_000,
  });

export const useShipmentQuery = (id: number | null) =>
  useQuery<ShipmentDetails>({
    queryKey: logisticsKeys.shipment(id ?? -1),
    queryFn: () => logisticsApi.getShipment(id as number),
    enabled: id !== null,
  });

export const useCreateShipmentMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (payload: CreateShipmentPayload) => logisticsApi.createShipment(payload),
    onSuccess: () => {
      void qc.invalidateQueries({ queryKey: logisticsKeys.all });
    },
  });
};

export const useReuseTransportMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({
      sourceId,
      payload,
    }: {
      sourceId: number;
      payload: CreateShipmentPayload;
    }) => logisticsApi.reuseTransport(sourceId, payload),
    onSuccess: () => {
      void qc.invalidateQueries({ queryKey: logisticsKeys.all });
    },
  });
};

export const useUpdateShipmentMutation = (id: number) => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (payload: UpdateShipmentPayload) => logisticsApi.updateShipment(id, payload),
    onSuccess: (updated) => {
      qc.setQueryData(logisticsKeys.shipment(id), updated);
      void qc.invalidateQueries({ queryKey: logisticsKeys.all });
    },
  });
};

export const useProposeMutation = (id: number) =>
  useMutation({
    mutationFn: (vehicleClass?: VehicleClass | null) =>
      logisticsApi.proposeItems(id, vehicleClass),
  });

export const useConfirmItemsMutation = (id: number) => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (items: ShipmentItemInput[]) => logisticsApi.confirmItems(id, items),
    onSuccess: () => {
      void qc.invalidateQueries({ queryKey: logisticsKeys.shipment(id) });
      void qc.invalidateQueries({ queryKey: sgpKeys.all });
    },
  });
};

export const useCompleteShipmentMutation = (id: number) => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: () => logisticsApi.completeShipment(id),
    onSuccess: () => invalidateRelated(qc),
  });
};

export const useCancelShipmentMutation = (id: number) => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: () => logisticsApi.cancelShipment(id),
    onSuccess: () => invalidateRelated(qc),
  });
};

export const useShipmentSheetMutation = (id: number) =>
  useMutation({
    mutationKey: ["logistics", "sheet", id],
    mutationFn: async () => {
      const result = await logisticsApi.downloadSheet(id);
      saveBlobAs(result.blob, result.filename);
      return result;
    },
  });

export const useCarriersQuery = (params?: { q?: string; activeOnly?: boolean }) => {
  const q = params?.q ?? "";
  const activeOnly = params?.activeOnly ?? false;
  return useQuery({
    queryKey: logisticsKeys.carriers(q, activeOnly),
    queryFn: () => logisticsApi.listCarriers({ q, activeOnly }),
    staleTime: 15_000,
  });
};

export const useMergeCarrierMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ id, intoId }: { id: number; intoId: number }) =>
      logisticsApi.mergeCarrier(id, intoId),
    onSuccess: () => {
      void qc.invalidateQueries({ queryKey: logisticsKeys.all });
    },
  });
};

export const usePileCatalogQuery = (q = "", enabled = true) =>
  useQuery({
    queryKey: logisticsKeys.pileCatalog(q),
    queryFn: () => logisticsApi.searchPileCatalog(q),
    enabled,
    staleTime: 60_000,
  });
