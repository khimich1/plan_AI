import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { gsmApi } from "@/features/gsm/api/gsmApi";
import type {
  CardCreatePayload,
  CardPatchPayload,
  DriverCreatePayload,
  DriverPatchPayload,
  GsmSeasonMode,
  GsmSettings,
  StationCreatePayload,
  StationPatchPayload,
  VehicleCreatePayload,
  VehiclePatchPayload,
  WaybillCreatePayload,
  WaybillExportPayload,
  WaybillGeneratePayload,
  WaybillListParams,
  WaybillPatchPayload,
  OverviewParams,
  TransactionListParams,
  UsageReportPayload,
  WaybillBulkGeneratePayload,
} from "@/features/gsm/types/gsm";
import { saveBlobAs } from "@/shared/lib/downloadFile";

export const gsmKeys = {
  all: ["gsm"] as const,
  vehicles: (activeOnly: boolean) => ["gsm", "vehicles", activeOnly] as const,
  drivers: (activeOnly: boolean) => ["gsm", "drivers", activeOnly] as const,
  cards: (includeArchived: boolean) => ["gsm", "cards", includeArchived] as const,
  stations: ["gsm", "stations"] as const,
  settings: ["gsm", "settings"] as const,
  waybills: (params: WaybillListParams | null) =>
    params
      ? (["gsm", "waybills", params.vehicleId ?? "all", params.periodFrom, params.periodTo] as const)
      : (["gsm", "waybills"] as const),
  routes: (vehicleId: number | null) =>
    vehicleId != null ? (["gsm", "routes", vehicleId] as const) : (["gsm", "routes"] as const),
  overview: (params: OverviewParams | null) =>
    params
      ? (["gsm", "overview", params.periodFrom, params.periodTo] as const)
      : (["gsm", "overview"] as const),
  transactions: (params: TransactionListParams | null) =>
    params
      ? ([
          "gsm",
          "transactions",
          params.periodFrom,
          params.periodTo,
          params.vehicleId ?? "all",
          params.serviceType ?? "all",
        ] as const)
      : (["gsm", "transactions"] as const),
};

const invalidateGsm = (qc: ReturnType<typeof useQueryClient>) => {
  void qc.invalidateQueries({ queryKey: gsmKeys.all });
};

export const useGsmVehiclesQuery = (activeOnly = true) =>
  useQuery({
    queryKey: gsmKeys.vehicles(activeOnly),
    queryFn: () => gsmApi.listVehicles({ activeOnly }),
    staleTime: 10_000,
  });

export const useGsmDriversQuery = (activeOnly = true) =>
  useQuery({
    queryKey: gsmKeys.drivers(activeOnly),
    queryFn: () => gsmApi.listDrivers({ activeOnly }),
    staleTime: 10_000,
  });

export const useGsmCardsQuery = (includeArchived = false) =>
  useQuery({
    queryKey: gsmKeys.cards(includeArchived),
    queryFn: () => gsmApi.listCards({ includeArchived }),
    staleTime: 10_000,
  });

export const useGsmStationsQuery = () =>
  useQuery({
    queryKey: gsmKeys.stations,
    queryFn: () => gsmApi.listStations(),
    staleTime: 10_000,
  });

export const useGsmSettingsQuery = () =>
  useQuery({
    queryKey: gsmKeys.settings,
    queryFn: () => gsmApi.getSettings(),
    staleTime: 30_000,
  });

export const useCreateVehicleMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (payload: VehicleCreatePayload) => gsmApi.createVehicle(payload),
    onSuccess: () => invalidateGsm(qc),
  });
};

export const usePatchVehicleMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ id, payload }: { id: number; payload: VehiclePatchPayload }) =>
      gsmApi.patchVehicle(id, payload),
    onSuccess: () => invalidateGsm(qc),
  });
};

export const useCreateDriverMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (payload: DriverCreatePayload) => gsmApi.createDriver(payload),
    onSuccess: () => invalidateGsm(qc),
  });
};

export const usePatchDriverMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ id, payload }: { id: number; payload: DriverPatchPayload }) =>
      gsmApi.patchDriver(id, payload),
    onSuccess: () => invalidateGsm(qc),
  });
};

export const useCreateCardMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (payload: CardCreatePayload) => gsmApi.createCard(payload),
    onSuccess: () => invalidateGsm(qc),
  });
};

export const usePatchCardMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ id, payload }: { id: number; payload: CardPatchPayload }) =>
      gsmApi.patchCard(id, payload),
    onSuccess: () => invalidateGsm(qc),
  });
};

export const useCreateStationMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (payload: StationCreatePayload) => gsmApi.createStation(payload),
    onSuccess: () => invalidateGsm(qc),
  });
};

export const usePatchStationMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ id, payload }: { id: number; payload: StationPatchPayload }) =>
      gsmApi.patchStation(id, payload),
    onSuccess: () => invalidateGsm(qc),
  });
};

export const usePutGsmSettingsMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (payload: GsmSettings) => gsmApi.putSettings(payload),
    onSuccess: () => {
      void qc.invalidateQueries({ queryKey: gsmKeys.settings });
    },
  });
};

export const useUpdateGsmSeasonMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ mode, date }: { mode: GsmSeasonMode; date: string }) =>
      gsmApi.updateSeason(mode, date),
    onSuccess: () => {
      void qc.invalidateQueries({ queryKey: gsmKeys.settings });
    },
  });
};

export const useImportGsmTransactionsMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (files: File[]) => gsmApi.importTransactions(files),
    onSuccess: () => invalidateGsm(qc),
  });
};

export const useGsmWaybillsQuery = (params: WaybillListParams | null) =>
  useQuery({
    queryKey: gsmKeys.waybills(params),
    queryFn: () => {
      if (!params) {
        throw new Error("waybill list params required");
      }
      return gsmApi.listWaybills(params);
    },
    enabled: Boolean(params?.periodFrom && params.periodTo),
    staleTime: 5_000,
  });

export const useGsmOverviewQuery = (params: OverviewParams | null) =>
  useQuery({
    queryKey: gsmKeys.overview(params),
    queryFn: () => {
      if (!params) {
        throw new Error("overview params required");
      }
      return gsmApi.getOverview(params);
    },
    enabled: Boolean(params?.periodFrom && params.periodTo),
    staleTime: 5_000,
  });

export const useGsmTransactionsQuery = (params: TransactionListParams | null) =>
  useQuery({
    queryKey: gsmKeys.transactions(params),
    queryFn: () => {
      if (!params) {
        throw new Error("transaction list params required");
      }
      return gsmApi.listTransactions(params);
    },
    enabled: Boolean(params?.periodFrom && params.periodTo),
    staleTime: 5_000,
  });

export const useBulkGenerateMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (payload: WaybillBulkGeneratePayload) => gsmApi.generateWaybillsBulk(payload),
    onSuccess: () => {
      void qc.invalidateQueries({ queryKey: gsmKeys.overview(null) });
      void qc.invalidateQueries({ queryKey: gsmKeys.waybills(null) });
      void qc.invalidateQueries({ queryKey: gsmKeys.transactions(null) });
    },
  });
};

export const useGenerateGsmWaybillsMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (payload: WaybillGeneratePayload) => gsmApi.generateWaybills(payload),
    onSuccess: (_data, variables) => {
      // Recalc/generate updates chain_broken and tail badges on the fleet overview.
      void qc.invalidateQueries({ queryKey: gsmKeys.overview(null) });
      void qc.invalidateQueries({
        queryKey: gsmKeys.waybills({
          vehicleId: variables.vehicle_id,
          periodFrom: variables.period_from,
          periodTo: variables.period_to,
        }),
      });
      void qc.invalidateQueries({ queryKey: gsmKeys.waybills(null) });
    },
  });
};

export const useGsmRoutesQuery = (vehicleId: number | null) =>
  useQuery({
    queryKey: gsmKeys.routes(vehicleId),
    queryFn: () => {
      if (vehicleId == null) {
        throw new Error("vehicleId required");
      }
      return gsmApi.listRoutes(vehicleId);
    },
    enabled: vehicleId != null && vehicleId > 0,
    staleTime: 30_000,
  });

const invalidateWaybills = (qc: ReturnType<typeof useQueryClient>) => {
  void qc.invalidateQueries({ queryKey: gsmKeys.waybills(null) });
};

export const usePatchGsmWaybillMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ id, payload }: { id: number; payload: WaybillPatchPayload }) =>
      gsmApi.patchWaybill(id, payload),
    onSuccess: () => invalidateWaybills(qc),
  });
};

export const useCreateGsmWaybillMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (payload: WaybillCreatePayload) => gsmApi.createWaybill(payload),
    onSuccess: () => invalidateWaybills(qc),
  });
};

export const useExportGsmWaybillsMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (payload: WaybillExportPayload) => {
      const result = await gsmApi.exportWaybills(payload);
      saveBlobAs(result.blob, result.filename);
      return result;
    },
    onSuccess: (_data, variables) => {
      const vehicleId = variables.vehicle_ids[0];
      if (vehicleId == null) {
        return;
      }
      void qc.invalidateQueries({
        queryKey: gsmKeys.waybills({
          vehicleId,
          periodFrom: variables.from,
          periodTo: variables.to,
        }),
      });
    },
  });
};

export const useDownloadGsmUsageReportMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (payload: UsageReportPayload) => {
      const result = await gsmApi.downloadUsageReport(payload);
      saveBlobAs(result.blob, result.filename);
      return result;
    },
    onSuccess: () => {
      // Kit export can close the tail; refresh overview badges and waybill lists.
      void qc.invalidateQueries({ queryKey: gsmKeys.overview(null) });
      void qc.invalidateQueries({ queryKey: gsmKeys.waybills(null) });
    },
  });
};

export const useGsmResetToAnchorsMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: () => gsmApi.resetToAnchors(),
    onSuccess: () => invalidateGsm(qc),
  });
};
