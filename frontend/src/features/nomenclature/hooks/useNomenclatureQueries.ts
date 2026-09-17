import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { nomenclatureApi } from "@/features/nomenclature/api/nomenclatureApi";
import type { GuidTasksResponse, NomenclatureProductKind } from "@/features/nomenclature/types/nomenclature";

export const nomenclatureKeys = {
  all: ["nomenclature"] as const,
  tasks: ["nomenclature", "tasks"] as const,
  duplicates: ["nomenclature", "duplicates"] as const,
};

export const useImport1cMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({
      file,
      productKind,
    }: {
      file: File;
      productKind?: NomenclatureProductKind | null;
    }) => nomenclatureApi.import1c(file, productKind),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: nomenclatureKeys.all });
    },
  });
};

export const useGuidTasksQuery = () =>
  useQuery<GuidTasksResponse>({
    queryKey: nomenclatureKeys.tasks,
    queryFn: () => nomenclatureApi.listTasks(),
  });

export const useGuidDuplicatesQuery = () =>
  useQuery<GuidTasksResponse>({
    queryKey: nomenclatureKeys.duplicates,
    queryFn: () => nomenclatureApi.listDuplicates(),
  });

export const useResolvePriceMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ guid, price }: { guid: string; price: number }) =>
      nomenclatureApi.resolvePrice(guid, price),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: nomenclatureKeys.all });
    },
  });
};

export const useResolveDuplicateMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (payload: {
      scope: string;
      key: string;
      chosen_guid: string;
      note?: string;
      product_kind?: string | null;
    }) => nomenclatureApi.resolveDuplicate(payload),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: nomenclatureKeys.all });
    },
  });
};
