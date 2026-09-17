import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { priceDeskApi } from "@/features/price-desk/api/priceDeskApi";

export const priceDeskKeys = {
  all: ["price-desk"] as const,
  status: ["price-desk", "status"] as const,
};

export const usePriceDeskStatusQuery = () =>
  useQuery({
    queryKey: priceDeskKeys.status,
    queryFn: () => priceDeskApi.status(),
  });

export const usePriceDeskPreviewMutation = () =>
  useMutation({
    mutationFn: (file: File) => priceDeskApi.preview(file),
  });

export const usePriceDeskApplyMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ file, fileSha256 }: { file: File; fileSha256: string }) =>
      priceDeskApi.apply(file, fileSha256),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: priceDeskKeys.status });
    },
  });
};
