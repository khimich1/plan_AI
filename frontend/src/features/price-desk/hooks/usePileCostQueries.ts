import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { pileCostApi } from "@/features/price-desk/api/pileCostApi";
import { priceDeskKeys } from "@/features/price-desk/hooks/usePriceDeskQueries";
import type { PileCostInputsBody } from "@/features/price-desk/types/pileCost";

export const pileCostKeys = {
  all: ["pile-cost"] as const,
  state: (mark: string | null) => ["pile-cost", "state", mark] as const,
};

const selectedMark = (mark?: string | null): string | null => {
  const trimmed = mark?.trim() ?? "";
  return trimmed.length > 0 ? trimmed : null;
};

export const usePileCostQuery = (mark?: string | null) => {
  const markName = selectedMark(mark);
  return useQuery({
    queryKey: pileCostKeys.state(markName),
    queryFn: () => pileCostApi.get(markName),
  });
};

/** PUT ручек. Кэш списка обновляется; статус стола прайсов — нет, pile_prices не менялись. */
export const useSavePileCostInputsMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (body: PileCostInputsBody) => pileCostApi.save(body),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: pileCostKeys.all });
    },
  });
};

/** Превью не пишет цены и не инвалидирует кэш. */
export const usePileCostPreviewMutation = () =>
  useMutation({
    mutationFn: () => pileCostApi.preview(),
  });

/** Apply только с inputs_sha256, который передал вызывающий (отпечаток превью). */
export const usePileCostApplyMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (inputsSha256: string) => pileCostApi.apply({ inputs_sha256: inputsSha256 }),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: pileCostKeys.all });
      void queryClient.invalidateQueries({ queryKey: priceDeskKeys.status });
    },
  });
};
