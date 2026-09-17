import { useQuery } from "@tanstack/react-query";
import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import type { GuidCheckResponse } from "@/features/commercial-offer/types/commercialOffer";

export const guidCheckKeys = {
  all: ["commercial", "guid-check"] as const,
  draft: (draftId: string) => ["commercial", "guid-check", draftId] as const,
};

export const useGuidCheckQuery = (draftId: string) =>
  useQuery<GuidCheckResponse>({
    queryKey: guidCheckKeys.draft(draftId),
    queryFn: () => commercialOfferApi.checkGuids(draftId),
    enabled: Boolean(draftId),
    staleTime: 15_000,
  });
