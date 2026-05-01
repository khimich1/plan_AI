import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import { useWizardDraftStore } from "@/features/commercial-offer/store/wizardDraftStore";
import type { CommercialDraftDetails, SaveMode, WidePlateAction } from "@/features/commercial-offer/types/commercialOffer";

const draftQueryKey = (draftId: string | null) => ["commercial-offer-draft", draftId] as const;

export const useCommercialOfferWizard = () => {
  const { state, dispatch } = useWizardDraftStore();
  const queryClient = useQueryClient();

  const managersQuery = useQuery({
    queryKey: ["commercial-offer-managers"],
    queryFn: commercialOfferApi.getManagers,
  });

  const syncDraft = (draft: CommercialDraftDetails) => {
    dispatch({ type: "hydrate-draft", payload: draft });
    queryClient.setQueryData(draftQueryKey(draft.draft_id), draft);
  };

  const draftQuery = useQuery({
    queryKey: draftQueryKey(state.draftId),
    enabled: Boolean(state.draftId),
    queryFn: async () => {
      if (!state.draftId) {
        throw new Error("Draft is not initialized.");
      }
      return commercialOfferApi.getDraft(state.draftId);
    },
    onSuccess: syncDraft,
  });

  const createDraftMutation = useMutation({
    mutationFn: commercialOfferApi.createDraft,
    onSuccess: syncDraft,
  });

  const updatePlatesMutation = useMutation({
    mutationFn: ({ draftId, text, image, mode }: { draftId: string; text: string; image: File | null; mode: "append" | "replace" }) =>
      commercialOfferApi.updateDraftPlates(draftId, { text, image, mode }),
    onSuccess: syncDraft,
  });

  const resolveWidePlatesMutation = useMutation({
    mutationFn: ({
      draftId,
      decisions,
    }: {
      draftId: string;
      decisions: Array<{ sourceLine: string; action: WidePlateAction; replacementText: string }>;
    }) => commercialOfferApi.resolveWidePlates(draftId, decisions),
    onSuccess: syncDraft,
  });

  const updateMetaMutation = useMutation({
    mutationFn: ({
      draftId,
      managerId,
      clientName,
      discountPercent,
      conditionsMode,
      deliveryConditions,
      paymentConditions,
      logisticsCost,
    }: {
      draftId: string;
      managerId?: number | null;
      clientName?: string;
      discountPercent?: number;
      conditionsMode?: "standard" | "custom";
      deliveryConditions?: string;
      paymentConditions?: string;
      logisticsCost?: number;
    }) =>
      commercialOfferApi.updateDraftMeta(draftId, {
        managerId,
        clientName,
        discountPercent,
        conditionsMode,
        deliveryConditions,
        paymentConditions,
        logisticsCost,
      }),
    onSuccess: syncDraft,
  });

  const calculateMutation = useMutation({
    mutationFn: (draftId: string) => commercialOfferApi.calculateDraft(draftId),
    onSuccess: syncDraft,
  });

  const generateFilesMutation = useMutation({
    mutationFn: (draftId: string) => commercialOfferApi.generateFiles(draftId),
    onSuccess: async (payload) => {
      if (!state.draftId || !state.lastDraft) {
        return;
      }
      const mergedDraft: CommercialDraftDetails = {
        ...state.lastDraft,
        files: payload.files,
      };
      syncDraft(mergedDraft);
    },
  });

  const saveDraftMutation = useMutation({
    mutationFn: ({
      draftId,
      mode,
      executionTermsInput,
    }: {
      draftId: string;
      mode: SaveMode;
      executionTermsInput?: string;
    }) => commercialOfferApi.saveDraft(draftId, { mode, executionTermsInput }),
    onSuccess: async (result) => {
      dispatch({ type: "set-save-result", payload: result });
      if (state.draftId) {
        const freshDraft = await commercialOfferApi.getDraft(state.draftId);
        syncDraft(freshDraft);
      }
    },
  });

  return {
    state,
    dispatch,
    managersQuery,
    draftQuery,
    createDraftMutation,
    updatePlatesMutation,
    resolveWidePlatesMutation,
    updateMetaMutation,
    calculateMutation,
    generateFilesMutation,
    saveDraftMutation,
    currentDraft: draftQuery.data ?? state.lastDraft,
  };
};
