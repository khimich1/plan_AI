import { useCallback, useEffect } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import { useWizardDraftStore } from "@/features/commercial-offer/store/wizardDraftStore";
import type { CommercialDraftDetails, SaveMode, WidePlateAction } from "@/features/commercial-offer/types/commercialOffer";

const draftQueryKey = (draftId: string | null) => ["commercial-offer-draft", draftId] as const;
const breakdownQueryKey = (draftId: string | null) => ["commercial-offer-breakdown", draftId] as const;

export const useCommercialOfferWizard = () => {
  const { state, dispatch } = useWizardDraftStore();
  const queryClient = useQueryClient();

  const managersQuery = useQuery({
    queryKey: ["commercial-offer-managers"],
    queryFn: commercialOfferApi.getManagers,
  });

  const draftQuery = useQuery({
    queryKey: draftQueryKey(state.draftId),
    enabled: Boolean(state.draftId),
    queryFn: async () => {
      if (!state.draftId) {
        throw new Error("Draft is not initialized.");
      }
      return commercialOfferApi.getDraft(state.draftId);
    },
  });

  const breakdownQuery = useQuery({
    queryKey: breakdownQueryKey(state.draftId),
    enabled: Boolean(state.draftId) && state.currentStep === "result",
    queryFn: async () => {
      if (!state.draftId) {
        throw new Error("Draft is not initialized.");
      }
      return commercialOfferApi.getBreakdown(state.draftId);
    },
  });

  useEffect(() => {
    if (draftQuery.data) {
      dispatch({ type: "hydrate-draft", payload: draftQuery.data });
    }
  }, [draftQuery.data, dispatch]);

  const invalidateDraft = useCallback(
    (draftId: string) => {
      void queryClient.invalidateQueries({ queryKey: draftQueryKey(draftId) });
      void queryClient.invalidateQueries({ queryKey: breakdownQueryKey(draftId) });
    },
    [queryClient],
  );

  const setDraftCache = useCallback(
    (draftId: string, draft: CommercialDraftDetails) => {
      queryClient.setQueryData(draftQueryKey(draftId), draft);
    },
    [queryClient],
  );

  const createDraftMutation = useMutation({
    mutationFn: commercialOfferApi.createDraft,
    onSuccess: (draft) => {
      dispatch({ type: "hydrate-draft", payload: draft });
      setDraftCache(draft.draft_id, draft);
      invalidateDraft(draft.draft_id);
    },
  });

  const updatePlatesMutation = useMutation({
    mutationFn: ({ draftId, text, image, mode }: { draftId: string; text: string; image: File | null; mode: "append" | "replace" }) =>
      commercialOfferApi.updateDraftPlates(draftId, { text, image, mode }),
    onSuccess: (draft, variables) => {
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const applyAiPlatesMutation = useMutation({
    mutationFn: ({
      draftId,
      instruction,
      image,
    }: {
      draftId: string;
      instruction: string;
      image: File | null;
    }) => commercialOfferApi.applyAiPlates(draftId, { instruction, image }),
    onSuccess: (draft, variables) => {
      dispatch({ type: "start-batch-review", payload: draft });
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const resolveWidePlatesMutation = useMutation({
    mutationFn: ({
      draftId,
      decisions,
    }: {
      draftId: string;
      decisions: Array<{ sourceLine: string; action: WidePlateAction; replacementText: string }>;
    }) => commercialOfferApi.resolveWidePlates(draftId, decisions),
    onSuccess: (draft, variables) => {
      dispatch({ type: "sync-after-wide-plates", payload: draft });
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
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
    onSuccess: (draft, variables) => {
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const calculateMutation = useMutation({
    mutationFn: (draftId: string) => commercialOfferApi.calculateDraft(draftId),
    onSuccess: (draft, draftId) => {
      setDraftCache(draftId, draft);
      invalidateDraft(draftId);
    },
  });

  const generateFilesMutation = useMutation({
    mutationFn: (draftId: string) => commercialOfferApi.generateFiles(draftId),
    onSettled: (_data, _error, draftId) => {
      invalidateDraft(draftId);
    },
  });

  const generateSchemaMutation = useMutation({
    mutationFn: (draftId: string) => commercialOfferApi.generateSchemaFiles(draftId),
    onSettled: (_data, _error, draftId) => {
      invalidateDraft(draftId);
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
    onSuccess: (result, variables) => {
      dispatch({ type: "set-save-result", payload: result });
      invalidateDraft(variables.draftId);
    },
  });

  return {
    state,
    dispatch,
    managersQuery,
    draftQuery,
    breakdownQuery,
    createDraftMutation,
    updatePlatesMutation,
    applyAiPlatesMutation,
    resolveWidePlatesMutation,
    updateMetaMutation,
    calculateMutation,
    generateFilesMutation,
    generateSchemaMutation,
    saveDraftMutation,
    currentDraft: draftQuery.data ?? state.lastDraft,
  };
};
