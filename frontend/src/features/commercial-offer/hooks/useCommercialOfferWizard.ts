import { useCallback, useEffect } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import { resolveDraftProductType, isSimpleKpProductType } from "@/features/commercial-offer/lib/wizardStepOrder";
import { useWizardDraftStore } from "@/features/commercial-offer/store/wizardDraftStore";
import type { CommercialDraftDetails, ProductType, SaveMode, WidePlateAction } from "@/features/commercial-offer/types/commercialOffer";

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

  const currentDraft = draftQuery.data ?? state.lastDraft;
  const draftProductType = resolveDraftProductType(currentDraft?.metadata.product_type ?? state.productType);
  const isPileDraft = draftProductType === "piles";
  const isStepDraft = draftProductType === "steps";
  const isMarchDraft = draftProductType === "marches";
  const isBridgePileDraft = draftProductType === "bridge_piles";
  const isFbsDraft = draftProductType === "fbs";
  const isSimpleKpDraft = isSimpleKpProductType(draftProductType);

  const breakdownQuery = useQuery({
    queryKey: breakdownQueryKey(state.draftId),
    enabled: Boolean(state.draftId) && state.currentStep === "result" && !isSimpleKpDraft,
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

  const updatePilesMutation = useMutation({
    mutationFn: ({ draftId, text, image, mode }: { draftId: string; text: string; image: File | null; mode: "append" | "replace" }) =>
      commercialOfferApi.updateDraftPiles(draftId, { text, image, mode }),
    onSuccess: (draft, variables) => {
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const updateStepsMutation = useMutation({
    mutationFn: ({ draftId, text, image, mode }: { draftId: string; text: string; image: File | null; mode: "append" | "replace" }) =>
      commercialOfferApi.updateDraftSteps(draftId, { text, image, mode }),
    onSuccess: (draft, variables) => {
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const updateMarchesMutation = useMutation({
    mutationFn: ({ draftId, text, image, mode }: { draftId: string; text: string; image: File | null; mode: "append" | "replace" }) =>
      commercialOfferApi.updateDraftMarches(draftId, { text, image, mode }),
    onSuccess: (draft, variables) => {
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const updateBridgePilesMutation = useMutation({
    mutationFn: ({ draftId, text, image, mode }: { draftId: string; text: string; image: File | null; mode: "append" | "replace" }) =>
      commercialOfferApi.updateDraftBridgePiles(draftId, { text, image, mode }),
    onSuccess: (draft, variables) => {
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const updateFbsMutation = useMutation({
    mutationFn: ({ draftId, text, image, mode }: { draftId: string; text: string; image: File | null; mode: "append" | "replace" }) =>
      commercialOfferApi.updateDraftFbs(draftId, { text, image, mode }),
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

  const applyAiPilesMutation = useMutation({
    mutationFn: ({
      draftId,
      instruction,
      image,
    }: {
      draftId: string;
      instruction: string;
      image: File | null;
    }) => commercialOfferApi.applyAiPiles(draftId, { instruction, image }),
    onSuccess: (draft, variables) => {
      dispatch({ type: "start-batch-review", payload: draft });
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const applyAiStepsMutation = useMutation({
    mutationFn: ({
      draftId,
      instruction,
      image,
    }: {
      draftId: string;
      instruction: string;
      image: File | null;
    }) => commercialOfferApi.applyAiSteps(draftId, { instruction, image }),
    onSuccess: (draft, variables) => {
      dispatch({ type: "start-batch-review", payload: draft });
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const applyAiMarchesMutation = useMutation({
    mutationFn: ({
      draftId,
      instruction,
      image,
    }: {
      draftId: string;
      instruction: string;
      image: File | null;
    }) => commercialOfferApi.applyAiMarches(draftId, { instruction, image }),
    onSuccess: (draft, variables) => {
      dispatch({ type: "start-batch-review", payload: draft });
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const applyAiBridgePilesMutation = useMutation({
    mutationFn: ({
      draftId,
      instruction,
      image,
    }: {
      draftId: string;
      instruction: string;
      image: File | null;
    }) => commercialOfferApi.applyAiBridgePiles(draftId, { instruction, image }),
    onSuccess: (draft, variables) => {
      dispatch({ type: "start-batch-review", payload: draft });
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const applyAiFbsMutation = useMutation({
    mutationFn: ({
      draftId,
      instruction,
      image,
    }: {
      draftId: string;
      instruction: string;
      image: File | null;
    }) => commercialOfferApi.applyAiFbs(draftId, { instruction, image }),
    onSuccess: (draft, variables) => {
      dispatch({ type: "start-batch-review", payload: draft });
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const updatePileGradesMutation = useMutation({
    mutationFn: ({ draftId, concreteGrade }: { draftId: string; concreteGrade: string }) =>
      commercialOfferApi.updatePileGrades(draftId, concreteGrade),
    onSuccess: (draft, variables) => {
      dispatch({ type: "hydrate-draft", payload: draft, refreshBatchText: true });
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const updateMarchGradesMutation = useMutation({
    mutationFn: ({ draftId, concreteGrade }: { draftId: string; concreteGrade: string }) =>
      commercialOfferApi.updateMarchGrades(draftId, concreteGrade),
    onSuccess: (draft, variables) => {
      dispatch({ type: "hydrate-draft", payload: draft, refreshBatchText: true });
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const updateBridgePileGradesMutation = useMutation({
    mutationFn: ({ draftId, concreteGrade }: { draftId: string; concreteGrade: string }) =>
      commercialOfferApi.updateBridgePileGrades(draftId, concreteGrade),
    onSuccess: (draft, variables) => {
      dispatch({ type: "hydrate-draft", payload: draft, refreshBatchText: true });
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const updateFbsGradesMutation = useMutation({
    mutationFn: ({ draftId, concreteGrade }: { draftId: string; concreteGrade: string }) =>
      commercialOfferApi.updateFbsGrades(draftId, concreteGrade),
    onSuccess: (draft, variables) => {
      dispatch({ type: "hydrate-draft", payload: draft, refreshBatchText: true });
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

  const resolveUnpricedPlatesMutation = useMutation({
    mutationFn: ({
      draftId,
      decisions,
    }: {
      draftId: string;
      decisions: Array<{
        lineId?: string;
        sourceLine: string;
        action: "replace_load" | "exclude";
        loadCode?: number | null;
      }>;
    }) => commercialOfferApi.resolveUnpricedPlates(draftId, decisions),
    onSuccess: (draft, variables) => {
      dispatch({ type: "sync-after-unpriced-plates", payload: draft });
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
      pileLogisticsCost,
      pileTripOverrides,
    }: {
      draftId: string;
      managerId?: number | null;
      clientName?: string;
      discountPercent?: number;
      conditionsMode?: "standard" | "custom";
      deliveryConditions?: string;
      paymentConditions?: string;
      logisticsCost?: number;
      pileLogisticsCost?: number;
      pileTripOverrides?: Record<string, number>;
    }) =>
      commercialOfferApi.updateDraftMeta(draftId, {
        managerId,
        clientName,
        discountPercent,
        conditionsMode,
        deliveryConditions,
        paymentConditions,
        logisticsCost,
        pileLogisticsCost,
        pileTripOverrides,
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
    mutationFn: ({ draftId, fileTypes }: { draftId: string; fileTypes?: Array<"pdf" | "xlsx" | "breakdown" | "schema"> }) =>
      commercialOfferApi.generateFiles(draftId, fileTypes),
    onSettled: (_data, _error, variables) => {
      invalidateDraft(variables.draftId);
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

  const startAppendCycleMutation = useMutation({
    mutationFn: ({ draftId, productType }: { draftId: string; productType: ProductType }) =>
      commercialOfferApi.startAppendCycle(draftId, productType),
    onSuccess: (draft, variables) => {
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const undoLastAppendBatchMutation = useMutation({
    mutationFn: (draftId: string) => commercialOfferApi.undoLastAppendBatch(draftId),
    onSuccess: (draft, draftId) => {
      setDraftCache(draftId, draft);
      invalidateDraft(draftId);
    },
  });

  const deleteDraftLineMutation = useMutation({
    mutationFn: ({ draftId, lineId }: { draftId: string; lineId: string }) =>
      commercialOfferApi.deleteDraftLine(draftId, lineId),
    onSuccess: (draft, variables) => {
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const patchDraftLineMutation = useMutation({
    mutationFn: ({
      draftId,
      lineId,
      payload,
    }: {
      draftId: string;
      lineId: string;
      payload: { qty?: number; source_text?: string };
    }) => commercialOfferApi.patchDraftLine(draftId, lineId, payload),
    onSuccess: (draft, variables) => {
      setDraftCache(variables.draftId, draft);
      invalidateDraft(variables.draftId);
    },
  });

  const restoreDraftLinesMutation = useMutation({
    mutationFn: ({
      draftId,
      payload,
    }: {
      draftId: string;
      payload: { index: number; lines: Record<string, unknown>[]; replace_line_ids?: string[] };
    }) => commercialOfferApi.restoreDraftLines(draftId, payload),
    onSuccess: (draft, variables) => {
      setDraftCache(variables.draftId, draft);
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
    updatePilesMutation,
    updateStepsMutation,
    updateMarchesMutation,
    updateBridgePilesMutation,
    updateFbsMutation,
    applyAiPlatesMutation,
    applyAiPilesMutation,
    applyAiStepsMutation,
    applyAiMarchesMutation,
    applyAiBridgePilesMutation,
    applyAiFbsMutation,
    updatePileGradesMutation,
    updateMarchGradesMutation,
    updateBridgePileGradesMutation,
    updateFbsGradesMutation,
    resolveWidePlatesMutation,
    resolveUnpricedPlatesMutation,
    updateMetaMutation,
    calculateMutation,
    generateFilesMutation,
    generateSchemaMutation,
    saveDraftMutation,
    startAppendCycleMutation,
    undoLastAppendBatchMutation,
    deleteDraftLineMutation,
    patchDraftLineMutation,
    restoreDraftLinesMutation,
    currentDraft,
    isPileDraft,
    isStepDraft,
    isMarchDraft,
    isBridgePileDraft,
    isFbsDraft,
    isSimpleKpDraft,
  };
};
