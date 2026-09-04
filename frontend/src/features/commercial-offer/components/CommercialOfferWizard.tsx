import { useCallback, useEffect, useRef, useState } from "react";
import { useNavigate } from "react-router";
import { useQueryClient } from "@tanstack/react-query";

import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import { archiveKeys } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { planApplyAiSessionSync } from "@/features/commercial-offer/lib/applyAiSession";
import { getBatches, getCurrentBatchReviewText, mergeEditedBatchIntoFullText } from "@/features/commercial-offer/lib/batchReview";
import { getDraftBatchCount } from "@/features/commercial-offer/lib/getDraftBatchCount";
import {
  buildPileLinesFromOrderData,
  buildPilePreviewRows,
} from "@/features/commercial-offer/lib/buildPilePreviewRows";
import {
  buildBridgePileLinesFromOrderData,
  buildBridgePilePreviewRows,
} from "@/features/commercial-offer/lib/buildBridgePilePreviewRows";
import {
  buildFbsLinesFromOrderData,
  buildFbsPreviewRows,
} from "@/features/commercial-offer/lib/buildFbsPreviewRows";
import {
  buildMarchLinesFromOrderData,
  buildMarchPreviewRows,
} from "@/features/commercial-offer/lib/buildMarchPreviewRows";
import {
  getProductInputStep,
  getWizardStepOrder,
  isInputStepBlockedWithoutAppendCycle,
  mapLegacyWizardStep,
  shouldSkipClientStep,
} from "@/features/commercial-offer/lib/wizardStepOrder";

import { useCommercialOfferWizard } from "@/features/commercial-offer/hooks/useCommercialOfferWizard";
import {
  useMultiPageRecognize,
  type RecognizePageArgs,
  type RerunPageArgs,
} from "@/features/commercial-offer/hooks/useMultiPageRecognize";
import { useRecognizedImagePreview } from "@/features/commercial-offer/hooks/useRecognizedImagePreview";
import { useSourceImageQueue } from "@/features/commercial-offer/hooks/useSourceImageQueue";
import { applyPromoteSourceImageQueue } from "@/features/commercial-offer/lib/promoteSourceImageQueue";
import { liveWidePlateLines } from "@/features/commercial-offer/lib/liveWidePlateLines";
import {
  buildMergedFlushText,
  flushThenResolveWidePlates,
} from "@/features/commercial-offer/lib/flushThenResolveWidePlates";

import { WizardProgress } from "@/features/commercial-offer/components/WizardProgress";
import { ProductTypePicker } from "@/features/commercial-offer/components/ProductTypePicker";
import { PlateInputStep } from "@/features/commercial-offer/components/steps/PlateInputStep";
import { PileInputStep } from "@/features/commercial-offer/components/steps/PileInputStep";
import { MarchInputStep } from "@/features/commercial-offer/components/steps/MarchInputStep";
import { BridgePileInputStep } from "@/features/commercial-offer/components/steps/BridgePileInputStep";
import { FbsInputStep } from "@/features/commercial-offer/components/steps/FbsInputStep";
import { StepInputStep } from "@/features/commercial-offer/components/steps/StepInputStep";
import { ClientConditionsStep } from "@/features/commercial-offer/components/steps/ClientConditionsStep";
import { CalculationResultStep } from "@/features/commercial-offer/components/steps/CalculationResultStep";

import type { ProductType, WizardStepId } from "@/features/commercial-offer/types/commercialOffer";
import type { LineRowHandlers, LineSavePayload } from "@/features/commercial-offer/lib/lineRowHandlers";
import { LINE_UNDO_TOAST_MS } from "@/features/commercial-offer/lib/lineRowHandlers";

import { getErrorMessage } from "@/shared/lib/apiError";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Spinner } from "@/shared/ui/Spinner";

export const CommercialOfferWizard = ({ productType: productTypeProp }: { productType: ProductType }) => {
  const {
    state,
    dispatch,
    managersQuery,
    draftQuery,
    breakdownQuery,
    currentDraft,
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
    resolveInvalidWidthsMutation,
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
    isPileDraft,
    isStepDraft,
    isMarchDraft,
    isBridgePileDraft,
    isFbsDraft,
    isSimpleKpDraft,
  } = useCommercialOfferWizard();
  const navigate = useNavigate();
  const queryClient = useQueryClient();

  const [aiInstruction, setAiInstruction] = useState("");
  const [stepError, setStepError] = useState<string | null>(null);
  const [lineUndoToastMessage, setLineUndoToastMessage] = useState<string | null>(null);
  const [lineRowError, setLineRowError] = useState<{ lineId: string; message: string } | null>(null);
  const lineUndoRef = useRef<
    | { kind: "qty"; lineId: string; qty: number }
    | { kind: "restore"; index: number; lines: Record<string, unknown>[]; replaceLineIds: string[] }
    | null
  >(null);
  const lineUndoTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const [widePlateError, setWidePlateError] = useState<string | null>(null);
  const [unpricedPlateError, setUnpricedPlateError] = useState<string | null>(null);
  const [invalidWidthError, setInvalidWidthError] = useState<string | null>(null);
  const {
    preview: recognizedImagePreview,
    setPreviewFromFile,
    clearPreview: clearRecognizedImagePreview,
    takePreview: takeRecognizedImagePreview,
  } = useRecognizedImagePreview();
  const sourceImageQueue = useSourceImageQueue();

  const productType = state.productType;
  const isPileFlow = productType === "piles";
  const isStepFlow = productType === "steps";
  const isMarchFlow = productType === "marches";
  const isBridgePileFlow = productType === "bridge_piles";
  const isFbsFlow = productType === "fbs";
  const isSimpleProductFlow = isPileFlow || isStepFlow || isMarchFlow || isBridgePileFlow || isFbsFlow;
  const skipClient = shouldSkipClientStep({
    clientName: state.clientName,
    appendBatches: state.lastDraft?.metadata.append_batches ?? currentDraft?.metadata.append_batches,
    resumeKpId: state.lastDraft?.metadata.resume_kp_id ?? currentDraft?.metadata.resume_kp_id ?? null,
  });
  const stepOrder = getWizardStepOrder(productType, { skipClient });
  const inputStep = getProductInputStep(productType);
  const stepIndex = (step: WizardStepId) => stepOrder.indexOf(step);
  const updateInputMutation = isFbsFlow
    ? updateFbsMutation
    : isBridgePileFlow
    ? updateBridgePilesMutation
    : isMarchFlow
      ? updateMarchesMutation
      : isStepFlow
        ? updateStepsMutation
        : isPileFlow
          ? updatePilesMutation
          : updatePlatesMutation;
  const applyAiMutation = isFbsFlow
    ? applyAiFbsMutation
    : isBridgePileFlow
    ? applyAiBridgePilesMutation
    : isMarchFlow
      ? applyAiMarchesMutation
      : isStepFlow
        ? applyAiStepsMutation
        : isPileFlow
          ? applyAiPilesMutation
          : applyAiPlatesMutation;

  const managers = managersQuery.data?.items ?? [];

  const recognizePage = useCallback(
    async ({ image, productType: pageProductType, draftId, isFirst }: RecognizePageArgs) => {
      let draft;
      if (!draftId) {
        draft = await createDraftMutation.mutateAsync({
          text: "",
          image,
          productType: pageProductType,
        });
      } else {
        draft = await updateInputMutation.mutateAsync({
          draftId,
          text: "",
          image,
          mode: isFirst ? "replace" : "append",
        });
      }
      const batchReviewText = getCurrentBatchReviewText(draft);
      if (isFirst) {
        dispatch({ type: "start-batch-review", payload: draft });
      } else {
        dispatch({ type: "hydrate-draft", payload: draft, refreshBatchText: false });
      }
      return { draft, batchReviewText };
    },
    [createDraftMutation, updateInputMutation, dispatch],
  );

  const rerunOcrPage = useCallback(
    async ({ image, draftId }: RerunPageArgs) => commercialOfferApi.ocrPage(draftId, image),
    [],
  );

  const multiPage = useMultiPageRecognize({ recognizePage, rerunPage: rerunOcrPage });

  useEffect(() => {
    if (state.isPickingProductType) {
      return;
    }
    if (productTypeProp !== state.productType) {
      dispatch({ type: "set-product-type", productType: productTypeProp });
    }
  }, [productTypeProp, state.productType, state.isPickingProductType, dispatch]);

  useEffect(() => {
    const step = mapLegacyWizardStep(state.currentStep);
    if (step !== state.currentStep) {
      dispatch({ type: "set-step", step });
      return;
    }
    if (skipClient && step === "client") {
      dispatch({ type: "set-step", step: state.draftId ? "result" : inputStep });
      return;
    }
    if (!stepOrder.includes(step)) {
      dispatch({ type: "set-step", step: inputStep });
      return;
    }
    if (step === "result" && !state.draftId) {
      dispatch({ type: "set-step", step: inputStep });
    }
  }, [state.currentStep, state.draftId, dispatch, stepOrder, inputStep, skipClient]);

  const resetSource = () => {
    multiPage.reset();
    dispatch({ type: "set-source", text: "", imageName: null });
  };

  const handleSourceTextChange = (value: string) => {
    if (multiPage.pages.length > 0 || sourceImageQueue.length > 0) {
      sourceImageQueue.clear();
      multiPage.reset();
      clearRecognizedImagePreview();
    }
    dispatch({ type: "set-source", text: value, imageName: null });
  };

  const handleAddFiles = (files: File[]) => {
    if (files.length === 0) {
      return;
    }
    const nextPages = multiPage.addFiles(files);
    dispatch({
      type: "set-source",
      text: "",
      imageName: nextPages[0]?.name ?? (nextPages.length > 0 ? `${nextPages.length} фото` : null),
    });
  };

  const handleRemovePage = (id: string) => {
    const remaining = multiPage.remove(id);
    dispatch({
      type: "set-source",
      text: state.sourceText,
      imageName: remaining[0]?.name ?? null,
    });
  };

  const handleSelectPage = (id: string) => {
    multiPage.setActive(id);
    const page = multiPage.pages.find((item) => item.id === id);
    if (page && (page.status === "ready" || page.status === "confirmed")) {
      dispatch({ type: "set-batch-review-text", text: page.batchReviewText });
    }
  };

  const handleRerecognize = useCallback(async () => {
    if (multiPage.hasStarted && multiPage.activeId) {
      const pageId = multiPage.activeId;
      const result = await multiPage.rerunPage(pageId);
      if (result) {
        dispatch({ type: "set-batch-review-text", text: result.normalized_text });
      }
      return;
    }
    const image = recognizedImagePreview?.file ?? null;
    const draftId = currentDraft?.draft_id;
    if (!image || !draftId) {
      return;
    }
    try {
      const draft = await updateInputMutation.mutateAsync({
        draftId,
        text: "",
        image,
        mode: "replace",
      });
      dispatch({ type: "hydrate-draft", payload: draft, refreshBatchText: true });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  }, [
    currentDraft?.draft_id,
    dispatch,
    multiPage,
    recognizedImagePreview?.file,
    updateInputMutation,
  ]);

  // Keep store batch text aligned when OCR marks a page ready / auto-focuses it.
  useEffect(() => {
    if (!multiPage.hasStarted || !multiPage.activePage) {
      return;
    }
    if (multiPage.activePage.status === "ready" || multiPage.activePage.status === "confirmed") {
      dispatch({ type: "set-batch-review-text", text: multiPage.activePage.batchReviewText });
    }
  }, [
    multiPage.hasStarted,
    multiPage.activeId,
    multiPage.activePage?.status,
    multiPage.activePage?.batchReviewText,
    dispatch,
  ]);

  const handlePrevPage = () => {
    const navigable = multiPage.pages.filter(
      (page) => page.status === "ready" || page.status === "confirmed",
    );
    const index = navigable.findIndex((page) => page.id === multiPage.activeId);
    if (index > 0) {
      handleSelectPage(navigable[index - 1]!.id);
    }
  };

  const handleNextPage = () => {
    const navigable = multiPage.pages.filter(
      (page) => page.status === "ready" || page.status === "confirmed",
    );
    const index = navigable.findIndex((page) => page.id === multiPage.activeId);
    if (index >= 0 && index < navigable.length - 1) {
      handleSelectPage(navigable[index + 1]!.id);
    }
  };

  const handleRecognize = async (mode: "append" | "replace") => {
    setStepError(null);
    const hasPages = multiPage.pages.length > 0;
    if (!state.sourceText.trim() && !hasPages) {
      setStepError(
        isFbsFlow
          ? "Введите текст списка ФБС или загрузите изображение."
          : isBridgePileFlow
          ? "Введите текст списка мостовых свай или загрузите изображение."
          : isMarchFlow
          ? "Введите текст списка маршей или загрузите изображение."
          : isStepFlow
          ? "Введите текст списка ступеней или загрузите изображение."
          : isPileFlow
            ? "Введите текст списка свай или загрузите изображение."
            : "Введите текст списка плит или загрузите изображение.",
      );
      return;
    }

    // Multi-page first replace pass: sequential OCR pipeline.
    if (hasPages && mode === "replace" && !currentDraft?.draft_id) {
      try {
        void multiPage.start({ productType });
      } catch (error) {
        setStepError(getErrorMessage(error));
      }
      return;
    }

    const sourceText = hasPages ? "" : state.sourceText;
    const imageForRecognition = hasPages ? multiPage.pages[0]?.file ?? null : null;

    try {
      if (imageForRecognition) {
        setPreviewFromFile(imageForRecognition);
      }
      let draft;
      if (currentDraft?.draft_id) {
        draft = await updateInputMutation.mutateAsync({
          draftId: currentDraft.draft_id,
          text: sourceText,
          image: imageForRecognition,
          mode,
        });
      } else {
        draft = await createDraftMutation.mutateAsync({
          text: sourceText,
          image: imageForRecognition,
          productType,
        });
      }
      dispatch({ type: "start-batch-review", payload: draft });
      resetSource();
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleApplyAi = async () => {
    setStepError(null);
    const instruction = aiInstruction.trim();
    if (!instruction) {
      setStepError("Введите инструкцию для помощника.");
      return;
    }
    if (!currentDraft?.draft_id) {
      setStepError(
        isMarchFlow
          ? "Сначала распознайте список маршей, затем используйте помощника."
          : isStepFlow
          ? "Сначала распознайте список ступеней, затем используйте помощника."
          : isPileFlow
            ? "Сначала распознайте список свай, затем используйте помощника."
            : "Сначала распознайте список плит, затем используйте помощника.",
      );
      return;
    }

    try {
      const image = multiPage.pages[0]?.file ?? null;
      if (image) {
        setPreviewFromFile(image);
      }
      const draft = await applyAiMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        instruction,
        image,
      });
      const sessionPlan = planApplyAiSessionSync({
        multiHasStarted: multiPage.hasStarted,
        activePageId: multiPage.activeId,
        draft,
      });
      if (sessionPlan.nextActivePageText !== null) {
        dispatch({ type: "set-batch-review-text", text: sessionPlan.nextActivePageText });
        if (sessionPlan.preserveMultiSession && multiPage.activeId) {
          multiPage.updatePageText(multiPage.activeId, sessionPlan.nextActivePageText);
        }
      }
      if (!sessionPlan.preserveMultiSession) {
        resetSource();
      }
      setAiInstruction("");
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleConfirmBatch = async () => {
    setStepError(null);
    if (!currentDraft?.draft_id) {
      setStepError("Не удалось загрузить данные черновика. Обновите страницу или начните заново.");
      return;
    }

    // Multi-page: confirm active ready page; finalize when all confirmed.
    if (multiPage.hasStarted && multiPage.pages.length > 0) {
      const active = multiPage.activePage;
      if (!active || active.status !== "ready") {
        return;
      }
      const editedActiveText = state.batchReviewText;
      multiPage.updatePageText(active.id, editedActiveText);
      const textsSnapshot = multiPage.pages.map((page) =>
        page.id === active.id ? editedActiveText.trim() : page.batchReviewText.trim(),
      );
      const { allConfirmed, nextBatchReviewText } = multiPage.confirmActive();
      if (!allConfirmed) {
        dispatch({ type: "set-batch-review-text", text: nextBatchReviewText });
        return;
      }

      const mergedText = textsSnapshot.filter(Boolean).join("\n");
      try {
        let draft = currentDraft;
        if (mergedText && mergedText !== (draft.metadata.normalized_text ?? "").trim()) {
          draft = await updateInputMutation.mutateAsync({
            draftId: draft.draft_id,
            text: mergedText,
            image: null,
            mode: "replace",
          });
          dispatch({ type: "hydrate-draft", payload: draft, refreshBatchText: true });
        }
        // Promote-then-reset: snapshot independent blob URLs before OCR reset.
        applyPromoteSourceImageQueue(sourceImageQueue, {
          pages: multiPage.pages,
          singlePreview: null,
        });
        multiPage.reset();
        clearRecognizedImagePreview();
        dispatch({
          type: "confirm-batch-review",
          batchCount: getDraftBatchCount(draft),
        });
      } catch (error) {
        setStepError(getErrorMessage(error));
      }
      return;
    }

    let draft = currentDraft;
    const batches = getBatches(draft);
    const lastBatch = batches.length > 0 ? batches[batches.length - 1] : undefined;
    const editedText = state.batchReviewText.trim();
    const originalBatchText = (lastBatch?.normalized_text ?? "").trim();

    if (editedText && editedText !== originalBatchText) {
      try {
        const mergedText = mergeEditedBatchIntoFullText(batches, editedText);
        draft = await updateInputMutation.mutateAsync({
          draftId: draft.draft_id,
          text: mergedText,
          image: null,
          mode: "replace",
        });
        dispatch({ type: "hydrate-draft", payload: draft, refreshBatchText: true });
        dispatch({
          type: "confirm-batch-review",
          batchCount: getDraftBatchCount(draft),
        });
      } catch (error) {
        setStepError(getErrorMessage(error));
        return;
      }
    } else {
      dispatch({ type: "confirm-batch-review", batchCount: getDraftBatchCount(draft) });
    }

    // Transfer File into queue (fresh object URL); takePreview already cleared OCR preview.
    const takenPreview = takeRecognizedImagePreview();
    applyPromoteSourceImageQueue(sourceImageQueue, {
      pages: [],
      singlePreview: takenPreview,
    });
  };

  const proceedFromInputStep = async (next: WizardStepId) => {
    if (next !== "result") {
      dispatch({ type: "set-step", step: next });
      return;
    }
    if (!currentDraft?.draft_id) {
      setStepError("Не удалось загрузить данные черновика. Обновите страницу или начните заново.");
      return;
    }
    try {
      const calculated = await calculateMutation.mutateAsync(currentDraft.draft_id);
      dispatch({ type: "hydrate-draft", payload: calculated });
      const ws = calculated.wizard_state;
      if (ws.current_step !== "result") {
        const msgs = (ws.validation_errors ?? []).filter(Boolean);
        setStepError(
          msgs.length > 0
            ? msgs.join(" ")
            : "Расчёт не завершён. Проверьте данные заказа и условия.",
        );
        return;
      }
      dispatch({ type: "set-step", step: "result" });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleFinishPlates = async () => {
    setStepError(null);
    if (state.pendingBatchReview) {
      setStepError("Сначала подтвердите список текущего источника — нажмите «Список верен».");
      return;
    }
    if (!currentDraft?.wizard_state || !currentDraft.draft_id) {
      setStepError("Не удалось загрузить данные черновика. Обновите страницу или начните заново.");
      return;
    }

    const draft = currentDraft;
    const next = draft.wizard_state.can_proceed_to[0];
    if (!next) {
      const serverMsgs = (draft.wizard_state.validation_errors ?? []).filter(Boolean);
      if (serverMsgs.length > 0) {
        setStepError(serverMsgs.join(" "));
      } else if (draft.wizard_state.next_required_action === "ingest_plates") {
        setStepError("Сначала распознайте и получите хотя бы одну позицию в заказе.");
      } else if (draft.wizard_state.next_required_action === "resolve_wide_plates") {
        setStepError("Сначала примите решение по позициям шире стандартной.");
      } else if (draft.wizard_state.next_required_action === "resolve_invalid_widths") {
        setStepError("Нестандартная ширина: замените на заводской рез или исключите позицию.");
      } else if (draft.wizard_state.next_required_action === "resolve_unpriced_plates") {
        setStepError("Сначала примите решение по позициям без цены в прайсе.");
      } else {
        setStepError("Нельзя перейти дальше — проверьте список плит и повторите.");
      }
      return;
    }
    await proceedFromInputStep(next);
  };

  const handleFinishPiles = async () => {
    setStepError(null);
    if (state.pendingBatchReview) {
      setStepError("Сначала подтвердите список текущего источника — нажмите «Список верен».");
      return;
    }
    if (!currentDraft?.wizard_state || !currentDraft.draft_id) {
      setStepError("Не удалось загрузить данные черновика. Обновите страницу или начните заново.");
      return;
    }

    const draft = currentDraft;
    const next = draft.wizard_state.can_proceed_to[0];
    if (!next) {
      const serverMsgs = (draft.wizard_state.validation_errors ?? []).filter(Boolean);
      if (serverMsgs.length > 0) {
        setStepError(serverMsgs.join(" "));
      } else if (draft.wizard_state.next_required_action === "ingest_piles") {
        setStepError("Сначала распознайте и получите хотя бы одну позицию в заказе.");
      } else {
        setStepError("Нельзя перейти дальше — проверьте список свай и повторите.");
      }
      return;
    }
    await proceedFromInputStep(next);
  };

  const handleFinishSteps = async () => {
    setStepError(null);
    if (state.pendingBatchReview) {
      setStepError("Сначала подтвердите список текущего источника — нажмите «Список верен».");
      return;
    }
    if (!currentDraft?.wizard_state || !currentDraft.draft_id) {
      setStepError("Не удалось загрузить данные черновика. Обновите страницу или начните заново.");
      return;
    }

    const draft = currentDraft;
    const next = draft.wizard_state.can_proceed_to[0];
    if (!next) {
      const serverMsgs = (draft.wizard_state.validation_errors ?? []).filter(Boolean);
      if (serverMsgs.length > 0) {
        setStepError(serverMsgs.join(" "));
      } else if (draft.wizard_state.next_required_action === "ingest_steps") {
        setStepError("Сначала распознайте и получите хотя бы одну позицию в заказе.");
      } else {
        setStepError("Нельзя перейти дальше — проверьте список ступеней и повторите.");
      }
      return;
    }
    await proceedFromInputStep(next);
  };

  const handleFinishMarches = async () => {
    setStepError(null);
    if (state.pendingBatchReview) {
      setStepError("Сначала подтвердите список текущего источника — нажмите «Список верен».");
      return;
    }
    if (!currentDraft?.wizard_state || !currentDraft.draft_id) {
      setStepError("Не удалось загрузить данные черновика. Обновите страницу или начните заново.");
      return;
    }

    const draft = currentDraft;
    const next = draft.wizard_state.can_proceed_to[0];
    if (!next) {
      const serverMsgs = (draft.wizard_state.validation_errors ?? []).filter(Boolean);
      if (serverMsgs.length > 0) {
        setStepError(serverMsgs.join(" "));
      } else if (draft.wizard_state.next_required_action === "ingest_marches") {
        setStepError("Сначала распознайте и получите хотя бы одну позицию в заказе.");
      } else {
        setStepError("Нельзя перейти дальше — проверьте список маршей и повторите.");
      }
      return;
    }
    await proceedFromInputStep(next);
  };
const handleFinishBridgePiles = async () => {
    setStepError(null);
    if (state.pendingBatchReview) {
      setStepError("Сначала подтвердите список текущего источника — нажмите «Список верен».");
      return;
    }
    if (!currentDraft?.wizard_state || !currentDraft.draft_id) {
      setStepError("Не удалось загрузить данные черновика. Обновите страницу или начните заново.");
      return;
    }

    const draft = currentDraft;
    const next = draft.wizard_state.can_proceed_to[0];
    if (!next) {
      const serverMsgs = (draft.wizard_state.validation_errors ?? []).filter(Boolean);
      if (serverMsgs.length > 0) {
        setStepError(serverMsgs.join(" "));
      } else if (draft.wizard_state.next_required_action === "ingest_bridge_piles") {
        setStepError("Сначала распознайте и получите хотя бы одну позицию в заказе.");
      } else {
        setStepError("Нельзя перейти дальше — проверьте список мостовых свай и повторите.");
      }
      return;
    }
    await proceedFromInputStep(next);
  };

  const handleFinishFbs = async () => {
    setStepError(null);
    if (state.pendingBatchReview) {
      setStepError("Сначала подтвердите список текущего источника — нажмите «Список верен».");
      return;
    }
    if (!currentDraft?.wizard_state || !currentDraft.draft_id) {
      setStepError("Не удалось загрузить данные черновика. Обновите страницу или начните заново.");
      return;
    }

    const draft = currentDraft;
    const next = draft.wizard_state.can_proceed_to[0];
    if (!next) {
      const serverMsgs = (draft.wizard_state.validation_errors ?? []).filter(Boolean);
      if (serverMsgs.length > 0) {
        setStepError(serverMsgs.join(" "));
      } else if (draft.wizard_state.next_required_action === "ingest_fbs") {
        setStepError("Сначала распознайте и получите хотя бы одну позицию в заказе.");
      } else {
        setStepError("Нельзя перейти дальше — проверьте список ФБС и повторите.");
      }
      return;
    }
    await proceedFromInputStep(next);
  };

  const handleApplyGradeToAll = async (grade: string) => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    try {
      const gradesMutation = isFbsFlow
        ? updateFbsGradesMutation
        : isBridgePileFlow
        ? updateBridgePileGradesMutation
        : isMarchFlow
          ? updateMarchGradesMutation
          : updatePileGradesMutation;
      await gradesMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        concreteGrade: grade,
      });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleLineGradeChange = async (lineIndex: number, grade: string) => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    // Preview shows sealed ∪ current. Grade re-ingest must send only unsealed lines
    // (build*LinesFromOrderData skips sealed) with append mode so sealed batches stay intact.
    const hasSealedLines = (currentDraft.order_data ?? []).some(
      (item) => String(item.append_batch_id ?? "").trim().length > 0,
    );
    const updateMode = hasSealedLines ? "append" : "replace";
    try {
      if (isFbsFlow) {
        const rows = buildFbsPreviewRows(currentDraft);
        if (rows[lineIndex]?.sealed) {
          return;
        }
        const updated = rows.map((row, idx) => (idx === lineIndex ? { ...row, concrete_grade: grade } : row));
        const text = buildFbsLinesFromOrderData(updated);
        await updateFbsMutation.mutateAsync({
          draftId: currentDraft.draft_id,
          text,
          image: null,
          mode: updateMode,
        });
        return;
      }
      if (isBridgePileFlow) {
        const rows = buildBridgePilePreviewRows(currentDraft);
        if (rows[lineIndex]?.sealed) {
          return;
        }
        const updated = rows.map((row, idx) => (idx === lineIndex ? { ...row, concrete_grade: grade } : row));
        const text = buildBridgePileLinesFromOrderData(updated);
        await updateBridgePilesMutation.mutateAsync({
          draftId: currentDraft.draft_id,
          text,
          image: null,
          mode: updateMode,
        });
        return;
      }
      if (isMarchFlow) {
        const rows = buildMarchPreviewRows(currentDraft);
        if (lineIndex < 0 || lineIndex >= rows.length || rows[lineIndex]?.sealed) {
          return;
        }
        const updated = rows.map((row, idx) => (idx === lineIndex ? { ...row, concrete_grade: grade } : row));
        const text = buildMarchLinesFromOrderData(updated);
        await updateMarchesMutation.mutateAsync({
          draftId: currentDraft.draft_id,
          text,
          image: null,
          mode: updateMode,
        });
        return;
      }

      const rows = buildPilePreviewRows(currentDraft);
      if (lineIndex < 0 || lineIndex >= rows.length || rows[lineIndex]?.sealed) {
        return;
      }
      const updated = rows.map((row, idx) => (idx === lineIndex ? { ...row, concrete_grade: grade } : row));
      const text = buildPileLinesFromOrderData(updated);
      await updatePilesMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        text,
        image: null,
        mode: updateMode,
      });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleApplyWidePlates = async () => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setWidePlateError(null);
    try {
      const flushText = buildMergedFlushText({
        hasStarted: multiPage.hasStarted,
        pages: multiPage.pages,
        activePageId: multiPage.activeId,
        editorText: state.batchReviewText,
        singlePageText: state.batchReviewText,
      });
      await flushThenResolveWidePlates({
        draftId: currentDraft.draft_id,
        flushText,
        persistedText: (currentDraft.metadata.input_text || currentDraft.metadata.normalized_text || "").trim(),
        liveLines: liveWidePlateLines(flushText),
        decisionsById: state.widePlateActions,
        currentWideLines: currentDraft.metadata.wide_plate_lines ?? [],
        updateInput: (payload) => updateInputMutation.mutateAsync(payload),
        resolveWidePlates: (payload) => resolveWidePlatesMutation.mutateAsync(payload),
      });
    } catch (error) {
      setWidePlateError(getErrorMessage(error));
    }
  };

  const handleApplyUnpricedPlates = async () => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setUnpricedPlateError(null);
    try {
      await resolveUnpricedPlatesMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        decisions: (currentDraft.metadata.unpriced_plate_lines ?? []).map((item) => {
          const decision = state.unpricedPlateActions[item.id];
          const fallbackLoad = item.replacements[0]?.load_code ?? null;
          return {
            lineId: item.id,
            sourceLine: item.line,
            action: decision?.action ?? (fallbackLoad != null ? "replace_load" : "exclude"),
            loadCode: decision?.loadCode ?? fallbackLoad,
          };
        }),
      });
    } catch (error) {
      setUnpricedPlateError(getErrorMessage(error));
    }
  };

  const handleApplyInvalidWidths = async () => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setInvalidWidthError(null);
    try {
      await resolveInvalidWidthsMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        decisions: (currentDraft.metadata.invalid_width_lines ?? []).map((item) => {
          const decision = state.invalidWidthActions[item.id];
          const upper = item.replacements.reduce<(typeof item.replacements)[number] | null>(
            (best, repl) => (best == null || repl.width_mm > best.width_mm ? repl : best),
            null,
          );
          return {
            lineId: item.id,
            sourceLine: item.line,
            action: decision?.action ?? (upper != null ? "replace_width" : "exclude"),
            widthMm: decision?.widthMm ?? upper?.width_mm ?? null,
          };
        }),
      });
    } catch (error) {
      setInvalidWidthError(getErrorMessage(error));
    }
  };

  const handleClientSubmit = async (payload: {
    managerId: number;
    clientName: string;
    conditionsMode: "standard" | "custom";
    deliveryConditions: string;
    paymentConditions: string;
  }) => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    dispatch({ type: "set-manager", managerId: payload.managerId });
    dispatch({
      type: "set-client-form",
      payload: {
        clientName: payload.clientName,
        conditionsMode: payload.conditionsMode,
        deliveryConditions: payload.deliveryConditions,
        paymentConditions: payload.paymentConditions,
      },
    });
    try {
      await updateMetaMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        managerId: payload.managerId,
        clientName: payload.clientName,
        conditionsMode: payload.conditionsMode,
        deliveryConditions: payload.deliveryConditions,
        paymentConditions: payload.paymentConditions,
      });
      const calculated = await calculateMutation.mutateAsync(currentDraft.draft_id);
      dispatch({ type: "hydrate-draft", payload: calculated });
      const ws = calculated.wizard_state;
      if (ws.current_step !== "result") {
        const msgs = (ws.validation_errors ?? []).filter(Boolean);
        setStepError(
          msgs.length > 0
            ? msgs.join(" ")
            : "Расчёт не завершён. Проверьте клиента, менеджера и условия.",
        );
        return;
      }
      dispatch({
        type: "set-step",
        step: ws.current_step,
      });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleGenerateFiles = async () => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    try {
      await generateFilesMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        fileTypes: isSimpleKpDraft ? ["pdf", "xlsx"] : undefined,
      });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleGenerateSchema = async () => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    try {
      await generateSchemaMutation.mutateAsync(currentDraft.draft_id);
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleSave = async (payload: { mode: "database" | "archive" | "skip"; executionTermsInput: string }) => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    try {
      await saveDraftMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        mode: payload.mode,
        executionTermsInput: payload.executionTermsInput,
      });
      // Archive/database save ends the sticky source-image session.
      sourceImageQueue.clear();
      const resumeKpId =
        currentDraft.metadata.resume_kp_id ?? state.lastDraft?.metadata.resume_kp_id ?? null;
      if (resumeKpId != null && Number(resumeKpId) > 0) {
        await queryClient.invalidateQueries({ queryKey: archiveKeys.all });
        setStepError(null);
        setWidePlateError(null);
        multiPage.reset();
        clearRecognizedImagePreview();
        dispatch({ type: "reset" });
        navigate("/archive");
      }
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleDiscountSubmit = async (discountPercent: number) => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    try {
      await updateMetaMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        discountPercent,
      });
      const calculated = await calculateMutation.mutateAsync(currentDraft.draft_id);
      dispatch({ type: "hydrate-draft", payload: calculated });
      const msgs = (calculated.wizard_state.validation_errors ?? []).filter(Boolean);
      if (msgs.length > 0) {
        setStepError(msgs.join(" "));
      }
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleLogisticsCostSubmit = async (logisticsCost: number) => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    try {
      await updateMetaMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        logisticsCost,
      });
      const calculated = await calculateMutation.mutateAsync(currentDraft.draft_id);
      dispatch({ type: "hydrate-draft", payload: calculated });
      const msgs = (calculated.wizard_state.validation_errors ?? []).filter(Boolean);
      if (msgs.length > 0) {
        setStepError(msgs.join(" "));
      }
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handlePileDeliverySubmit = async (payload: {
    pileLogisticsCost?: number;
    pileTripOverrides?: Record<string, number>;
  }) => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    try {
      await updateMetaMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        pileLogisticsCost: payload.pileLogisticsCost,
        pileTripOverrides: payload.pileTripOverrides,
      });
      const calculated = await calculateMutation.mutateAsync(currentDraft.draft_id);
      dispatch({ type: "hydrate-draft", payload: calculated });
      const msgs = (calculated.wizard_state.validation_errors ?? []).filter(Boolean);
      if (msgs.length > 0) {
        setStepError(msgs.join(" "));
      }
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleCreateNewOffer = () => {
    setStepError(null);
    setWidePlateError(null);
    sourceImageQueue.clear();
    multiPage.reset();
    clearRecognizedImagePreview();
    dispatch({ type: "reset" });
  };

  const handleAddOtherNomenclature = () => {
    setStepError(null);
    sourceImageQueue.clear();
    multiPage.reset();
    clearRecognizedImagePreview();
    dispatch({ type: "start-append-cycle" });
  };

  const handleCancelAppendPick = () => {
    setStepError(null);
    dispatch({ type: "cancel-append-pick" });
  };

  const handleAppendProductTypeSelect = async (nextProductType: ProductType) => {
    if (!state.draftId) {
      dispatch({ type: "set-product-type", productType: nextProductType });
      return;
    }
    setStepError(null);
    try {
      const draft = await startAppendCycleMutation.mutateAsync({
        draftId: state.draftId,
        productType: nextProductType,
      });
      dispatch({ type: "set-product-type", productType: nextProductType });
      dispatch({ type: "hydrate-draft", payload: draft });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleUndoLastBatch = async () => {
    if (!state.draftId) {
      return;
    }
    setStepError(null);
    try {
      const draft = await undoLastAppendBatchMutation.mutateAsync(state.draftId);
      dispatch({ type: "hydrate-draft", payload: draft });
      dispatch({ type: "set-step", step: "result" });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const clearLineUndo = useCallback(() => {
    if (lineUndoTimerRef.current) {
      clearTimeout(lineUndoTimerRef.current);
      lineUndoTimerRef.current = null;
    }
    lineUndoRef.current = null;
    setLineUndoToastMessage(null);
  }, []);

  const armLineUndo = useCallback(
    (
      op:
        | { kind: "qty"; lineId: string; qty: number }
        | { kind: "restore"; index: number; lines: Record<string, unknown>[]; replaceLineIds: string[] },
      message: string,
    ) => {
      clearLineUndo();
      lineUndoRef.current = op;
      setLineUndoToastMessage(message);
      lineUndoTimerRef.current = setTimeout(() => {
        lineUndoRef.current = null;
        setLineUndoToastMessage(null);
        lineUndoTimerRef.current = null;
      }, LINE_UNDO_TOAST_MS);
    },
    [clearLineUndo],
  );

  useEffect(() => {
    clearLineUndo();
    setLineRowError(null);
  }, [state.currentStep, clearLineUndo]);

  const handleUndoLineOp = async () => {
    const op = lineUndoRef.current;
    if (!state.draftId || !op) {
      return;
    }
    setLineRowError(null);
    try {
      const draft =
        op.kind === "qty"
          ? await patchDraftLineMutation.mutateAsync({
              draftId: state.draftId,
              lineId: op.lineId,
              payload: { qty: op.qty },
            })
          : await restoreDraftLinesMutation.mutateAsync({
              draftId: state.draftId,
              payload: {
                index: op.index,
                lines: op.lines,
                replace_line_ids: op.replaceLineIds,
              },
            });
      dispatch({ type: "hydrate-draft", payload: draft });
      clearLineUndo();
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleSaveLine = async (lineId: string, payload: LineSavePayload) => {
    if (!state.draftId || !currentDraft) {
      return;
    }
    const order = currentDraft.order_data ?? [];
    const index = order.findIndex((item) => String(item.line_id ?? "") === lineId);
    if (index < 0) {
      return;
    }
    const snapshot = { ...order[index] };
    const oldQty = Number(snapshot.qty) || 0;
    setLineRowError(null);
    setStepError(null);
    try {
      const body =
        payload.sourceText !== undefined
          ? { source_text: payload.sourceText }
          : { qty: payload.qty };
      const draft = await patchDraftLineMutation.mutateAsync({
        draftId: state.draftId,
        lineId,
        payload: body,
      });
      dispatch({ type: "hydrate-draft", payload: draft });
      if (payload.sourceText !== undefined) {
        const kept = new Set(
          order.map((item) => String(item.line_id ?? "")).filter((id) => id && id !== lineId),
        );
        const replaceLineIds = (draft.order_data ?? [])
          .map((item) => String(item.line_id ?? ""))
          .filter((id) => id && !kept.has(id));
        armLineUndo(
          { kind: "restore", index, lines: [snapshot], replaceLineIds },
          "Строка изменена",
        );
      } else if (payload.qty !== undefined) {
        armLineUndo({ kind: "qty", lineId, qty: oldQty }, "Количество изменено");
      }
    } catch (error) {
      setLineRowError({ lineId, message: getErrorMessage(error) });
    }
  };

  const handleDeleteLine = async (lineId: string) => {
    if (!state.draftId || !currentDraft) {
      return;
    }
    const order = currentDraft.order_data ?? [];
    const index = order.findIndex((item) => String(item.line_id ?? "") === lineId);
    if (index < 0) {
      return;
    }
    const snapshot = { ...order[index] };
    setStepError(null);
    setLineRowError(null);
    try {
      const draft = await deleteDraftLineMutation.mutateAsync({ draftId: state.draftId, lineId });
      dispatch({ type: "hydrate-draft", payload: draft });
      armLineUndo({ kind: "restore", index, lines: [snapshot], replaceLineIds: [] }, "Строка удалена");
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const hasUnresolvedWidePlates =
    Boolean(currentDraft?.metadata.wide_plate_lines?.length) && !currentDraft?.metadata.wide_plates_resolved;
  const hasUnresolvedUnpricedPlates =
    Boolean(currentDraft?.metadata.unpriced_plate_lines?.length) &&
    !currentDraft?.metadata.unpriced_plates_resolved;
  const hasUnresolvedInvalidWidths =
    Boolean(currentDraft?.metadata.invalid_width_lines?.length) &&
    !currentDraft?.metadata.invalid_widths_resolved;

  const canNavigateToStep = (step: WizardStepId): boolean => {
    if (
      isInputStepBlockedWithoutAppendCycle({
        currentStep: state.currentStep,
        targetStep: step,
        inputStep,
        draftWizardStep: currentDraft?.wizard_state?.current_step as WizardStepId | undefined,
      })
    ) {
      return false;
    }
    if (step === inputStep) {
      return true;
    }
    if (!currentDraft?.wizard_state) {
      return false;
    }
    const ws = currentDraft.wizard_state;
    if (step === state.currentStep) {
      return true;
    }
    if (stepIndex(step) < stepIndex(ws.current_step as WizardStepId)) {
      return true;
    }
    return ws.can_proceed_to.includes(step);
  };

  const handleSidebarStepClick = (step: WizardStepId) => {
    if (!canNavigateToStep(step)) {
      if (!currentDraft && step !== inputStep) {
        setStepError(
          isMarchFlow
            ? "Сначала распознайте и обработайте список маршей."
            : isStepFlow
            ? "Сначала распознайте и обработайте список ступеней."
            : isPileFlow
              ? "Сначала распознайте и обработайте список свай."
              : "Сначала распознайте и обработайте список плит.",
        );
        return;
      }
      if (
        step === "client" &&
        !isSimpleProductFlow &&
        (hasUnresolvedWidePlates ||
          hasUnresolvedInvalidWidths ||
          hasUnresolvedUnpricedPlates ||
          state.pendingBatchReview)
      ) {
        setStepError(
          state.pendingBatchReview
            ? "Сначала подтвердите список текущего источника — «Список верен»."
            : hasUnresolvedWidePlates
              ? "Сначала примите решение по позициям шире стандартной."
              : hasUnresolvedInvalidWidths
                ? "Нестандартная ширина: замените на заводской рез или исключите позицию."
                : "Сначала примите решение по позициям без цены в прайсе.",
        );
        return;
      }
      if (step === "client" && isSimpleProductFlow && state.pendingBatchReview) {
        setStepError("Сначала подтвердите список текущего источника — «Список верен».");
        return;
      }
      return;
    }
    setStepError(null);
    dispatch({ type: "set-step", step });
  };

  const pageProgressLabel =
    multiPage.hasStarted && multiPage.pages.length > 0
      ? `Распознано ${multiPage.progress.recognized}/${multiPage.progress.total}`
      : null;
  const activeReviewPage = multiPage.activePage;
  const multiPendingReview = multiPage.hasStarted && !multiPage.allConfirmed;
  const pendingBatchReview = multiPendingReview || state.pendingBatchReview;
  const reviewImageUrl =
    activeReviewPage?.recognizedImageUrl ??
    activeReviewPage?.previewUrl ??
    recognizedImagePreview?.url ??
    null;
  const reviewImageName = activeReviewPage?.name ?? recognizedImagePreview?.name ?? null;
  const reviewBatchText =
    multiPage.hasStarted && activeReviewPage
      ? activeReviewPage.batchReviewText
      : state.batchReviewText;
  const canConfirmActivePage = multiPage.hasStarted
    ? activeReviewPage?.status === "ready"
    : true;
  const isRecognizingMulti = multiPage.isRecognizing;
  const isRerecognizing =
    (activeReviewPage?.status === "running" && Boolean(activeReviewPage.file)) ||
    (!multiPage.hasStarted && updateInputMutation.isPending && Boolean(recognizedImagePreview?.file));

  const multiPageStepProps = {
    pages: multiPage.pages,
    activePageId: multiPage.activeId,
    softCapMessage: multiPage.softCapMessage,
    pageProgressLabel,
    recognitionStarted: multiPage.hasStarted,
    onAddFiles: handleAddFiles,
    onRemovePage: handleRemovePage,
    onSelectPage: handleSelectPage,
    onPrevPage: handlePrevPage,
    onNextPage: handleNextPage,
    canConfirmActivePage,
  };

  const lineRowHandlers: LineRowHandlers = {
    onSaveLine: (lineId, payload) => void handleSaveLine(lineId, payload),
    onDeleteLine: (lineId) => void handleDeleteLine(lineId),
    undoToast: lineUndoToastMessage
      ? { message: lineUndoToastMessage, onUndo: () => void handleUndoLineOp() }
      : null,
    rowError: lineRowError,
  };

  const currentStepContent =
    state.currentStep === "fbs" ? (
      <FbsInputStep
        draft={currentDraft}
        pendingBatchReview={pendingBatchReview}
        sourceText={state.sourceText}
        batchReviewText={multiPage.hasStarted ? reviewBatchText : state.batchReviewText}
        normalizedText={state.normalizedText}
        {...multiPageStepProps}
        recognizedImageUrl={reviewImageUrl}
        recognizedImageName={reviewImageName}
        sourceQueue={sourceImageQueue.items}
        errorMessage={stepError}
        isRecognizing={isRecognizingMulti || createDraftMutation.isPending || updateFbsMutation.isPending}
        isAiProcessing={applyAiFbsMutation.isPending}
        isUpdatingGrades={updateFbsGradesMutation.isPending}
        isConfirmingBatch={updateFbsMutation.isPending}
        isProceeding={false}
        aiInstruction={aiInstruction}
        onAiInstructionChange={setAiInstruction}
        onApplyAi={() => void handleApplyAi()}
        onTextChange={handleSourceTextChange}
        onBatchReviewTextChange={(value) => {
          dispatch({ type: "set-batch-review-text", text: value });
          if (multiPage.activeId) {
            multiPage.updatePageText(multiPage.activeId, value);
          }
        }}
        
        onRecognize={handleRecognize}
        onRerecognize={() => void handleRerecognize()}
        isRerecognizing={isRerecognizing}
        onConfirmBatch={() => void handleConfirmBatch()}
        onFinishFbs={() => void handleFinishFbs()}
        onApplyGradeToAll={(grade) => void handleApplyGradeToAll(grade)}
        onLineGradeChange={(lineIndex, grade) => void handleLineGradeChange(lineIndex, grade)}
        onReset={handleCreateNewOffer}
        lineRowHandlers={lineRowHandlers}
      />
    ) :     state.currentStep === "bridge_piles" ? (
      <BridgePileInputStep
        draft={currentDraft}
        pendingBatchReview={pendingBatchReview}
        sourceText={state.sourceText}
        batchReviewText={multiPage.hasStarted ? reviewBatchText : state.batchReviewText}
        normalizedText={state.normalizedText}
        {...multiPageStepProps}
        recognizedImageUrl={reviewImageUrl}
        recognizedImageName={reviewImageName}
        sourceQueue={sourceImageQueue.items}
        errorMessage={stepError}
        isRecognizing={isRecognizingMulti || createDraftMutation.isPending || updateBridgePilesMutation.isPending}
        isAiProcessing={applyAiBridgePilesMutation.isPending}
        isUpdatingGrades={updateBridgePileGradesMutation.isPending}
        isConfirmingBatch={updateBridgePilesMutation.isPending}
        isProceeding={false}
        aiInstruction={aiInstruction}
        onAiInstructionChange={setAiInstruction}
        onApplyAi={() => void handleApplyAi()}
        onTextChange={handleSourceTextChange}
        onBatchReviewTextChange={(value) => {
          dispatch({ type: "set-batch-review-text", text: value });
          if (multiPage.activeId) {
            multiPage.updatePageText(multiPage.activeId, value);
          }
        }}
        
        onRecognize={handleRecognize}
        onRerecognize={() => void handleRerecognize()}
        isRerecognizing={isRerecognizing}
        onConfirmBatch={() => void handleConfirmBatch()}
        onFinishBridgePiles={() => void handleFinishBridgePiles()}
        onApplyGradeToAll={(grade) => void handleApplyGradeToAll(grade)}
        onLineGradeChange={(lineIndex, grade) => void handleLineGradeChange(lineIndex, grade)}
        onReset={handleCreateNewOffer}
        lineRowHandlers={lineRowHandlers}
      />
    ) : state.currentStep === "marches" ? (
      <MarchInputStep
        draft={currentDraft}
        pendingBatchReview={pendingBatchReview}
        sourceText={state.sourceText}
        batchReviewText={multiPage.hasStarted ? reviewBatchText : state.batchReviewText}
        normalizedText={state.normalizedText}
        {...multiPageStepProps}
        recognizedImageUrl={reviewImageUrl}
        recognizedImageName={reviewImageName}
        sourceQueue={sourceImageQueue.items}
        errorMessage={stepError}
        isRecognizing={isRecognizingMulti || createDraftMutation.isPending || updateMarchesMutation.isPending}
        isAiProcessing={applyAiMarchesMutation.isPending}
        isUpdatingGrades={updateMarchGradesMutation.isPending}
        isConfirmingBatch={updateMarchesMutation.isPending}
        isProceeding={false}
        aiInstruction={aiInstruction}
        onAiInstructionChange={setAiInstruction}
        onApplyAi={() => void handleApplyAi()}
        onTextChange={handleSourceTextChange}
        onBatchReviewTextChange={(value) => {
          dispatch({ type: "set-batch-review-text", text: value });
          if (multiPage.activeId) {
            multiPage.updatePageText(multiPage.activeId, value);
          }
        }}
        
        onRecognize={handleRecognize}
        onRerecognize={() => void handleRerecognize()}
        isRerecognizing={isRerecognizing}
        onConfirmBatch={() => void handleConfirmBatch()}
        onFinishMarches={() => void handleFinishMarches()}
        onApplyGradeToAll={(grade) => void handleApplyGradeToAll(grade)}
        onLineGradeChange={(lineIndex, grade) => void handleLineGradeChange(lineIndex, grade)}
        onReset={handleCreateNewOffer}
        lineRowHandlers={lineRowHandlers}
      />
    ) : state.currentStep === "steps" ? (
      <StepInputStep
        draft={currentDraft}
        pendingBatchReview={pendingBatchReview}
        sourceText={state.sourceText}
        batchReviewText={multiPage.hasStarted ? reviewBatchText : state.batchReviewText}
        normalizedText={state.normalizedText}
        {...multiPageStepProps}
        recognizedImageUrl={reviewImageUrl}
        recognizedImageName={reviewImageName}
        sourceQueue={sourceImageQueue.items}
        errorMessage={stepError}
        isRecognizing={isRecognizingMulti || createDraftMutation.isPending || updateStepsMutation.isPending}
        isAiProcessing={applyAiStepsMutation.isPending}
        isConfirmingBatch={updateStepsMutation.isPending}
        isProceeding={false}
        aiInstruction={aiInstruction}
        onAiInstructionChange={setAiInstruction}
        onApplyAi={() => void handleApplyAi()}
        onTextChange={handleSourceTextChange}
        onBatchReviewTextChange={(value) => {
          dispatch({ type: "set-batch-review-text", text: value });
          if (multiPage.activeId) {
            multiPage.updatePageText(multiPage.activeId, value);
          }
        }}
        
        onRecognize={handleRecognize}
        onRerecognize={() => void handleRerecognize()}
        isRerecognizing={isRerecognizing}
        onConfirmBatch={() => void handleConfirmBatch()}
        onFinishSteps={() => void handleFinishSteps()}
        onReset={handleCreateNewOffer}
        lineRowHandlers={lineRowHandlers}
      />
    ) : state.currentStep === "piles" ? (
      <PileInputStep
        draft={currentDraft}
        pendingBatchReview={pendingBatchReview}
        sourceText={state.sourceText}
        batchReviewText={multiPage.hasStarted ? reviewBatchText : state.batchReviewText}
        normalizedText={state.normalizedText}
        {...multiPageStepProps}
        recognizedImageUrl={reviewImageUrl}
        recognizedImageName={reviewImageName}
        sourceQueue={sourceImageQueue.items}
        errorMessage={stepError}
        isRecognizing={isRecognizingMulti || createDraftMutation.isPending || updatePilesMutation.isPending}
        isAiProcessing={applyAiPilesMutation.isPending}
        isUpdatingGrades={updatePileGradesMutation.isPending}
        isConfirmingBatch={updatePilesMutation.isPending}
        isProceeding={false}
        aiInstruction={aiInstruction}
        onAiInstructionChange={setAiInstruction}
        onApplyAi={() => void handleApplyAi()}
        onTextChange={handleSourceTextChange}
        onBatchReviewTextChange={(value) => {
          dispatch({ type: "set-batch-review-text", text: value });
          if (multiPage.activeId) {
            multiPage.updatePageText(multiPage.activeId, value);
          }
        }}
        
        onRecognize={handleRecognize}
        onRerecognize={() => void handleRerecognize()}
        isRerecognizing={isRerecognizing}
        onConfirmBatch={() => void handleConfirmBatch()}
        onFinishPiles={() => void handleFinishPiles()}
        onApplyGradeToAll={(grade) => void handleApplyGradeToAll(grade)}
        onLineGradeChange={(lineIndex, grade) => void handleLineGradeChange(lineIndex, grade)}
        onReset={handleCreateNewOffer}
        lineRowHandlers={lineRowHandlers}
      />
    ) : state.currentStep === "plates" ? (
      <PlateInputStep
        draft={currentDraft}
        pendingBatchReview={pendingBatchReview}
        sourceText={state.sourceText}
        batchReviewText={multiPage.hasStarted ? reviewBatchText : state.batchReviewText}
        normalizedText={state.normalizedText}
        {...multiPageStepProps}
        recognizedImageUrl={reviewImageUrl}
        recognizedImageName={reviewImageName}
        sourceQueue={sourceImageQueue.items}
        errorMessage={stepError}
        widePlateErrorMessage={widePlateError}
        unpricedPlateErrorMessage={unpricedPlateError}
        invalidWidthErrorMessage={invalidWidthError}
        isRecognizing={isRecognizingMulti || createDraftMutation.isPending || updatePlatesMutation.isPending}
        isAiProcessing={applyAiPlatesMutation.isPending}
        isResolvingWidePlates={resolveWidePlatesMutation.isPending}
        isResolvingUnpricedPlates={resolveUnpricedPlatesMutation.isPending}
        isResolvingInvalidWidths={resolveInvalidWidthsMutation.isPending}
        isConfirmingBatch={updatePlatesMutation.isPending}
        isProceeding={false}
        widePlateDecisions={state.widePlateActions}
        unpricedPlateDecisions={state.unpricedPlateActions}
        invalidWidthDecisions={state.invalidWidthActions}
        aiInstruction={aiInstruction}
        onAiInstructionChange={setAiInstruction}
        onApplyAi={() => void handleApplyAi()}
        onTextChange={handleSourceTextChange}
        onBatchReviewTextChange={(value) => {
          dispatch({ type: "set-batch-review-text", text: value });
          if (multiPage.activeId) {
            multiPage.updatePageText(multiPage.activeId, value);
          }
        }}
        
        onRecognize={handleRecognize}
        onRerecognize={() => void handleRerecognize()}
        isRerecognizing={isRerecognizing}
        onConfirmBatch={() => void handleConfirmBatch()}
        onFinishPlates={() => void handleFinishPlates()}
        onWidePlateDecisionChange={(lineId, action, replacementText) =>
          dispatch({ type: "set-wide-action", lineId, action, replacementText })
        }
        onApplyWidePlates={() => void handleApplyWidePlates()}
        onUnpricedPlateDecisionChange={(lineId, action, loadCode) =>
          dispatch({ type: "set-unpriced-action", lineId, action, loadCode })
        }
        onApplyUnpricedPlates={() => void handleApplyUnpricedPlates()}
        onInvalidWidthDecisionChange={(lineId, action, widthMm) =>
          dispatch({ type: "set-invalid-width-action", lineId, action, widthMm })
        }
        onApplyInvalidWidths={() => void handleApplyInvalidWidths()}
        onReset={handleCreateNewOffer}
        lineRowHandlers={lineRowHandlers}
      />
    ) : state.currentStep === "client" ? (
      <ClientConditionsStep
        managers={managers}
        selectedManagerId={state.managerId}
        defaultValues={{
          clientName: state.clientName,
          conditionsMode: state.conditionsMode,
          deliveryConditions: state.deliveryConditions,
          paymentConditions: state.paymentConditions,
        }}
        errorMessage={stepError}
        isPending={updateMetaMutation.isPending || calculateMutation.isPending}
        onBack={() => dispatch({ type: "set-step", step: inputStep })}
        onManagerChange={(managerId) => dispatch({ type: "set-manager", managerId })}
        onSubmit={handleClientSubmit}
      />
    ) : state.currentStep === "result" && currentDraft ? (
      <CalculationResultStep
        draft={currentDraft}
        breakdownTables={breakdownQuery.data?.items ?? []}
        isBreakdownLoading={breakdownQuery.isPending || breakdownQuery.isFetching}
        errorMessage={stepError}
        isPileDraft={isPileDraft}
        isStepDraft={isStepDraft}
        isMarchDraft={isMarchDraft}
        isBridgePileDraft={isBridgePileDraft}
        isFbsDraft={isFbsDraft}
        isSimpleKpDraft={isSimpleKpDraft}
        isGeneratingFiles={generateFilesMutation.isPending}
        isGeneratingSchema={generateSchemaMutation.isPending}
        isSaving={saveDraftMutation.isPending}
        lastSaveResult={state.lastSaveResult}
        onBack={
          skipClient ? undefined : () => dispatch({ type: "set-step", step: "client" })
        }
        onCreateNew={handleCreateNewOffer}
        onGenerateFiles={handleGenerateFiles}
        onGenerateSchema={handleGenerateSchema}
        onSave={handleSave}
        isUpdatingDiscount={updateMetaMutation.isPending || calculateMutation.isPending}
        onDiscountSubmit={handleDiscountSubmit}
        onLogisticsCostSubmit={handleLogisticsCostSubmit}
        onPileDeliverySubmit={handlePileDeliverySubmit}
        onAddOtherNomenclature={handleAddOtherNomenclature}
        onUndoLastBatch={() => void handleUndoLastBatch()}
        onDeleteLine={(lineId) => void handleDeleteLine(lineId)}
        onSaveLine={(lineId, payload) => void handleSaveLine(lineId, payload)}
        lineUndoToast={lineRowHandlers.undoToast}
        lineRowError={lineRowError}
      />
    ) : null;

  const renderDraftStepFallback = () => {
    if (!state.draftId) {
      return (
        <div style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}>
          <Spinner /> Восстанавливаю шаг мастера…
        </div>
      );
    }
    if (draftQuery.isPending || draftQuery.isFetching) {
      return (
        <div style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}>
          <Spinner /> Загружаю черновик…
        </div>
      );
    }
    if (draftQuery.isError) {
      return (
        <div style={{ display: "grid", gap: "1rem" }}>
          <Alert tone="error">{getErrorMessage(draftQuery.error)}</Alert>
          <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
            <Button type="button" onClick={() => void draftQuery.refetch()}>
              Повторить
            </Button>
            <Button type="button" variant="secondary" onClick={handleCreateNewOffer}>
              Начать заново
            </Button>
          </div>
        </div>
      );
    }
    return (
      <div style={{ display: "grid", gap: "1rem" }}>
        <Alert tone="error">
          Черновик недоступен. Проверьте сеть и авторизацию или начните новое коммерческое предложение.
        </Alert>
        <Button type="button" onClick={handleCreateNewOffer}>
          Начать заново
        </Button>
      </div>
    );
  };

  const resolvedStepContent =
    currentStepContent ??
    (state.currentStep === "result"
      ? renderDraftStepFallback()
      : (
          <div style={{ display: "grid", gap: "1rem" }}>
            <Alert tone="error">Некорректное состояние мастера.</Alert>
            <Button type="button" onClick={handleCreateNewOffer}>
              Сбросить и начать сначала
            </Button>
          </div>
        ));

  const appendOrderLines =
    currentDraft?.order_data ?? state.lastDraft?.order_data ?? [];
  const appendSelectedProductTypes = (() => {
    const seen = new Set<ProductType>();
    const ordered: ProductType[] = [];
    for (const line of appendOrderLines) {
      const raw = String(line.product_type ?? "").trim();
      if (!raw || seen.has(raw as ProductType)) {
        continue;
      }
      // Only known product types appear as "already in KP".
      if (
        raw === "plates" ||
        raw === "piles" ||
        raw === "steps" ||
        raw === "marches" ||
        raw === "bridge_piles" ||
        raw === "fbs"
      ) {
        seen.add(raw);
        ordered.push(raw);
      }
    }
    return ordered;
  })();
  const appendManagerName =
    state.lastDraft?.metadata.manager_name?.trim() ||
    currentDraft?.metadata.manager_name?.trim() ||
    managers.find((manager) => manager.id === state.managerId)?.fio ||
    "";

  return (
    <div style={{ display: "grid", gap: "1.25rem" }}>
      {managersQuery.error && <Alert tone="error">{getErrorMessage(managersQuery.error)}</Alert>}
      {state.isPickingProductType ? (
        <div style={{ display: "grid", gap: "1rem" }}>
          {stepError ? <Alert tone="error">{stepError}</Alert> : null}
          {state.discountPercent > 0 ? (
            <Alert tone="info">Скидка: {state.discountPercent}%</Alert>
          ) : null}
          <ProductTypePicker
            mode="append"
            selectedProductTypes={appendSelectedProductTypes}
            orderLines={appendOrderLines}
            managerName={appendManagerName}
            clientName={state.clientName}
            onSelect={(nextType) => void handleAppendProductTypeSelect(nextType)}
            onBackToResult={handleCancelAppendPick}
          />
        </div>
      ) : (
        <div className="wizard-shell">
          <div className="wizard-main">{resolvedStepContent}</div>

          <aside className="wizard-sidebar">
            <div className="wizard-sidebar__inner">
              <WizardProgress
                productType={productType}
                currentStep={state.currentStep}
                onStepClick={handleSidebarStepClick}
                canNavigateToStep={canNavigateToStep}
                skipClient={skipClient}
              />
            </div>
          </aside>
        </div>
      )}
    </div>
  );
};
