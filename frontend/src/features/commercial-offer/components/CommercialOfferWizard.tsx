import { useCallback, useEffect, useRef, useState } from "react";
import { useNavigate } from "react-router";
import { useQueryClient } from "@tanstack/react-query";

import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import { archiveKeys } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { planApplyAiSessionSync } from "@/features/commercial-offer/lib/applyAiSession";
import {
  getBatches,
  getCurrentBatchReviewText,
  mergeEditedBatchIntoFullText,
} from "@/features/commercial-offer/lib/batchReview";
import { getDraftBatchCount } from "@/features/commercial-offer/lib/getDraftBatchCount";
import { getProductTypePreview } from "@/features/commercial-offer/lib/productTypePreview";
import type { ConcreteSpecPatch } from "@/features/commercial-offer/lib/concreteSpec";
import {
  getProductInputStep,
  getWizardStepOrder,
  isInputStepBlockedWithoutAppendCycle,
  mapLegacyWizardStep,
  shouldSkipClientStep,
} from "@/features/commercial-offer/lib/wizardStepOrder";
import {
  getProductTypeConfig,
  INGEST_REQUIRED_MESSAGE,
  PRODUCT_TYPE_CONFIG,
  RESOLVE_GATE_MESSAGES,
} from "@/features/commercial-offer/lib/productTypeConfig";

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
  carryWideSlots,
  chooseWideSlotAction,
  commitWideSlots,
  photoQueueWidePlateBlocksFinish,
  prepareTextReviewConfirm,
  routeWidePlateApply,
  TEXT_REVIEW_SLOT_PAGE_ID,
  uniqueWideSlotsByLine,
} from "@/features/commercial-offer/lib/widePlateSlots";
import {
  buildWidePlateApplyFlushText,
  findWidePlateDecision,
  flushThenResolveWidePlates,
} from "@/features/commercial-offer/lib/flushThenResolveWidePlates";
import {
  applyWideResolveToActivePage,
  syncWidePlateQueuePages,
} from "@/features/commercial-offer/lib/syncWidePlateQueuePages";
import {
  wideKeysInText,
  undecidedWideKeys,
} from "@/features/commercial-offer/lib/widePlatePageGate";

import { WizardProgress } from "@/features/commercial-offer/components/WizardProgress";
import { ProductTypePicker } from "@/features/commercial-offer/components/ProductTypePicker";
import { PlateInputStep } from "@/features/commercial-offer/components/steps/PlateInputStep";
import { SimpleProductInputStep } from "@/features/commercial-offer/components/steps/SimpleProductInputStep";
import { ClientConditionsStep } from "@/features/commercial-offer/components/steps/ClientConditionsStep";
import { CalculationResultStep } from "@/features/commercial-offer/components/steps/CalculationResultStep";
import { defaultInvalidWidthDecision } from "@/features/commercial-offer/components/InvalidWidthsInlineSection";

import type {
  CommercialDraftDetails,
  ProductType,
  SimpleKpProductType,
  WizardStepId,
} from "@/features/commercial-offer/types/commercialOffer";
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
    updateInputMutation,
    applyAiMutation,
    updateGradesMutation,
    updateConcreteSpecMutation,
    resolveWidePlatesMutation,
    rememberResolvedDraft,
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
    draftProductType,
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
  const wideSlotTextRef = useRef<Record<string, string>>({});
  const wideSlotsByPageRef = useRef(state.wideSlotsByPage);
  wideSlotsByPageRef.current = state.wideSlotsByPage;
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
  const productConfig = getProductTypeConfig(productType);
  const isSimpleProductFlow = productConfig.isSimpleKp;
  const skipClient = shouldSkipClientStep({
    clientName: state.clientName,
    counterpartyId: state.counterpartyId,
    appendBatches: state.lastDraft?.metadata.append_batches ?? currentDraft?.metadata.append_batches,
    resumeKpId: state.lastDraft?.metadata.resume_kp_id ?? currentDraft?.metadata.resume_kp_id ?? null,
  });
  const stepOrder = getWizardStepOrder(productType, { skipClient });
  const inputStep = getProductInputStep(productType);
  const stepIndex = (step: WizardStepId) => stepOrder.indexOf(step);

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
          productType: pageProductType,
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
        productType,
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
    productType,
    recognizedImagePreview?.file,
    updateInputMutation,
  ]);

  // Photo review shows the active page. Hydrate-draft can overwrite the store with the last batch.
  useEffect(() => {
    if (!multiPage.hasStarted || !multiPage.activePage) {
      return;
    }
    const pageReady =
      multiPage.activePage.status === "ready" || multiPage.activePage.status === "confirmed";
    if (!pageReady || state.batchReviewText === multiPage.activePage.batchReviewText) {
      return;
    }
    dispatch({ type: "set-batch-review-text", text: multiPage.activePage.batchReviewText });
  }, [
    multiPage.hasStarted,
    multiPage.activeId,
    multiPage.activePage?.status,
    multiPage.activePage?.batchReviewText,
    state.batchReviewText,
    state.lastDraft,
    dispatch,
  ]);

  const wideChoicePageText = multiPage.hasStarted
    ? multiPage.pages.map((page) => page.batchReviewText).join("\n")
    : "";

  useEffect(() => {
    if (!state.widePlateChoicesHeld || !multiPage.hasStarted) {
      return;
    }
    dispatch({ type: "prune-wide-plate-choices", pageText: wideChoicePageText });
  }, [dispatch, multiPage.hasStarted, state.widePlateChoicesHeld, wideChoicePageText]);

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
        `Введите текст списка ${productConfig.labels.nounGenitivePlural} или загрузите изображение.`,
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
          productType,
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
        `Сначала распознайте список ${productConfig.labels.nounGenitivePlural}, затем используйте помощника.`,
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
        productType,
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
      const editedActiveText = active.batchReviewText;
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
            productType,
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
    const plan = prepareTextReviewConfirm(
      state.batchReviewText,
      wideSlotsByPageRef.current[TEXT_REVIEW_SLOT_PAGE_ID] ?? [],
      lastBatch?.normalized_text ?? "",
    );
    wideSlotTextRef.current[TEXT_REVIEW_SLOT_PAGE_ID] = plan.text;
    wideSlotsByPageRef.current = {
      ...wideSlotsByPageRef.current,
      [TEXT_REVIEW_SLOT_PAGE_ID]: plan.slots,
    };
    dispatch({ type: "set-batch-review-text", text: plan.text });
    dispatch({
      type: "set-page-wide-slots",
      pageId: TEXT_REVIEW_SLOT_PAGE_ID,
      slots: plan.slots,
    });
    if (!plan.closeReview) {
      setWidePlateError(plan.error);
      return;
    }
    setWidePlateError(null);

    if (plan.uploadText) {
      try {
        const mergedText = mergeEditedBatchIntoFullText(batches, plan.uploadText);
        draft = await updateInputMutation.mutateAsync({
          draftId: draft.draft_id,
          productType,
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

  const handleFinishInput = async () => {
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
      } else if (draft.wizard_state.next_required_action === productConfig.ingestAction) {
        setStepError(INGEST_REQUIRED_MESSAGE);
      } else {
        const action = draft.wizard_state.next_required_action;
        const resolveMessage =
          action === "resolve_wide_plates" ||
          action === "resolve_invalid_widths" ||
          action === "resolve_unpriced_plates"
            ? RESOLVE_GATE_MESSAGES[action]
            : undefined;
        setStepError(
          resolveMessage ??
            `Нельзя перейти дальше — проверьте список ${productConfig.labels.nounGenitivePlural} и повторите.`,
        );
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
      await updateGradesMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        productType,
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
      // Only reachable for grade-supporting simple products — the wizard passes
      // onLineGradeChange solely when productConfig.supportsGrades.
      const preview = getProductTypePreview(productType);
      if (!preview) {
        return;
      }
      const rows = preview.buildRows(currentDraft);
      if (lineIndex < 0 || lineIndex >= rows.length || rows[lineIndex]?.sealed) {
        return;
      }
      const updated = rows.map((row, idx) => (idx === lineIndex ? { ...row, concrete_grade: grade } : row));
      const text = preview.buildLines(updated);
      await updateInputMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        productType,
        text,
        image: null,
        mode: updateMode,
      });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleConcreteSpecChange = async (lineId: string, patch: ConcreteSpecPatch) => {
    if (!currentDraft?.draft_id || !lineId) {
      return;
    }
    setStepError(null);
    try {
      await updateConcreteSpecMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        lineId,
        payload: patch,
      });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const openTextReview = state.pendingBatchReview && !multiPage.hasStarted;

  const wideSlotWriteTarget = (): { pageId: string; lines: string[] } | null => {
    const photoPageId = multiPage.activeId;
    if (photoPageId) {
      const pageText = multiPage.pages.find((page) => page.id === photoPageId)?.batchReviewText ?? "";
      return { pageId: photoPageId, lines: pageText.split(/\r?\n/) };
    }
    if (multiPage.hasStarted) {
      return null;
    }
    return {
      pageId: TEXT_REVIEW_SLOT_PAGE_ID,
      lines: state.batchReviewText.split(/\r?\n/),
    };
  };

  const applyLocalWideSlots = (slotIds: readonly string[]) => {
    const target = wideSlotWriteTarget();
    if (!target) {
      return;
    }
    const { pageId, lines } = target;
    const stored = uniqueWideSlotsByLine(
      wideSlotsByPageRef.current[pageId] ?? carryWideSlots(lines, lines, []),
    );
    setWidePlateError(null);
    const committed = commitWideSlots(lines, stored, slotIds);
    const nextText = committed.lines.join("\n");
    wideSlotTextRef.current[pageId] = nextText;
    wideSlotsByPageRef.current = {
      ...wideSlotsByPageRef.current,
      [pageId]: committed.slots,
    };
    if (multiPage.activeId) {
      multiPage.updatePageText(multiPage.activeId, nextText);
    }
    dispatch({ type: "set-batch-review-text", text: nextText });
    dispatch({
      type: "set-page-wide-slots",
      pageId,
      slots: committed.slots,
    });
    if (committed.error) {
      setWidePlateError(committed.error);
    }
  };

  const handleApplyWidePlates = async () => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setWidePlateError(null);
    const photoQueue = multiPage.hasStarted && multiPage.pages.length > 0;
    // Confirmed photo queue still resolves. Open photo review and open text review write locally.
    if (routeWidePlateApply(photoQueue && !multiPage.allConfirmed, { textReview: openTextReview }) === "local") {
      const pageId = multiPage.activeId ?? (openTextReview ? TEXT_REVIEW_SLOT_PAGE_ID : null);
      const openIds = (pageId ? (state.wideSlotsByPage[pageId] ?? []) : [])
        .filter((slot) => !slot.accepted)
        .map((slot) => slot.id);
      applyLocalWideSlots(openIds);
      return;
    }
    const pagesSnapshot = multiPage.pages;
    const activePageId = multiPage.activeId;
    const batchesBefore = currentDraft.metadata.plate_batches ?? [];
    const activePageText = pagesSnapshot.find((page) => page.id === activePageId)?.batchReviewText ?? "";
    const wideLines = photoQueue
      ? liveWidePlateLines(activePageText)
      : (currentDraft.metadata.wide_plate_lines ?? []);
    const decisionsById = state.widePlateActions;
    const flushText = buildWidePlateApplyFlushText({
      hasStarted: multiPage.hasStarted,
      pages: pagesSnapshot,
      activePageId,
      storeText: state.batchReviewText,
    });
    if (photoQueue) {
      dispatch({ type: "hold-wide-plate-choices" });
    }
    try {
      let resolvedDraft: CommercialDraftDetails | null = null;
      await flushThenResolveWidePlates({
        draftId: currentDraft.draft_id,
        flushText,
        persistedText: (currentDraft.metadata.input_text || currentDraft.metadata.normalized_text || "").trim(),
        liveLines: liveWidePlateLines(flushText),
        decisionsById,
        currentWideLines: wideLines,
        photoQueue,
        updateInput: (payload) => updateInputMutation.mutateAsync({ ...payload, productType }),
        resolveWidePlates: async (payload) => {
          resolvedDraft = await resolveWidePlatesMutation.mutateAsync({
            ...payload,
            deferPhotoQueueSync: photoQueue,
          });
        },
      });
      // Text input has no photo queue. Its mutation already synced the store.
      if (!photoQueue || resolvedDraft === null) {
        if (photoQueue) {
          dispatch({ type: "release-wide-plate-choices" });
        }
        return;
      }
      const layout = syncWidePlateQueuePages({
        pages: pagesSnapshot,
        activePageId,
        draft: resolvedDraft,
      });
      if (!layout.ok) {
        setWidePlateError(layout.message);
        // The page was not written. Keep the hold so the next hydrate does not drop replace/exclude.
        return;
      }
      const activePage = pagesSnapshot.find((page) => page.id === activePageId);
      const pageDecisions = liveWidePlateLines(activePage?.batchReviewText ?? "").map((line) => {
        const decision = findWidePlateDecision(line, decisionsById);
        return {
          sourceLine: decision?.sourceLine || line.line,
          action: decision?.action ?? "confirm",
          replacementText: decision?.replacementText ?? "",
        };
      });
      const writtenPages = applyWideResolveToActivePage({
        pages: pagesSnapshot,
        activePageId,
        batchesBefore,
        resolvedPages: layout.pages,
        decisions: pageDecisions,
      });
      for (const page of writtenPages) {
        if (page.id === activePageId) {
          multiPage.updatePageText(page.id, page.batchReviewText);
        }
      }
      if (activePageId) {
        dispatch({
          type: "settle-wide-plate-page",
          pageId: activePageId,
          keys: wideKeysInText(activePage?.batchReviewText ?? ""),
        });
      }
      const activeWritten =
        writtenPages.find((page) => page.id === activePageId)?.batchReviewText ?? layout.activePageText;
      dispatch({ type: "release-wide-plate-choices" });
      dispatch({ type: "sync-after-wide-plates", payload: resolvedDraft });
      dispatch({ type: "set-batch-review-text", text: activeWritten });
      rememberResolvedDraft(resolvedDraft);
    } catch (error) {
      if (photoQueue) {
        dispatch({ type: "release-wide-plate-choices" });
      }
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
          const fallback = defaultInvalidWidthDecision(item);
          return {
            lineId: item.id,
            sourceLine: item.line,
            action: decision?.action ?? fallback.action,
            widthMm: decision?.widthMm ?? fallback.widthMm,
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
    counterpartyId: number | null;
    counterpartyCode1c: string;
    counterpartyInn: string | null;
    counterpartyKpp: string | null;
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
        counterpartyId: payload.counterpartyId,
        counterpartyCode1c: payload.counterpartyCode1c,
        counterpartyInn: payload.counterpartyInn,
        counterpartyKpp: payload.counterpartyKpp,
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
        counterpartyId: payload.counterpartyId,
        clientInn: payload.counterpartyInn,
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
    longPileDelivery?: Record<string, { trip_cost: number }>;
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
        longPileDelivery: payload.longPileDelivery,
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

  const handleSetOneoffPrice = async (lineId: string, unitPrice: number | null) => {
    if (!state.draftId) {
      return;
    }
    setLineRowError(null);
    setStepError(null);
    try {
      const draft = await patchDraftLineMutation.mutateAsync({
        draftId: state.draftId,
        lineId,
        payload: { unit_price: unitPrice },
      });
      dispatch({ type: "hydrate-draft", payload: draft });
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

  const activeReviewPage = multiPage.activePage;
  const multiPendingReview = multiPage.hasStarted && !multiPage.allConfirmed;
  const activeReviewPageId = activeReviewPage?.id ?? null;
  const activeReviewPageText = activeReviewPage?.batchReviewText;
  useEffect(() => {
    if (!multiPendingReview || !activeReviewPageId || activeReviewPageText === undefined) {
      return;
    }
    if (wideSlotTextRef.current[activeReviewPageId] === activeReviewPageText) {
      return;
    }
    const previous = wideSlotTextRef.current[activeReviewPageId];
    const stored = uniqueWideSlotsByLine(wideSlotsByPageRef.current[activeReviewPageId] ?? []);
    const previousLines =
      previous === undefined ? activeReviewPageText.split(/\r?\n/) : previous.split(/\r?\n/);
    const nextSlots = carryWideSlots(previousLines, activeReviewPageText.split(/\r?\n/), stored);
    wideSlotTextRef.current[activeReviewPageId] = activeReviewPageText;
    dispatch({ type: "set-page-wide-slots", pageId: activeReviewPageId, slots: nextSlots });
  }, [activeReviewPageId, activeReviewPageText, dispatch, multiPendingReview]);
  useEffect(() => {
    if (!openTextReview) {
      return;
    }
    const text = state.batchReviewText;
    if (wideSlotTextRef.current[TEXT_REVIEW_SLOT_PAGE_ID] === text) {
      return;
    }
    const previous = wideSlotTextRef.current[TEXT_REVIEW_SLOT_PAGE_ID];
    const stored = uniqueWideSlotsByLine(wideSlotsByPageRef.current[TEXT_REVIEW_SLOT_PAGE_ID] ?? []);
    const previousLines = previous === undefined ? text.split(/\r?\n/) : previous.split(/\r?\n/);
    const nextSlots = carryWideSlots(previousLines, text.split(/\r?\n/), stored);
    wideSlotTextRef.current[TEXT_REVIEW_SLOT_PAGE_ID] = text;
    wideSlotsByPageRef.current = {
      ...wideSlotsByPageRef.current,
      [TEXT_REVIEW_SLOT_PAGE_ID]: nextSlots,
    };
    dispatch({ type: "set-page-wide-slots", pageId: TEXT_REVIEW_SLOT_PAGE_ID, slots: nextSlots });
  }, [dispatch, openTextReview, state.batchReviewText]);
  const pendingBatchReview = multiPendingReview || state.pendingBatchReview;
  const reviewBatchText =
    multiPage.hasStarted && activeReviewPage ? activeReviewPage.batchReviewText : state.batchReviewText;
  const reopenKnownWideLines =
    multiPage.hasStarted &&
    Boolean(activeReviewPage) &&
    undecidedWideKeys(
      reviewBatchText,
      state.widePlateSettledKeysByPage[activeReviewPage?.id ?? ""] ?? [],
    ).length > 0;
  const hasUnresolvedWidePlates = multiPage.hasStarted
    ? photoQueueWidePlateBlocksFinish(
        multiPage.pages.map((page) => ({
          text: page.id === multiPage.activeId ? reviewBatchText : page.batchReviewText,
          slots: state.wideSlotsByPage[page.id] ?? [],
        })),
      )
    : photoQueueWidePlateBlocksFinish([
        {
          text: state.batchReviewText,
          slots: state.wideSlotsByPage[TEXT_REVIEW_SLOT_PAGE_ID] ?? [],
        },
      ]);
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
    if (step === "client" && !isSimpleProductFlow && hasUnresolvedWidePlates) {
      return false;
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
          `Сначала распознайте и обработайте список ${productConfig.labels.nounGenitivePlural}.`,
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
  const reviewImageUrl =
    activeReviewPage?.recognizedImageUrl ??
    activeReviewPage?.previewUrl ??
    recognizedImagePreview?.url ??
    null;
  const reviewImageName = activeReviewPage?.name ?? recognizedImagePreview?.name ?? null;
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
    state.currentStep === inputStep && productConfig.isSimpleKp ? (
      <SimpleProductInputStep
        productType={productType as SimpleKpProductType}
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
        isRecognizing={isRecognizingMulti || createDraftMutation.isPending || updateInputMutation.isPending}
        isAiProcessing={applyAiMutation.isPending}
        isUpdatingGrades={updateGradesMutation.isPending}
        isConfirmingBatch={updateInputMutation.isPending}
        isProceeding={calculateMutation.isPending}
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
        onFinishInput={() => void handleFinishInput()}
        onApplyGradeToAll={
          productConfig.supportsGrades ? (grade) => void handleApplyGradeToAll(grade) : undefined
        }
        onLineGradeChange={
          productConfig.supportsGrades
            ? (lineIndex, grade) => void handleLineGradeChange(lineIndex, grade)
            : undefined
        }
        onConcreteSpecChange={
          productConfig.supportsGrades ? (lineId, patch) => void handleConcreteSpecChange(lineId, patch) : undefined
        }
        onReset={handleCreateNewOffer}
        lineRowHandlers={lineRowHandlers}
        onSetOneoffPrice={(lineId, unitPrice) => void handleSetOneoffPrice(lineId, unitPrice)}
        oneoffError={lineRowError}
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
        isRecognizing={isRecognizingMulti || createDraftMutation.isPending || updateInputMutation.isPending}
        isAiProcessing={applyAiMutation.isPending}
        isResolvingWidePlates={resolveWidePlatesMutation.isPending}
        isResolvingUnpricedPlates={resolveUnpricedPlatesMutation.isPending}
        isResolvingInvalidWidths={resolveInvalidWidthsMutation.isPending}
        isConfirmingBatch={updateInputMutation.isPending}
        isProceeding={calculateMutation.isPending}
        widePlateDecisions={state.widePlateActions}
        wideSlots={
          multiPendingReview && activeReviewPage
            ? (state.wideSlotsByPage[activeReviewPage.id] ?? [])
            : !multiPage.hasStarted
              ? (state.wideSlotsByPage[TEXT_REVIEW_SLOT_PAGE_ID] ?? [])
              : undefined
        }
        wideSlotsByPage={state.wideSlotsByPage}
        collapseKey={
          multiPendingReview && activeReviewPage
            ? activeReviewPage.id
            : !multiPage.hasStarted
              ? TEXT_REVIEW_SLOT_PAGE_ID
              : undefined
        }
        onAcceptWideSlot={(slotId) => applyLocalWideSlots([slotId])}
        onApplyWideSlots={(slotIds) => applyLocalWideSlots(slotIds)}
        onWideSlotAction={(slotId, action) => {
          const target = wideSlotWriteTarget();
          if (!target) {
            return;
          }
          const { pageId, lines } = target;
          const stored = uniqueWideSlotsByLine(wideSlotsByPageRef.current[pageId] ?? []);
          const slots = uniqueWideSlotsByLine(
            stored.map((slot) => {
              if (slot.id !== slotId) {
                return slot;
              }
              const line = lines[slot.lineIndex] ?? "";
              const qty = liveWidePlateLines(line)[0]?.qty ?? 1;
              return chooseWideSlotAction(slot, action, line, qty);
            }),
          );
          wideSlotsByPageRef.current = { ...wideSlotsByPageRef.current, [pageId]: slots };
          dispatch({
            type: "set-page-wide-slots",
            pageId,
            slots,
          });
        }}
        onWideSlotReplacement={(slotId, replacementText) => {
          const target = wideSlotWriteTarget();
          if (!target) {
            return;
          }
          const { pageId } = target;
          const stored = uniqueWideSlotsByLine(wideSlotsByPageRef.current[pageId] ?? []);
          const slots = uniqueWideSlotsByLine(
            stored.map((slot) => (slot.id === slotId ? { ...slot, replacementText } : slot)),
          );
          wideSlotsByPageRef.current = { ...wideSlotsByPageRef.current, [pageId]: slots };
          dispatch({
            type: "set-page-wide-slots",
            pageId,
            slots,
          });
        }}
        reopenKnownWideLines={reopenKnownWideLines}
        unpricedPlateDecisions={state.unpricedPlateActions}
        invalidWidthDecisions={state.invalidWidthActions}
        aiInstruction={aiInstruction}
        onAiInstructionChange={setAiInstruction}
        onApplyAi={() => void handleApplyAi()}
        onTextChange={handleSourceTextChange}
        onBatchReviewTextChange={(value) => {
          dispatch({ type: "set-batch-review-text", text: value });
          const pageId = multiPage.activeId;
          if (!pageId) {
            if (openTextReview) {
              const previous = wideSlotTextRef.current[TEXT_REVIEW_SLOT_PAGE_ID] ?? state.batchReviewText;
              const nextSlots = carryWideSlots(
                previous.split(/\r?\n/),
                value.split(/\r?\n/),
                uniqueWideSlotsByLine(wideSlotsByPageRef.current[TEXT_REVIEW_SLOT_PAGE_ID] ?? []),
              );
              wideSlotTextRef.current[TEXT_REVIEW_SLOT_PAGE_ID] = value;
              wideSlotsByPageRef.current = {
                ...wideSlotsByPageRef.current,
                [TEXT_REVIEW_SLOT_PAGE_ID]: nextSlots,
              };
              dispatch({
                type: "set-page-wide-slots",
                pageId: TEXT_REVIEW_SLOT_PAGE_ID,
                slots: nextSlots,
              });
            }
            return;
          }
          if (multiPage.hasStarted && !multiPage.allConfirmed) {
            const previous = multiPage.pages.find((page) => page.id === pageId)?.batchReviewText ?? "";
            const nextSlots = carryWideSlots(
              previous.split(/\r?\n/),
              value.split(/\r?\n/),
              uniqueWideSlotsByLine(wideSlotsByPageRef.current[pageId] ?? []),
            );
            wideSlotTextRef.current[pageId] = value;
            wideSlotsByPageRef.current = { ...wideSlotsByPageRef.current, [pageId]: nextSlots };
            dispatch({ type: "set-page-wide-slots", pageId, slots: nextSlots });
          }
          multiPage.updatePageText(pageId, value);
        }}
        
        onRecognize={handleRecognize}
        onRerecognize={() => void handleRerecognize()}
        isRerecognizing={isRerecognizing}
        onConfirmBatch={() => void handleConfirmBatch()}
        onFinishPlates={() => void handleFinishInput()}
        onWidePlateDecisionChange={(lineId, action, replacementText, sourceLine) =>
          dispatch({ type: "set-wide-action", lineId, action, replacementText, sourceLine })
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
        onConcreteSpecChange={(lineId, patch) => void handleConcreteSpecChange(lineId, patch)}
        lineRowHandlers={lineRowHandlers}
      />
    ) : state.currentStep === "client" ? (
      <ClientConditionsStep
        managers={managers}
        selectedManagerId={state.managerId}
        defaultValues={{
          clientName: state.clientName,
          counterpartyId: state.counterpartyId,
          counterpartyCode1c: state.counterpartyCode1c,
          counterpartyInn: state.counterpartyInn,
          counterpartyKpp: state.counterpartyKpp,
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
        draftProductType={draftProductType}
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
      if (Object.prototype.hasOwnProperty.call(PRODUCT_TYPE_CONFIG, raw)) {
        seen.add(raw as ProductType);
        ordered.push(raw as ProductType);
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
