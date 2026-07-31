import { useEffect, useState } from "react";

import { mergeEditedBatchIntoFullText } from "@/features/commercial-offer/lib/batchReview";
import {
  buildPileLinesFromOrderData,
  buildPilePreviewRows,
} from "@/features/commercial-offer/lib/buildPilePreviewRows";
import {
  getProductInputStep,
  getWizardStepOrder,
  mapLegacyWizardStep,
  wizardStepIndex,
} from "@/features/commercial-offer/lib/wizardStepOrder";

import { useCommercialOfferWizard } from "@/features/commercial-offer/hooks/useCommercialOfferWizard";
import { useRecognizedImagePreview } from "@/features/commercial-offer/hooks/useRecognizedImagePreview";

import { WizardProgress } from "@/features/commercial-offer/components/WizardProgress";
import { PlateInputStep } from "@/features/commercial-offer/components/steps/PlateInputStep";
import { PileInputStep } from "@/features/commercial-offer/components/steps/PileInputStep";
import { ClientConditionsStep } from "@/features/commercial-offer/components/steps/ClientConditionsStep";
import { CalculationResultStep } from "@/features/commercial-offer/components/steps/CalculationResultStep";

import type { ProductType, WizardStepId } from "@/features/commercial-offer/types/commercialOffer";

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
    applyAiPlatesMutation,
    applyAiPilesMutation,
    updatePileGradesMutation,
    resolveWidePlatesMutation,
    updateMetaMutation,
    calculateMutation,
    generateFilesMutation,
    generateSchemaMutation,
    saveDraftMutation,
    isPileDraft,
  } = useCommercialOfferWizard();

  const [selectedImage, setSelectedImage] = useState<File | null>(null);
  const [aiInstruction, setAiInstruction] = useState("");
  const [stepError, setStepError] = useState<string | null>(null);
  const [widePlateError, setWidePlateError] = useState<string | null>(null);
  const {
    preview: recognizedImagePreview,
    setPreviewFromFile,
    clearPreview: clearRecognizedImagePreview,
  } = useRecognizedImagePreview();

  const productType = state.productType;
  const isPileFlow = productType === "piles";
  const stepOrder = getWizardStepOrder(productType);
  const inputStep = getProductInputStep(productType);
  const updateInputMutation = isPileFlow ? updatePilesMutation : updatePlatesMutation;
  const applyAiMutation = isPileFlow ? applyAiPilesMutation : applyAiPlatesMutation;

  const managers = managersQuery.data?.items ?? [];

  useEffect(() => {
    if (productTypeProp !== state.productType) {
      dispatch({ type: "set-product-type", productType: productTypeProp });
    }
  }, [productTypeProp, state.productType, dispatch]);

  useEffect(() => {
    const step = mapLegacyWizardStep(state.currentStep);
    if (step !== state.currentStep) {
      dispatch({ type: "set-step", step });
      return;
    }
    if (!stepOrder.includes(step)) {
      dispatch({ type: "set-step", step: inputStep });
      return;
    }
    if (step === "result" && !state.draftId) {
      dispatch({ type: "set-step", step: inputStep });
    }
  }, [state.currentStep, state.draftId, dispatch, stepOrder, inputStep]);

  const resetSource = () => {
    setSelectedImage(null);
    dispatch({ type: "set-source", text: "", imageName: null });
  };

  const handleSourceTextChange = (value: string) => {
    if (selectedImage) {
      setSelectedImage(null);
    }
    dispatch({ type: "set-source", text: value, imageName: null });
  };

  const handleImageSelect = (file: File | null) => {
    setSelectedImage(file);
    dispatch({ type: "set-source", text: file ? "" : state.sourceText, imageName: file?.name ?? null });
  };

  const handleRecognize = async (mode: "append" | "replace") => {
    setStepError(null);
    if (!state.sourceText.trim() && !selectedImage) {
      setStepError(
        isPileFlow
          ? "Введите текст списка свай или загрузите изображение."
          : "Введите текст списка плит или загрузите изображение.",
      );
      return;
    }

    const sourceText = selectedImage ? "" : state.sourceText;
    const imageForRecognition = selectedImage;

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
          productType: isPileFlow ? "piles" : "plates",
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
        isPileFlow
          ? "Сначала распознайте список свай, затем используйте помощника."
          : "Сначала распознайте список плит, затем используйте помощника.",
      );
      return;
    }

    try {
      if (selectedImage) {
        setPreviewFromFile(selectedImage);
      }
      await applyAiMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        instruction,
        image: selectedImage,
      });
      resetSource();
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

    let draft = currentDraft;
    const batches = isPileFlow
      ? (draft.metadata.pile_batches ?? [])
      : (draft.metadata.plate_batches ?? []);
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
          batchCount: isPileFlow
            ? (draft.metadata.pile_batches?.length ?? 0)
            : (draft.metadata.plate_batches?.length ?? 0),
        });
      } catch (error) {
        setStepError(getErrorMessage(error));
        return;
      }
    } else {
      dispatch({ type: "confirm-batch-review", batchCount: batches.length });
    }

    clearRecognizedImagePreview();
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
      } else {
        setStepError("Нельзя перейти дальше — проверьте список плит и повторите.");
      }
      return;
    }
    dispatch({ type: "set-step", step: next });
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
    dispatch({ type: "set-step", step: next });
  };

  const handleApplyGradeToAll = async (grade: string) => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    try {
      await updatePileGradesMutation.mutateAsync({
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
    try {
      const rows = buildPilePreviewRows(currentDraft);
      if (lineIndex < 0 || lineIndex >= rows.length) {
        return;
      }
      const updated = rows.map((row, idx) => (idx === lineIndex ? { ...row, concrete_grade: grade } : row));
      const text = buildPileLinesFromOrderData(updated);
      await updatePilesMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        text,
        image: null,
        mode: "replace",
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
      await resolveWidePlatesMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        decisions: currentDraft.metadata.wide_plate_lines.map((item) => ({
          lineId: item.id,
          sourceLine: item.line,
          action: state.widePlateActions[item.id]?.action ?? "confirm",
          replacementText: state.widePlateActions[item.id]?.replacementText ?? "",
        })),
      });
    } catch (error) {
      setWidePlateError(getErrorMessage(error));
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
        fileTypes: isPileDraft ? ["pdf", "xlsx"] : undefined,
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

  const handleCreateNewOffer = () => {
    setStepError(null);
    setWidePlateError(null);
    setSelectedImage(null);
    clearRecognizedImagePreview();
    dispatch({ type: "reset" });
  };

  const hasUnresolvedWidePlates =
    Boolean(currentDraft?.metadata.wide_plate_lines?.length) && !currentDraft?.metadata.wide_plates_resolved;

  const canNavigateToStep = (step: WizardStepId): boolean => {
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
    if (wizardStepIndex(step, productType) < wizardStepIndex(ws.current_step as WizardStepId, productType)) {
      return true;
    }
    return ws.can_proceed_to.includes(step);
  };

  const handleSidebarStepClick = (step: WizardStepId) => {
    if (!canNavigateToStep(step)) {
      if (!currentDraft && step !== inputStep) {
        setStepError(
          isPileFlow
            ? "Сначала распознайте и обработайте список свай."
            : "Сначала распознайте и обработайте список плит.",
        );
        return;
      }
      if (step === "client" && !isPileFlow && (hasUnresolvedWidePlates || state.pendingBatchReview)) {
        setStepError(
          state.pendingBatchReview
            ? "Сначала подтвердите список текущего источника — «Список верен»."
            : "Сначала примите решение по позициям шире стандартной.",
        );
        return;
      }
      if (step === "client" && isPileFlow && state.pendingBatchReview) {
        setStepError("Сначала подтвердите список текущего источника — «Список верен».");
        return;
      }
      return;
    }
    setStepError(null);
    dispatch({ type: "set-step", step });
  };

  const currentStepContent =
    state.currentStep === "piles" ? (
      <PileInputStep
        draft={currentDraft}
        pendingBatchReview={state.pendingBatchReview}
        sourceText={state.sourceText}
        batchReviewText={state.batchReviewText}
        normalizedText={state.normalizedText}
        selectedImageName={state.selectedImageName}
        recognizedImageUrl={recognizedImagePreview?.url ?? null}
        recognizedImageName={recognizedImagePreview?.name ?? null}
        errorMessage={stepError}
        isRecognizing={createDraftMutation.isPending || updatePilesMutation.isPending}
        isAiProcessing={applyAiPilesMutation.isPending}
        isUpdatingGrades={updatePileGradesMutation.isPending}
        isConfirmingBatch={updatePilesMutation.isPending}
        isProceeding={false}
        aiInstruction={aiInstruction}
        onAiInstructionChange={setAiInstruction}
        onApplyAi={() => void handleApplyAi()}
        onTextChange={handleSourceTextChange}
        onBatchReviewTextChange={(value) => dispatch({ type: "set-batch-review-text", text: value })}
        onFileChange={handleImageSelect}
        onImagePaste={handleImageSelect}
        onRecognize={handleRecognize}
        onConfirmBatch={() => void handleConfirmBatch()}
        onFinishPiles={() => void handleFinishPiles()}
        onApplyGradeToAll={(grade) => void handleApplyGradeToAll(grade)}
        onLineGradeChange={(lineIndex, grade) => void handleLineGradeChange(lineIndex, grade)}
        onReset={handleCreateNewOffer}
      />
    ) : state.currentStep === "plates" ? (
      <PlateInputStep
        draft={currentDraft}
        pendingBatchReview={state.pendingBatchReview}
        sourceText={state.sourceText}
        batchReviewText={state.batchReviewText}
        normalizedText={state.normalizedText}
        selectedImageName={state.selectedImageName}
        recognizedImageUrl={recognizedImagePreview?.url ?? null}
        recognizedImageName={recognizedImagePreview?.name ?? null}
        errorMessage={stepError}
        widePlateErrorMessage={widePlateError}
        isRecognizing={createDraftMutation.isPending || updatePlatesMutation.isPending}
        isAiProcessing={applyAiPlatesMutation.isPending}
        isResolvingWidePlates={resolveWidePlatesMutation.isPending}
        isConfirmingBatch={updatePlatesMutation.isPending}
        isProceeding={false}
        widePlateDecisions={state.widePlateActions}
        aiInstruction={aiInstruction}
        onAiInstructionChange={setAiInstruction}
        onApplyAi={() => void handleApplyAi()}
        onTextChange={handleSourceTextChange}
        onBatchReviewTextChange={(value) => dispatch({ type: "set-batch-review-text", text: value })}
        onFileChange={handleImageSelect}
        onImagePaste={handleImageSelect}
        onRecognize={handleRecognize}
        onConfirmBatch={() => void handleConfirmBatch()}
        onFinishPlates={() => void handleFinishPlates()}
        onWidePlateDecisionChange={(lineId, action, replacementText) =>
          dispatch({ type: "set-wide-action", lineId, action, replacementText })
        }
        onApplyWidePlates={() => void handleApplyWidePlates()}
        onReset={handleCreateNewOffer}
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
        isGeneratingFiles={generateFilesMutation.isPending}
        isGeneratingSchema={generateSchemaMutation.isPending}
        isSaving={saveDraftMutation.isPending}
        lastSaveResult={state.lastSaveResult}
        executionTermsInput={state.executionTermsInput}
        onBack={() => dispatch({ type: "set-step", step: "client" })}
        onCreateNew={handleCreateNewOffer}
        onGenerateFiles={handleGenerateFiles}
        onGenerateSchema={handleGenerateSchema}
        onExecutionTermsChange={(value) => dispatch({ type: "set-execution-terms", value })}
        onSave={handleSave}
        isUpdatingDiscount={updateMetaMutation.isPending || calculateMutation.isPending}
        onDiscountSubmit={handleDiscountSubmit}
        onLogisticsCostSubmit={handleLogisticsCostSubmit}
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

  return (
    <div style={{ display: "grid", gap: "1.25rem" }}>
      {managersQuery.error && <Alert tone="error">{getErrorMessage(managersQuery.error)}</Alert>}

      <div className="wizard-shell">
        <div className="wizard-main">{resolvedStepContent}</div>

        <aside className="wizard-sidebar">
          <div className="wizard-sidebar__inner">
            <WizardProgress
              productType={productType}
              currentStep={state.currentStep}
              onStepClick={handleSidebarStepClick}
              canNavigateToStep={canNavigateToStep}
            />
          </div>
        </aside>
      </div>
    </div>
  );
};
