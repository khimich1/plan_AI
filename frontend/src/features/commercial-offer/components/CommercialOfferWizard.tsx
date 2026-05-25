import { useEffect, useState } from "react";
import { useCommercialOfferWizard } from "@/features/commercial-offer/hooks/useCommercialOfferWizard";
import { useRecognizedImagePreview } from "@/features/commercial-offer/hooks/useRecognizedImagePreview";
import { WIZARD_STEP_ORDER, wizardStepIndex } from "@/features/commercial-offer/lib/wizardStepOrder";
import { WizardProgress } from "@/features/commercial-offer/components/WizardProgress";
import { PlateInputStep } from "@/features/commercial-offer/components/steps/PlateInputStep";
import { WidePlateReviewStep } from "@/features/commercial-offer/components/steps/WidePlateReviewStep";
import { ManagerStep } from "@/features/commercial-offer/components/steps/ManagerStep";
import { ClientConditionsStep } from "@/features/commercial-offer/components/steps/ClientConditionsStep";
import { CalculationResultStep } from "@/features/commercial-offer/components/steps/CalculationResultStep";
import type { WizardStepId } from "@/features/commercial-offer/types/commercialOffer";
import { getErrorMessage } from "@/shared/lib/apiError";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Spinner } from "@/shared/ui/Spinner";

export const CommercialOfferWizard = () => {
  const {
    state,
    dispatch,
    managersQuery,
    draftQuery,
    currentDraft,
    createDraftMutation,
    updatePlatesMutation,
    resolveWidePlatesMutation,
    updateMetaMutation,
    calculateMutation,
    generateFilesMutation,
    generateSchemaMutation,
    saveDraftMutation,
  } = useCommercialOfferWizard();
  const [selectedImage, setSelectedImage] = useState<File | null>(null);
  const [stepError, setStepError] = useState<string | null>(null);
  const {
    preview: recognizedImagePreview,
    setPreviewFromFile,
    clearPreview: clearRecognizedImagePreview,
  } = useRecognizedImagePreview();

  const managers = managersQuery.data?.items ?? [];

  useEffect(() => {
    const step = state.currentStep;
    if (!WIZARD_STEP_ORDER.includes(step)) {
      dispatch({ type: "set-step", step: "plates" });
      return;
    }
    if ((step === "wide-plates" || step === "result") && !state.draftId) {
      dispatch({ type: "set-step", step: "plates" });
    }
  }, [state.currentStep, state.draftId, dispatch]);

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
      setStepError("Введите текст списка плит или загрузите изображение.");
      return;
    }

    const sourceText = selectedImage ? "" : state.sourceText;
    const imageForRecognition = selectedImage;

    try {
      if (imageForRecognition) {
        setPreviewFromFile(imageForRecognition);
      }
      if (currentDraft?.draft_id) {
        await updatePlatesMutation.mutateAsync({
          draftId: currentDraft.draft_id,
          text: sourceText,
          image: imageForRecognition,
          mode,
        });
      } else {
        await createDraftMutation.mutateAsync({
          text: sourceText,
          image: imageForRecognition,
        });
      }
      resetSource();
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleProcess = async () => {
    setStepError(null);
    if (!currentDraft?.wizard_state || !currentDraft.draft_id) {
      setStepError("Нет состояния мастера с сервера. Обновите черновик.");
      return;
    }

    let draft = currentDraft;
    const editedNormalizedText = state.normalizedText.trim();
    if (editedNormalizedText && editedNormalizedText !== draft.metadata.normalized_text) {
      try {
        draft = await updatePlatesMutation.mutateAsync({
          draftId: draft.draft_id,
          text: editedNormalizedText,
          image: null,
          mode: "replace",
        });
      } catch (error) {
        setStepError(getErrorMessage(error));
        return;
      }
    }

    const next = draft.wizard_state.can_proceed_to[0];
    if (!next) {
      const serverMsgs = (draft.wizard_state.validation_errors ?? []).filter(Boolean);
      if (serverMsgs.length > 0) {
        setStepError(serverMsgs.join(" "));
      } else if (draft.wizard_state.next_required_action === "ingest_plates") {
        setStepError("Сначала распознайте и получите хотя бы одну позицию в заказе.");
      } else {
        setStepError("Сервер не разрешает переход на следующий шаг. Проверьте данные и повторите запрос.");
      }
      return;
    }
    dispatch({ type: "set-step", step: next });
  };
  const handleWideSubmit = async (): Promise<boolean> => {
    if (!currentDraft?.draft_id) {
      return false;
    }
    setStepError(null);
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
      return true;
    } catch (error) {
      setStepError(getErrorMessage(error));
      return false;
    }
  };

  const handleManagerSubmit = async () => {
    if (!currentDraft?.draft_id || !state.managerId) {
      setStepError("Выберите менеджера.");
      return;
    }
    setStepError(null);
    try {
      await updateMetaMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        managerId: state.managerId,
      });
      dispatch({ type: "set-step", step: "client" });
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleClientSubmit = async (payload: {
    clientName: string;
    conditionsMode: "standard" | "custom";
    deliveryConditions: string;
    paymentConditions: string;
  }) => {
    if (!currentDraft?.draft_id) {
      return;
    }
    setStepError(null);
    dispatch({ type: "set-client-form", payload });
    try {
      await updateMetaMutation.mutateAsync({
        draftId: currentDraft.draft_id,
        managerId: state.managerId,
        clientName: payload.clientName,
        conditionsMode: payload.conditionsMode,
        deliveryConditions: payload.deliveryConditions,
        paymentConditions: payload.paymentConditions,
      });
      const calculated = await calculateMutation.mutateAsync(currentDraft.draft_id);
      const ws = calculated.wizard_state;
      if (ws.current_step !== "result") {
        const msgs = (ws.validation_errors ?? []).filter(Boolean);
        setStepError(
          msgs.length > 0
            ? msgs.join(" ")
            : "Сервер не перевёл черновик на шаг результата. Проверьте данные клиента и условия.",
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
      await generateFilesMutation.mutateAsync(currentDraft.draft_id);
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
    setSelectedImage(null);
    clearRecognizedImagePreview();
    dispatch({ type: "reset" });
  };

  const canNavigateToStep = (step: WizardStepId): boolean => {
    if (step === "plates") {
      return true;
    }
    if (!currentDraft?.wizard_state) {
      return false;
    }
    const ws = currentDraft.wizard_state;
    if (step === state.currentStep) {
      return true;
    }
    if (wizardStepIndex(step) < wizardStepIndex(ws.current_step)) {
      return true;
    }
    return ws.can_proceed_to.includes(step);
  };
  const handleSidebarStepClick = (step: WizardStepId) => {
    if (!canNavigateToStep(step)) {
      if (!currentDraft && step !== "plates") {
        setStepError("Сначала распознайте и обработайте список плит.");
        return;
      }
      if (step === "wide-plates") {
        setStepError("Нет проблемных плит для отдельной проверки.");
      }
      return;
    }
    setStepError(null);
    dispatch({ type: "set-step", step });
  };

  const currentStepContent =
    state.currentStep === "plates" ? (
      <PlateInputStep
        draft={currentDraft}
        sourceText={state.sourceText}
        normalizedText={state.normalizedText}
        selectedImageName={state.selectedImageName}
        recognizedImageUrl={recognizedImagePreview?.url ?? null}
        recognizedImageName={recognizedImagePreview?.name ?? null}
        errorMessage={stepError}
        isRecognizing={createDraftMutation.isPending || updatePlatesMutation.isPending}
        onTextChange={handleSourceTextChange}
        onNormalizedTextChange={(value) => dispatch({ type: "set-normalized-text", text: value })}
        onFileChange={handleImageSelect}
        onImagePaste={handleImageSelect}
        onRecognize={handleRecognize}
        onProcess={() => void handleProcess()}
        onReset={handleCreateNewOffer}
      />
    ) : state.currentStep === "wide-plates" && currentDraft ? (
      <WidePlateReviewStep
        draft={currentDraft}
        decisions={state.widePlateActions}
        errorMessage={stepError}
        isPending={resolveWidePlatesMutation.isPending}
        onDecisionChange={(lineId, action, replacementText) =>
          dispatch({ type: "set-wide-action", lineId, action, replacementText })
        }
        onBack={() => dispatch({ type: "set-step", step: "plates" })}
        onSubmit={async () => {
          const success = await handleWideSubmit();
          if (success) {
            dispatch({ type: "set-step", step: "manager" });
          }
        }}
      />
    ) : state.currentStep === "manager" ? (
      <ManagerStep
        managers={managers}
        selectedManagerId={state.managerId}
        errorMessage={stepError}
        isPending={updateMetaMutation.isPending}
        onSelect={(managerId) => dispatch({ type: "set-manager", managerId })}
        onBack={() =>
          dispatch({
            type: "set-step",
            step: currentDraft?.metadata?.wide_plate_lines?.length ? "wide-plates" : "plates",
          })
        }
        onNext={handleManagerSubmit}
      />
    ) : state.currentStep === "client" ? (
      <ClientConditionsStep
        defaultValues={{
          clientName: state.clientName,
          conditionsMode: state.conditionsMode,
          deliveryConditions: state.deliveryConditions,
          paymentConditions: state.paymentConditions,
        }}
        errorMessage={stepError}
        isPending={updateMetaMutation.isPending || calculateMutation.isPending}
        onBack={() => dispatch({ type: "set-step", step: "manager" })}
        onSubmit={handleClientSubmit}
      />
    ) : state.currentStep === "result" && currentDraft ? (
      <CalculationResultStep
        draft={currentDraft}
        errorMessage={stepError}
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
    (state.currentStep === "wide-plates" || state.currentStep === "result"
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
