import { useState } from "react";
import { useCommercialOfferWizard } from "@/features/commercial-offer/hooks/useCommercialOfferWizard";
import { WizardProgress } from "@/features/commercial-offer/components/WizardProgress";
import { PlateInputStep } from "@/features/commercial-offer/components/steps/PlateInputStep";
import { WidePlateReviewStep } from "@/features/commercial-offer/components/steps/WidePlateReviewStep";
import { ManagerStep } from "@/features/commercial-offer/components/steps/ManagerStep";
import { ClientConditionsStep } from "@/features/commercial-offer/components/steps/ClientConditionsStep";
import { CalculationResultStep } from "@/features/commercial-offer/components/steps/CalculationResultStep";
import type { CommercialDraftDetails, WidePlateAction, WizardStepId } from "@/features/commercial-offer/types/commercialOffer";
import { getErrorMessage } from "@/shared/lib/apiError";
import { Alert } from "@/shared/ui/Alert";

const getNextStepFromDraft = (draft: CommercialDraftDetails): WizardStepId => {
  if (draft.metadata.wide_plate_lines.length > 0 && !draft.metadata.wide_plates_resolved) {
    return "wide-plates";
  }
  if (!draft.metadata.manager_id) {
    return "manager";
  }
  if (!draft.metadata.client_name) {
    return "client";
  }
  return "result";
};

export const CommercialOfferWizard = () => {
  const {
    state,
    dispatch,
    managersQuery,
    currentDraft,
    createDraftMutation,
    updatePlatesMutation,
    resolveWidePlatesMutation,
    updateMetaMutation,
    generateFilesMutation,
    saveDraftMutation,
  } = useCommercialOfferWizard();
  const [selectedImage, setSelectedImage] = useState<File | null>(null);
  const [stepError, setStepError] = useState<string | null>(null);

  const managers = managersQuery.data?.items ?? [];

  const resetSource = () => {
    setSelectedImage(null);
    dispatch({ type: "set-source", text: "", imageName: null });
  };

  const handleRecognize = async (mode: "append" | "replace") => {
    setStepError(null);
    if (!state.sourceText.trim() && !selectedImage) {
      setStepError("Введите текст списка плит или загрузите изображение.");
      return;
    }

    try {
      if (currentDraft?.draft_id) {
        await updatePlatesMutation.mutateAsync({
          draftId: currentDraft.draft_id,
          text: state.sourceText,
          image: selectedImage,
          mode,
        });
      } else {
        await createDraftMutation.mutateAsync({
          text: state.sourceText,
          image: selectedImage,
        });
      }
      resetSource();
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleProcess = () => {
    setStepError(null);
    if (!currentDraft) {
      setStepError("Сначала обработайте список плит.");
      return;
    }
    dispatch({ type: "set-step", step: getNextStepFromDraft(currentDraft) });
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
      dispatch({ type: "set-step", step: "result" });
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
    } catch (error) {
      setStepError(getErrorMessage(error));
    }
  };

  const handleCreateNewOffer = () => {
    setStepError(null);
    setSelectedImage(null);
    dispatch({ type: "reset" });
  };

  const currentStepContent =
    state.currentStep === "plates" ? (
      <PlateInputStep
        draft={currentDraft}
        sourceText={state.sourceText}
        selectedImageName={state.selectedImageName}
        errorMessage={stepError}
        isRecognizing={createDraftMutation.isPending || updatePlatesMutation.isPending}
        onTextChange={(value) => dispatch({ type: "set-source", text: value, imageName: selectedImage?.name ?? null })}
        onFileChange={(file) => {
          setSelectedImage(file);
          dispatch({ type: "set-source", text: state.sourceText, imageName: file?.name ?? null });
        }}
        onRecognize={handleRecognize}
        onProcess={handleProcess}
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
            step: currentDraft?.metadata.wide_plate_lines.length ? "wide-plates" : "plates",
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
        isPending={updateMetaMutation.isPending}
        onBack={() => dispatch({ type: "set-step", step: "manager" })}
        onSubmit={handleClientSubmit}
      />
    ) : state.currentStep === "result" && currentDraft ? (
      <CalculationResultStep
        draft={currentDraft}
        errorMessage={stepError}
        isGeneratingFiles={generateFilesMutation.isPending}
        isSaving={saveDraftMutation.isPending}
        lastSaveResult={state.lastSaveResult}
        executionTermsInput={state.executionTermsInput}
        onBack={() => dispatch({ type: "set-step", step: "client" })}
        onCreateNew={handleCreateNewOffer}
        onGenerateFiles={handleGenerateFiles}
        onExecutionTermsChange={(value) => dispatch({ type: "set-execution-terms", value })}
        onSave={handleSave}
        isUpdatingDiscount={updateMetaMutation.isPending}
        onDiscountSubmit={handleDiscountSubmit}
      />
    ) : null;

  return (
    <div style={{ display: "grid", gap: "1.25rem" }}>
      {managersQuery.error && <Alert tone="error">{getErrorMessage(managersQuery.error)}</Alert>}

      <div className="wizard-shell">
        <div className="wizard-main">{currentStepContent}</div>

        <aside className="wizard-sidebar">
          <div className="wizard-sidebar__inner">
            <WizardProgress currentStep={state.currentStep} />
          </div>
        </aside>
      </div>
    </div>
  );
};
