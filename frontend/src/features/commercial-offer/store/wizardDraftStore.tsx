import {
  createContext,
  useContext,
  useEffect,
  useMemo,
  useReducer,
  type Dispatch,
  type PropsWithChildren,
} from "react";
import { getCurrentBatchReviewText, needsBatchReview } from "@/features/commercial-offer/lib/batchReview";
import { draftStorage } from "@/features/commercial-offer/store/draftStorage";
import { mapLegacyWizardStep, WIZARD_STEP_ORDER } from "@/features/commercial-offer/lib/wizardStepOrder";
import type { CommercialDraftDetails, CommercialSaveResult, WizardStepId, WizardStoreState } from "@/features/commercial-offer/types/commercialOffer";

const mergeWizardStepWithServer = (local: WizardStepId, server: WizardStepId | string | undefined): WizardStepId => {
  const normalizedLocal = mapLegacyWizardStep(local);
  const normalizedServer = server ? mapLegacyWizardStep(server) : undefined;

  if (!normalizedServer || !WIZARD_STEP_ORDER.includes(normalizedServer)) {
    return normalizedLocal;
  }
  const li = WIZARD_STEP_ORDER.indexOf(normalizedLocal);
  const si = WIZARD_STEP_ORDER.indexOf(normalizedServer);
  if (li < 0) {
    return normalizedServer;
  }
  if (si < 0) {
    return normalizedLocal;
  }
  // After recognize, stay on plates until user clicks "Обработать"
  if (normalizedLocal === "plates" && si > li) {
    return normalizedLocal;
  }
  return WIZARD_STEP_ORDER[Math.max(li, si)];
};

type WizardDraftAction =
  | { type: "set-step"; step: WizardStepId }
  | { type: "set-source"; text: string; imageName: string | null }
  | { type: "set-normalized-text"; text: string }
  | { type: "set-batch-review-text"; text: string }
  | { type: "start-batch-review"; payload: CommercialDraftDetails }
  | { type: "confirm-batch-review"; batchCount: number }
  | { type: "set-manager"; managerId: number | null }
  | {
      type: "set-client-form";
      payload: {
        clientName: string;
        conditionsMode: WizardStoreState["conditionsMode"];
        deliveryConditions: string;
        paymentConditions: string;
      };
    }
  | {
      type: "set-wide-action";
      lineId: string;
      action: WizardStoreState["widePlateActions"][string]["action"];
      replacementText: string;
    }
  | { type: "set-execution-terms"; value: string }
  | { type: "set-draft-id"; draftId: string }
  | { type: "hydrate-draft"; payload: CommercialDraftDetails; refreshBatchText?: boolean }
  | { type: "sync-after-wide-plates"; payload: CommercialDraftDetails }
  | { type: "set-save-result"; payload: CommercialSaveResult | null }
  | { type: "reset" };

const initialState: WizardStoreState = {
  draftId: null,
  currentStep: "plates",
  sourceText: "",
  selectedImageName: null,
  normalizedText: "",
  batchReviewText: "",
  pendingBatchReview: false,
  confirmedBatchCount: 0,
  lastPlateMode: "replace",
  managerId: null,
  clientName: "",
  discountPercent: 0,
  conditionsMode: "standard",
  deliveryConditions: "",
  paymentConditions: "",
  executionTermsInput: "",
  widePlateActions: {},
  lastDraft: null,
  lastSaveResult: null,
};

const reducer = (state: WizardStoreState, action: WizardDraftAction): WizardStoreState => {
  switch (action.type) {
    case "set-step":
      return { ...state, currentStep: mapLegacyWizardStep(action.step) };
    case "set-source":
      return { ...state, sourceText: action.text, selectedImageName: action.imageName };
    case "set-normalized-text":
      return { ...state, normalizedText: action.text };
    case "set-batch-review-text":
      return { ...state, batchReviewText: action.text };
    case "start-batch-review": {
      const batchCount = action.payload.metadata.plate_batches?.length ?? 0;
      return {
        ...state,
        draftId: action.payload.draft_id,
        currentStep: mergeWizardStepWithServer(
          state.currentStep,
          action.payload.wizard_state?.current_step,
        ),
        managerId: action.payload.metadata.manager_id,
        clientName: action.payload.metadata.client_name,
        discountPercent: action.payload.metadata.discount_percent,
        conditionsMode: action.payload.metadata.conditions_mode,
        deliveryConditions: action.payload.metadata.delivery_conditions,
        paymentConditions: action.payload.metadata.payment_conditions,
        executionTermsInput:
          (action.payload.metadata.execution_terms || action.payload.saved_offer?.execution_terms || "").trim() ||
          state.executionTermsInput,
        normalizedText: action.payload.metadata.normalized_text ?? "",
        pendingBatchReview: batchCount > 0,
        batchReviewText: getCurrentBatchReviewText(action.payload),
        confirmedBatchCount: Math.max(0, batchCount - 1),
        lastDraft: action.payload,
      };
    }
    case "confirm-batch-review":
      return {
        ...state,
        pendingBatchReview: false,
        confirmedBatchCount: action.batchCount,
      };
    case "set-manager":
      return { ...state, managerId: action.managerId };
    case "set-client-form":
      return { ...state, ...action.payload };
    case "set-wide-action":
      return {
        ...state,
        widePlateActions: {
          ...state.widePlateActions,
          [action.lineId]: {
            action: action.action,
            replacementText: action.replacementText,
          },
        },
      };
    case "set-execution-terms":
      return { ...state, executionTermsInput: action.value };
    case "set-draft-id":
      return { ...state, draftId: action.draftId };
    case "hydrate-draft": {
      const batchReviewPending = needsBatchReview(action.payload, state.confirmedBatchCount);
      const shouldRefreshBatchText =
        action.refreshBatchText ||
        batchReviewPending ||
        (state.pendingBatchReview && (action.payload.metadata.plate_batches?.length ?? 0) > 0);
      return {
        ...state,
        draftId: action.payload.draft_id,
        currentStep: mergeWizardStepWithServer(
          state.currentStep,
          action.payload.wizard_state?.current_step,
        ),
        managerId: action.payload.metadata.manager_id,
        clientName: action.payload.metadata.client_name,
        discountPercent: action.payload.metadata.discount_percent,
        conditionsMode: action.payload.metadata.conditions_mode,
        deliveryConditions: action.payload.metadata.delivery_conditions,
        paymentConditions: action.payload.metadata.payment_conditions,
        executionTermsInput:
          (action.payload.metadata.execution_terms || action.payload.saved_offer?.execution_terms || "").trim() ||
          state.executionTermsInput,
        normalizedText: action.payload.metadata.normalized_text ?? "",
        pendingBatchReview: batchReviewPending,
        batchReviewText: shouldRefreshBatchText
          ? getCurrentBatchReviewText(action.payload)
          : state.batchReviewText,
        confirmedBatchCount: Math.min(
          state.confirmedBatchCount,
          action.payload.metadata.plate_batches?.length ?? 0,
        ),
        widePlateActions: action.payload.metadata.wide_plates_resolved ? {} : state.widePlateActions,
        lastDraft: action.payload,
      };
    }
    case "sync-after-wide-plates": {
      const batchReviewPending = needsBatchReview(action.payload, state.confirmedBatchCount);
      return {
        ...state,
        draftId: action.payload.draft_id,
        currentStep: mergeWizardStepWithServer(
          state.currentStep,
          action.payload.wizard_state?.current_step,
        ),
        normalizedText: action.payload.metadata.normalized_text ?? "",
        pendingBatchReview: batchReviewPending,
        batchReviewText: getCurrentBatchReviewText(action.payload),
        confirmedBatchCount: Math.min(
          state.confirmedBatchCount,
          action.payload.metadata.plate_batches?.length ?? 0,
        ),
        widePlateActions: {},
        lastDraft: action.payload,
      };
    }
    case "set-save-result":
      return { ...state, lastSaveResult: action.payload };
    case "reset":
      return initialState;
    default:
      return state;
  }
};

type WizardDraftContextValue = {
  state: WizardStoreState;
  dispatch: Dispatch<WizardDraftAction>;
};

const WizardDraftContext = createContext<WizardDraftContextValue | null>(null);

export const WizardDraftProvider = ({ children }: PropsWithChildren) => {
  const [state, dispatch] = useReducer(reducer, initialState, (value) => {
    const loaded = draftStorage.load();
    if (!loaded) {
      return value;
    }
    return {
      ...value,
      ...loaded,
      currentStep: mapLegacyWizardStep(loaded.currentStep),
      normalizedText: loaded.normalizedText ?? "",
    };
  });

  useEffect(() => {
    draftStorage.save(state);
  }, [state]);

  const contextValue = useMemo(() => ({ state, dispatch }), [state]);

  return <WizardDraftContext.Provider value={contextValue}>{children}</WizardDraftContext.Provider>;
};

export const useWizardDraftStore = (): WizardDraftContextValue => {
  const context = useContext(WizardDraftContext);
  if (!context) {
    throw new Error("useWizardDraftStore must be used within WizardDraftProvider");
  }
  return context;
};
