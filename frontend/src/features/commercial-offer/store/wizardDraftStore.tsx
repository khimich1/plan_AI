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
import { normalizeLineKey, pruneWidePlateActions } from "@/features/commercial-offer/lib/flushThenResolveWidePlates";
import { getDraftBatchCount } from "@/features/commercial-offer/lib/getDraftBatchCount";
import type { WideSlot } from "@/features/commercial-offer/lib/widePlateSlots";
import { draftStorage } from "@/features/commercial-offer/store/draftStorage";
import {
  getProductInputStep,
  getWizardStepOrder,
  mapLegacyWizardStep,
  resolveDraftProductType,
} from "@/features/commercial-offer/lib/wizardStepOrder";
import type {
  CommercialDraftDetails,
  CommercialSaveResult,
  ProductType,
  WizardStepId,
  WizardStoreState,
} from "@/features/commercial-offer/types/commercialOffer";

const mergeWizardStepWithServer = (
  local: WizardStepId,
  server: WizardStepId | string | undefined,
  productType: ProductType,
): WizardStepId => {
  const stepOrder = getWizardStepOrder(productType);
  const normalizedLocal = mapLegacyWizardStep(local);
  const normalizedServer = server ? mapLegacyWizardStep(server) : undefined;

  if (!normalizedServer || !stepOrder.includes(normalizedServer)) {
    return normalizedLocal;
  }
  const li = stepOrder.indexOf(normalizedLocal);
  const si = stepOrder.indexOf(normalizedServer);
  if (li < 0) {
    return normalizedServer;
  }
  if (si < 0) {
    return normalizedLocal;
  }
  const inputStep = getProductInputStep(productType);
  if (normalizedLocal === inputStep && si > li) {
    return normalizedLocal;
  }
  return stepOrder[Math.max(li, si)]!;
};

const rememberWidePlateChoice = (
  actions: WizardStoreState["widePlateActions"],
  lineId: string,
  decision: WizardStoreState["widePlateActions"][string],
): WizardStoreState["widePlateActions"] => {
  const next = { ...actions, [lineId]: decision };
  if (!decision.sourceLine) {
    return next;
  }
  const markKey = normalizeLineKey(decision.sourceLine);
  if (!markKey || markKey === lineId) {
    return next;
  }
  return { ...next, [markKey]: decision };
};

const counterpartySnapshotFrom = (
  state: WizardStoreState,
  metadata: CommercialDraftDetails["metadata"],
): Pick<
  WizardStoreState,
  "counterpartyId" | "counterpartyCode1c" | "counterpartyInn" | "counterpartyKpp"
> => {
  const counterpartyId = metadata.counterparty_id ?? null;
  const keepSnapshot = counterpartyId != null && counterpartyId === state.counterpartyId;
  const cardInn = metadata.client_inn?.trim() || null;
  return {
    counterpartyId,
    counterpartyCode1c: keepSnapshot ? state.counterpartyCode1c : "",
    counterpartyInn: keepSnapshot ? state.counterpartyInn : cardInn,
    counterpartyKpp: keepSnapshot ? state.counterpartyKpp : null,
  };
};

type WizardDraftAction =
  | { type: "set-product-type"; productType: ProductType }
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
        counterpartyId: number | null;
        counterpartyCode1c: string;
        counterpartyInn: string | null;
        counterpartyKpp: string | null;
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
      sourceLine?: string;
    }
  | {
      type: "set-unpriced-action";
      lineId: string;
      action: WizardStoreState["unpricedPlateActions"][string]["action"];
      loadCode: number | null;
    }
  | {
      type: "set-invalid-width-action";
      lineId: string;
      action: WizardStoreState["invalidWidthActions"][string]["action"];
      widthMm: number | null;
    }
  | { type: "set-execution-terms"; value: string }
  | { type: "set-draft-id"; draftId: string }
  | { type: "set-page-wide-slots"; pageId: string; slots: WideSlot[] }
  | { type: "hydrate-draft"; payload: CommercialDraftDetails; refreshBatchText?: boolean }
  | { type: "hold-wide-plate-choices" }
  | { type: "release-wide-plate-choices" }
  | { type: "prune-wide-plate-choices"; pageText: string }
  | { type: "settle-wide-plate-page"; pageId: string; keys: string[] }
  | { type: "sync-after-wide-plates"; payload: CommercialDraftDetails }
  | { type: "sync-after-unpriced-plates"; payload: CommercialDraftDetails }
  | { type: "sync-after-invalid-widths"; payload: CommercialDraftDetails }
  | { type: "set-save-result"; payload: CommercialSaveResult | null }
  | { type: "start-append-cycle" }
  | { type: "cancel-append-pick" }
  | { type: "reset" };

const initialState: WizardStoreState = {
  productType: "plates",
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
  counterpartyId: null,
  counterpartyCode1c: "",
  counterpartyInn: null,
  counterpartyKpp: null,
  discountPercent: 0,
  conditionsMode: "standard",
  deliveryConditions: "",
  paymentConditions: "",
  executionTermsInput: "",
  widePlateActions: {},
  widePlateChoicesHeld: false,
  widePlateSettledKeysByPage: {},
  wideSlotsByPage: {},
  unpricedPlateActions: {},
  invalidWidthActions: {},
  lastDraft: null,
  lastSaveResult: null,
  isPickingProductType: false,
};

const reducer = (state: WizardStoreState, action: WizardDraftAction): WizardStoreState => {
  switch (action.type) {
    case "set-product-type":
      return {
        ...state,
        productType: resolveDraftProductType(action.productType),
        currentStep: getProductInputStep(action.productType),
        isPickingProductType: false,
      };
    case "start-append-cycle":
      return {
        ...state,
        sourceText: "",
        selectedImageName: null,
        normalizedText: "",
        batchReviewText: "",
        pendingBatchReview: false,
        confirmedBatchCount: 0,
        widePlateActions: {},
        widePlateChoicesHeld: false,
        widePlateSettledKeysByPage: {},
        unpricedPlateActions: {},
        invalidWidthActions: {},
        isPickingProductType: true,
      };
    case "cancel-append-pick":
      return {
        ...state,
        isPickingProductType: false,
        currentStep: "result",
      };
    case "set-step":
      return { ...state, currentStep: mapLegacyWizardStep(action.step) };
    case "set-source":
      return { ...state, sourceText: action.text, selectedImageName: action.imageName };
    case "set-normalized-text":
      return { ...state, normalizedText: action.text };
    case "set-batch-review-text":
      return { ...state, batchReviewText: action.text };
    case "start-batch-review": {
      const productType = resolveDraftProductType(action.payload.metadata.product_type ?? state.productType);
      const batchCount = getDraftBatchCount(action.payload);
      return {
        ...state,
        productType,
        draftId: action.payload.draft_id,
        currentStep: mergeWizardStepWithServer(
          state.currentStep,
          action.payload.wizard_state?.current_step,
          productType,
        ),
        managerId: action.payload.metadata.manager_id,
        clientName: action.payload.metadata.client_name,
        ...counterpartySnapshotFrom(state, action.payload.metadata),
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
        widePlateChoicesHeld: false,
        widePlateSettledKeysByPage: {},
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
        widePlateActions: rememberWidePlateChoice(state.widePlateActions, action.lineId, {
          action: action.action,
          replacementText: action.replacementText,
          ...(action.sourceLine ? { sourceLine: action.sourceLine } : {}),
        }),
      };
    case "set-unpriced-action":
      return {
        ...state,
        unpricedPlateActions: {
          ...state.unpricedPlateActions,
          [action.lineId]: {
            action: action.action,
            loadCode: action.loadCode,
          },
        },
      };
    case "set-invalid-width-action":
      return {
        ...state,
        invalidWidthActions: {
          ...state.invalidWidthActions,
          [action.lineId]: {
            action: action.action,
            widthMm: action.widthMm,
          },
        },
      };
    case "set-execution-terms":
      return { ...state, executionTermsInput: action.value };
    case "set-draft-id":
      return { ...state, draftId: action.draftId };
    case "set-page-wide-slots":
      return {
        ...state,
        wideSlotsByPage: {
          ...state.wideSlotsByPage,
          [action.pageId]: action.slots,
        },
      };
    case "hold-wide-plate-choices":
      return { ...state, widePlateChoicesHeld: true };
    case "release-wide-plate-choices":
      return { ...state, widePlateChoicesHeld: false };
    case "settle-wide-plate-page":
      return {
        ...state,
        widePlateSettledKeysByPage: {
          ...state.widePlateSettledKeysByPage,
          [action.pageId]: action.keys,
        },
      };
    case "prune-wide-plate-choices": {
      const widePlateActions = pruneWidePlateActions(state.widePlateActions, action.pageText);
      if (widePlateActions === state.widePlateActions) {
        return state;
      }
      const stillHeld = Object.values(widePlateActions).some(
        (decision) => decision.action === "replace" || decision.action === "exclude",
      );
      return {
        ...state,
        widePlateActions,
        widePlateChoicesHeld: stillHeld ? state.widePlateChoicesHeld : false,
      };
    }
    case "hydrate-draft": {
      const productType = resolveDraftProductType(action.payload.metadata.product_type ?? state.productType);
      const batchReviewPending = needsBatchReview(action.payload, state.confirmedBatchCount);
      const batchCount = getDraftBatchCount(action.payload);
      const autoRefreshBatchText =
        batchReviewPending || (state.pendingBatchReview && batchCount > 0);
      const shouldRefreshBatchText =
        Boolean(action.refreshBatchText) || (!state.widePlateChoicesHeld && autoRefreshBatchText);
      return {
        ...state,
        productType,
        draftId: action.payload.draft_id,
        currentStep: mergeWizardStepWithServer(
          state.currentStep,
          action.payload.wizard_state?.current_step,
          productType,
        ),
        managerId: action.payload.metadata.manager_id,
        clientName: action.payload.metadata.client_name,
        ...counterpartySnapshotFrom(state, action.payload.metadata),
        discountPercent: action.payload.metadata.discount_percent,
        conditionsMode: action.payload.metadata.conditions_mode,
        deliveryConditions: action.payload.metadata.delivery_conditions,
        paymentConditions: action.payload.metadata.payment_conditions,
        executionTermsInput:
          (action.payload.metadata.execution_terms || action.payload.saved_offer?.execution_terms || "").trim() ||
          state.executionTermsInput,
        normalizedText: action.payload.metadata.normalized_text ?? "",
        pendingBatchReview: batchReviewPending,
        batchReviewText: shouldRefreshBatchText ? getCurrentBatchReviewText(action.payload) : state.batchReviewText,
        confirmedBatchCount: Math.min(state.confirmedBatchCount, batchCount),
        widePlateActions:
          action.payload.metadata.wide_plates_resolved && !state.widePlateChoicesHeld ? {} : state.widePlateActions,
        unpricedPlateActions: action.payload.metadata.unpriced_plates_resolved ? {} : state.unpricedPlateActions,
        invalidWidthActions: action.payload.metadata.invalid_widths_resolved ? {} : state.invalidWidthActions,
        lastDraft: action.payload,
        lastSaveResult: null,
      };
    }
    case "sync-after-wide-plates": {
      if (state.widePlateChoicesHeld) {
        return state;
      }
      const productType = resolveDraftProductType(action.payload.metadata.product_type ?? state.productType);
      const batchReviewPending = needsBatchReview(action.payload, state.confirmedBatchCount);
      const batchCount = getDraftBatchCount(action.payload);
      return {
        ...state,
        productType,
        draftId: action.payload.draft_id,
        currentStep: mergeWizardStepWithServer(
          state.currentStep,
          action.payload.wizard_state?.current_step,
          productType,
        ),
        normalizedText: action.payload.metadata.normalized_text ?? "",
        pendingBatchReview: batchReviewPending,
        batchReviewText: getCurrentBatchReviewText(action.payload),
        confirmedBatchCount: Math.min(state.confirmedBatchCount, batchCount),
        widePlateActions: {},
        lastDraft: action.payload,
      };
    }
    case "sync-after-unpriced-plates": {
      const productType = resolveDraftProductType(action.payload.metadata.product_type ?? state.productType);
      const batchReviewPending = needsBatchReview(action.payload, state.confirmedBatchCount);
      const batchCount = getDraftBatchCount(action.payload);
      return {
        ...state,
        productType,
        draftId: action.payload.draft_id,
        currentStep: mergeWizardStepWithServer(
          state.currentStep,
          action.payload.wizard_state?.current_step,
          productType,
        ),
        normalizedText: action.payload.metadata.normalized_text ?? "",
        pendingBatchReview: batchReviewPending,
        batchReviewText: getCurrentBatchReviewText(action.payload),
        confirmedBatchCount: Math.min(state.confirmedBatchCount, batchCount),
        unpricedPlateActions: {},
        lastDraft: action.payload,
      };
    }
    case "sync-after-invalid-widths": {
      const productType = resolveDraftProductType(action.payload.metadata.product_type ?? state.productType);
      const batchReviewPending = needsBatchReview(action.payload, state.confirmedBatchCount);
      const batchCount = getDraftBatchCount(action.payload);
      return {
        ...state,
        productType,
        draftId: action.payload.draft_id,
        currentStep: mergeWizardStepWithServer(
          state.currentStep,
          action.payload.wizard_state?.current_step,
          productType,
        ),
        normalizedText: action.payload.metadata.normalized_text ?? "",
        pendingBatchReview: batchReviewPending,
        batchReviewText: getCurrentBatchReviewText(action.payload),
        confirmedBatchCount: Math.min(state.confirmedBatchCount, batchCount),
        invalidWidthActions: {},
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
    const productType = resolveDraftProductType(loaded.productType);
    return {
      ...value,
      ...loaded,
      productType,
      currentStep: mapLegacyWizardStep(loaded.currentStep),
      normalizedText: loaded.normalizedText ?? "",
      isPickingProductType: loaded.isPickingProductType ?? false,
      widePlateChoicesHeld: false,
      widePlateSettledKeysByPage: loaded.widePlateSettledKeysByPage ?? {},
      wideSlotsByPage: loaded.wideSlotsByPage ?? {},
      unpricedPlateActions: loaded.unpricedPlateActions ?? {},
      invalidWidthActions: loaded.invalidWidthActions ?? {},
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
