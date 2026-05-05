import {
  createContext,
  useContext,
  useEffect,
  useMemo,
  useReducer,
  type Dispatch,
  type PropsWithChildren,
} from "react";
import { draftStorage } from "@/features/commercial-offer/store/draftStorage";
import type { CommercialDraftDetails, CommercialSaveResult, WizardStepId, WizardStoreState } from "@/features/commercial-offer/types/commercialOffer";

type WizardDraftAction =
  | { type: "set-step"; step: WizardStepId }
  | { type: "set-source"; text: string; imageName: string | null }
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
  | { type: "hydrate-draft"; payload: CommercialDraftDetails }
  | { type: "set-save-result"; payload: CommercialSaveResult | null }
  | { type: "reset" };

const initialState: WizardStoreState = {
  draftId: null,
  currentStep: "plates",
  sourceText: "",
  selectedImageName: null,
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
      return { ...state, currentStep: action.step };
    case "set-source":
      return { ...state, sourceText: action.text, selectedImageName: action.imageName };
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
    case "hydrate-draft":
      return {
        ...state,
        draftId: action.payload.draft_id,
        currentStep: state.currentStep,
        managerId: action.payload.metadata.manager_id,
        clientName: action.payload.metadata.client_name,
        discountPercent: action.payload.metadata.discount_percent,
        conditionsMode: action.payload.metadata.conditions_mode,
        deliveryConditions: action.payload.metadata.delivery_conditions,
        paymentConditions: action.payload.metadata.payment_conditions,
        executionTermsInput:
          (action.payload.metadata.execution_terms || action.payload.saved_offer?.execution_terms || "").trim() ||
          state.executionTermsInput,
        lastDraft: action.payload,
      };
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
  const [state, dispatch] = useReducer(reducer, initialState, (value) => draftStorage.load() ?? value);

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
