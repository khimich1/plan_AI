import {
  createContext,
  useCallback,
  useContext,
  useMemo,
  type PropsWithChildren,
} from "react";
import type { WizardStoreState } from "@/features/commercial-offer/types/commercialOffer";
import { useWizardDraftStore } from "@/features/commercial-offer/store/wizardDraftStore";
import { draftStorage } from "@/features/commercial-offer/store/draftStorage";

export type CommercialDraftHeaderBridge = {
  /** True when navigating to «Создать КП» should confirm discarding draft data */
  hasDraft: boolean;
  /** Clears wizard state and persisted draft storage */
  resetDraft: () => void;
};

const noop = () => {};

const defaultBridge: CommercialDraftHeaderBridge = {
  hasDraft: false,
  resetDraft: noop,
};

const CommercialDraftHeaderBridgeContext = createContext<CommercialDraftHeaderBridge>(defaultBridge);

const computeHasDraft = (state: WizardStoreState): boolean => {
  if (state.draftId) {
    return true;
  }
  if (state.sourceText && state.sourceText.trim().length > 0) {
    return true;
  }
  if (state.selectedImageName) {
    return true;
  }
  if (state.managerId || state.clientName || state.lastSaveResult) {
    return true;
  }
  return false;
};

/** Bridges wizard draft state to the app shell header (mounted from AppLayout). */
export const CommercialOfferHeaderBridgeProvider = ({ children }: PropsWithChildren) => {
  const { state, dispatch } = useWizardDraftStore();

  const resetDraft = useCallback(() => {
    dispatch({ type: "reset" });
    draftStorage.clear();
  }, [dispatch]);

  const value = useMemo(
    () => ({
      hasDraft: computeHasDraft(state),
      resetDraft,
    }),
    [state, resetDraft],
  );

  return (
    <CommercialDraftHeaderBridgeContext.Provider value={value}>
      {children}
    </CommercialDraftHeaderBridgeContext.Provider>
  );
};

export function useCommercialDraftHeaderBridge(): CommercialDraftHeaderBridge {
  return useContext(CommercialDraftHeaderBridgeContext);
}
