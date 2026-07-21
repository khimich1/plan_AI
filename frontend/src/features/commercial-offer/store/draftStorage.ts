import type { WizardStoreState } from "@/features/commercial-offer/types/commercialOffer";

const STORAGE_KEY = "commercial-offer-wizard:v1";

export const draftStorage = {
  load(): WizardStoreState | null {
    try {
      const raw = window.sessionStorage.getItem(STORAGE_KEY);
      if (!raw) {
        return null;
      }
      return JSON.parse(raw) as WizardStoreState;
    } catch {
      return null;
    }
  },

  save(state: WizardStoreState): void {
    window.sessionStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  },

  clear(): void {
    window.sessionStorage.removeItem(STORAGE_KEY);
  },
};
