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
    // Persistence is best-effort: a full draft (OCR texts) can exceed the
    // sessionStorage quota, and storage may be unavailable in private mode.
    try {
      window.sessionStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch {
      // ignore — losing the draft snapshot must not break the wizard
    }
  },

  clear(): void {
    try {
      window.sessionStorage.removeItem(STORAGE_KEY);
    } catch {
      // ignore — same rationale as save()
    }
  },
};
