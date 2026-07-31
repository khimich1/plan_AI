import type { LegacyWizardStepId, ProductType, WizardStepId } from "@/features/commercial-offer/types/commercialOffer";

export const PLATES_WIZARD_STEP_ORDER: WizardStepId[] = ["plates", "client", "result"];
export const PILES_WIZARD_STEP_ORDER: WizardStepId[] = ["piles", "client", "result"];

/** @deprecated use getWizardStepOrder(productType) */
export const WIZARD_STEP_ORDER = PLATES_WIZARD_STEP_ORDER;

export const getProductInputStep = (productType: ProductType): WizardStepId =>
  productType === "piles" ? "piles" : "plates";

export const getWizardStepOrder = (productType: ProductType): WizardStepId[] =>
  productType === "piles" ? PILES_WIZARD_STEP_ORDER : PLATES_WIZARD_STEP_ORDER;

const LEGACY_WIZARD_STEP_MAP: Record<LegacyWizardStepId | "calculate", WizardStepId> = {
  "wide-plates": "plates",
  manager: "client",
  calculate: "client",
};

export const mapLegacyWizardStep = (step: string | null | undefined): WizardStepId => {
  const raw = String(step ?? "").trim().toLowerCase();
  if (!raw) {
    return "plates";
  }
  if (raw in LEGACY_WIZARD_STEP_MAP) {
    return LEGACY_WIZARD_STEP_MAP[raw as LegacyWizardStepId | "calculate"];
  }
  if (raw === "piles" || raw === "plates" || raw === "client" || raw === "result") {
    return raw;
  }
  return "plates";
};

export const wizardStepIndex = (step: WizardStepId, productType: ProductType = "plates"): number =>
  getWizardStepOrder(productType).indexOf(step);

export const resolveDraftProductType = (productType: ProductType | null | undefined): ProductType =>
  productType === "piles" ? "piles" : "plates";
