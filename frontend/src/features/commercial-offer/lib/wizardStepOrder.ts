import type { LegacyWizardStepId, ProductType, WizardStepId } from "@/features/commercial-offer/types/commercialOffer";
import { getProductTypeConfig, PRODUCT_TYPE_CONFIG } from "@/features/commercial-offer/lib/productTypeConfig";

const WIZARD_TRAILING_STEPS: WizardStepId[] = ["client", "result"];

const wizardOrderFor = (productType: ProductType): WizardStepId[] => [
  PRODUCT_TYPE_CONFIG[productType].inputStep,
  ...WIZARD_TRAILING_STEPS,
];

/** Module-level arrays — getWizardStepOrder must return a stable reference when skipClient is false. */
export const PLATES_WIZARD_STEP_ORDER: WizardStepId[] = wizardOrderFor("plates");
export const PILES_WIZARD_STEP_ORDER: WizardStepId[] = wizardOrderFor("piles");
export const STEPS_WIZARD_STEP_ORDER: WizardStepId[] = wizardOrderFor("steps");
export const MARCHES_WIZARD_STEP_ORDER: WizardStepId[] = wizardOrderFor("marches");
export const BRIDGE_PILES_WIZARD_STEP_ORDER: WizardStepId[] = wizardOrderFor("bridge_piles");
export const FBS_WIZARD_STEP_ORDER: WizardStepId[] = wizardOrderFor("fbs");

const FULL_WIZARD_STEP_ORDER: Record<ProductType, WizardStepId[]> = {
  plates: PLATES_WIZARD_STEP_ORDER,
  piles: PILES_WIZARD_STEP_ORDER,
  steps: STEPS_WIZARD_STEP_ORDER,
  marches: MARCHES_WIZARD_STEP_ORDER,
  bridge_piles: BRIDGE_PILES_WIZARD_STEP_ORDER,
  fbs: FBS_WIZARD_STEP_ORDER,
};

export type SkipClientStepInput = {
  clientName?: string | null;
  counterpartyId?: number | null;
  appendBatches?: ReadonlyArray<unknown> | null;
  resumeKpId?: number | null;
};

export type WizardStepOrderOptions = {
  skipClient?: boolean;
};

/** Aligns with BE CommercialWizardStepService.should_skip_client_step:
 * new KP skips only when counterpartyId is set; append/resume unchanged.
 * Free-text clientName does not skip. */
export const shouldSkipClientStep = (input: SkipClientStepInput): boolean => {
  const appendBatches = input.appendBatches ?? [];
  if (appendBatches.length > 0) {
    return true;
  }
  if (input.resumeKpId != null) {
    return true;
  }
  return input.counterpartyId != null;
};

export const getProductInputStep = (productType: ProductType): WizardStepId =>
  getProductTypeConfig(productType).inputStep;

const fullWizardStepOrder = (productType: ProductType): WizardStepId[] =>
  FULL_WIZARD_STEP_ORDER[getProductTypeConfig(productType).productType];

export const getWizardStepOrder = (
  productType: ProductType,
  options?: WizardStepOrderOptions,
): WizardStepId[] => {
  const order = fullWizardStepOrder(productType);
  if (options?.skipClient) {
    return order.filter((step) => step !== "client");
  }
  return order;
};

export const isSimpleKpProductType = (productType: ProductType): boolean =>
  getProductTypeConfig(productType).isSimpleKp;

const LEGACY_WIZARD_STEP_MAP: Record<LegacyWizardStepId | "calculate", WizardStepId> = {
  "wide-plates": "plates",
  manager: "client",
  calculate: "client",
};

const WIZARD_STEP_IDS: ReadonlySet<string> = new Set([
  ...Object.keys(PRODUCT_TYPE_CONFIG),
  "client",
  "result",
]);

export const mapLegacyWizardStep = (step: string | null | undefined): WizardStepId => {
  const raw = String(step ?? "").trim().toLowerCase();
  if (!raw) {
    return "plates";
  }
  if (Object.prototype.hasOwnProperty.call(LEGACY_WIZARD_STEP_MAP, raw)) {
    return LEGACY_WIZARD_STEP_MAP[raw as LegacyWizardStepId | "calculate"];
  }
  if (WIZARD_STEP_IDS.has(raw)) {
    return raw as WizardStepId;
  }
  return "plates";
};

export const wizardStepIndex = (step: WizardStepId, productType: ProductType = "plates"): number =>
  getWizardStepOrder(productType).indexOf(step);

/**
 * From Result (or after Result was reached), the product input step is only
 * reachable via a new append cycle — not via sidebar / «Назад».
 */
export const isInputStepBlockedWithoutAppendCycle = ({
  currentStep,
  targetStep,
  inputStep,
  draftWizardStep,
}: {
  currentStep: WizardStepId;
  targetStep: WizardStepId;
  inputStep: WizardStepId;
  draftWizardStep?: WizardStepId | null;
}): boolean => {
  if (targetStep !== inputStep) {
    return false;
  }
  if (currentStep === inputStep) {
    return false;
  }
  if (currentStep === "result") {
    return true;
  }
  return draftWizardStep === "result";
};

export const resolveDraftProductType = (productType: ProductType | null | undefined): ProductType =>
  getProductTypeConfig(productType).productType;
