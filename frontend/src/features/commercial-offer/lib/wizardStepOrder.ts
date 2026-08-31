import type { LegacyWizardStepId, ProductType, WizardStepId } from "@/features/commercial-offer/types/commercialOffer";

export const PLATES_WIZARD_STEP_ORDER: WizardStepId[] = ["plates", "client", "result"];
export const PILES_WIZARD_STEP_ORDER: WizardStepId[] = ["piles", "client", "result"];
export const STEPS_WIZARD_STEP_ORDER: WizardStepId[] = ["steps", "client", "result"];
export const MARCHES_WIZARD_STEP_ORDER: WizardStepId[] = ["marches", "client", "result"];
export const BRIDGE_PILES_WIZARD_STEP_ORDER: WizardStepId[] = ["bridge_piles", "client", "result"];
export const FBS_WIZARD_STEP_ORDER: WizardStepId[] = ["fbs", "client", "result"];

/** @deprecated use getWizardStepOrder(productType) */
export const WIZARD_STEP_ORDER = PLATES_WIZARD_STEP_ORDER;

export type SkipClientStepInput = {
  clientName?: string | null;
  appendBatches?: ReadonlyArray<unknown> | null;
  resumeKpId?: number | null;
};

export type WizardStepOrderOptions = {
  skipClient?: boolean;
};

/** Aligns with BE CommercialWizardStepService.should_skip_client_step. */
export const shouldSkipClientStep = (input: SkipClientStepInput): boolean => {
  const clientName = String(input.clientName ?? "").trim();
  if (clientName) {
    return true;
  }
  const appendBatches = input.appendBatches ?? [];
  if (appendBatches.length > 0) {
    return true;
  }
  return input.resumeKpId != null;
};

export const getProductInputStep = (productType: ProductType): WizardStepId => {
  if (productType === "piles") {
    return "piles";
  }
  if (productType === "steps") {
    return "steps";
  }
  if (productType === "marches") {
    return "marches";
  }
  if (productType === "bridge_piles") {
    return "bridge_piles";
  }
  if (productType === "fbs") {
    return "fbs";
  }
  return "plates";
};

const fullWizardStepOrder = (productType: ProductType): WizardStepId[] => {
  if (productType === "piles") {
    return PILES_WIZARD_STEP_ORDER;
  }
  if (productType === "steps") {
    return STEPS_WIZARD_STEP_ORDER;
  }
  if (productType === "marches") {
    return MARCHES_WIZARD_STEP_ORDER;
  }
  if (productType === "bridge_piles") {
    return BRIDGE_PILES_WIZARD_STEP_ORDER;
  }
  if (productType === "fbs") {
    return FBS_WIZARD_STEP_ORDER;
  }
  return PLATES_WIZARD_STEP_ORDER;
};

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
  productType === "piles" || productType === "steps" || productType === "marches" || productType === "bridge_piles" || productType === "fbs";

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
  if (raw === "piles" || raw === "plates" || raw === "steps" || raw === "marches" || raw === "bridge_piles" || raw === "fbs" || raw === "client" || raw === "result") {
    return raw;
  }
  return "plates";
};

export const wizardStepIndex = (step: WizardStepId, productType: ProductType = "plates"): number =>
  getWizardStepOrder(productType).indexOf(step);

export const resolveDraftProductType = (productType: ProductType | null | undefined): ProductType => {
  if (productType === "piles" || productType === "steps" || productType === "marches" || productType === "bridge_piles" || productType === "fbs") {
    return productType;
  }
  return "plates";
};
