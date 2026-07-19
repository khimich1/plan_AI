import type { LegacyWizardStepId, WizardStepId } from "@/features/commercial-offer/types/commercialOffer";

export const WIZARD_STEP_ORDER: WizardStepId[] = ["plates", "client", "result"];

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
  if ((WIZARD_STEP_ORDER as string[]).includes(raw)) {
    return raw as WizardStepId;
  }
  return "plates";
};

export const wizardStepIndex = (step: WizardStepId): number => WIZARD_STEP_ORDER.indexOf(step);
