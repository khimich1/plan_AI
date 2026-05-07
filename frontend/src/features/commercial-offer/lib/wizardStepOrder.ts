import type { WizardStepId } from "@/features/commercial-offer/types/commercialOffer";

export const WIZARD_STEP_ORDER: WizardStepId[] = ["plates", "wide-plates", "manager", "client", "result"];

export const wizardStepIndex = (step: WizardStepId): number => WIZARD_STEP_ORDER.indexOf(step);
