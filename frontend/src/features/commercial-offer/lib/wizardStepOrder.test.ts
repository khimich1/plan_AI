import { describe, expect, it } from "vitest";
import {
  getProductInputStep,
  getWizardStepOrder,
  mapLegacyWizardStep,
  PLATES_WIZARD_STEP_ORDER,
  PILES_WIZARD_STEP_ORDER,
} from "@/features/commercial-offer/lib/wizardStepOrder";

describe("wizardStepOrder", () => {
  it("defines plate and pile step orders", () => {
    expect(PLATES_WIZARD_STEP_ORDER).toEqual(["plates", "client", "result"]);
    expect(PILES_WIZARD_STEP_ORDER).toEqual(["piles", "client", "result"]);
  });

  it("returns product-specific step order", () => {
    expect(getWizardStepOrder("plates")).toEqual(PLATES_WIZARD_STEP_ORDER);
    expect(getWizardStepOrder("piles")).toEqual(PILES_WIZARD_STEP_ORDER);
  });

  it("maps piles step id", () => {
    expect(mapLegacyWizardStep("piles")).toBe("piles");
    expect(getProductInputStep("piles")).toBe("piles");
  });

  it("maps legacy wide-plates and manager steps", () => {
    expect(mapLegacyWizardStep("wide-plates")).toBe("plates");
    expect(mapLegacyWizardStep("manager")).toBe("client");
    expect(mapLegacyWizardStep("calculate")).toBe("client");
  });

  it("falls back to plates for unknown values", () => {
    expect(mapLegacyWizardStep("not-a-real-step")).toBe("plates");
    expect(mapLegacyWizardStep("")).toBe("plates");
  });
});
