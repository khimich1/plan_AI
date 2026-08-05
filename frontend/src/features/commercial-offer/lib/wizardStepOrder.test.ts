import { describe, expect, it } from "vitest";
import {
  BRIDGE_PILES_WIZARD_STEP_ORDER,
  FBS_WIZARD_STEP_ORDER,
  getProductInputStep,
  getWizardStepOrder,
  isSimpleKpProductType,
  mapLegacyWizardStep,
  MARCHES_WIZARD_STEP_ORDER,
  PLATES_WIZARD_STEP_ORDER,
  PILES_WIZARD_STEP_ORDER,
  STEPS_WIZARD_STEP_ORDER,
} from "@/features/commercial-offer/lib/wizardStepOrder";

describe("wizardStepOrder", () => {
  it("defines plate, pile, step, march, bridge pile, and fbs step orders", () => {
    expect(PLATES_WIZARD_STEP_ORDER).toEqual(["plates", "client", "result"]);
    expect(PILES_WIZARD_STEP_ORDER).toEqual(["piles", "client", "result"]);
    expect(STEPS_WIZARD_STEP_ORDER).toEqual(["steps", "client", "result"]);
    expect(MARCHES_WIZARD_STEP_ORDER).toEqual(["marches", "client", "result"]);
    expect(BRIDGE_PILES_WIZARD_STEP_ORDER).toEqual(["bridge_piles", "client", "result"]);
    expect(FBS_WIZARD_STEP_ORDER).toEqual(["fbs", "client", "result"]);
  });

  it("returns product-specific step order", () => {
    expect(getWizardStepOrder("plates")).toEqual(PLATES_WIZARD_STEP_ORDER);
    expect(getWizardStepOrder("piles")).toEqual(PILES_WIZARD_STEP_ORDER);
    expect(getWizardStepOrder("steps")).toEqual(STEPS_WIZARD_STEP_ORDER);
    expect(getWizardStepOrder("marches")).toEqual(MARCHES_WIZARD_STEP_ORDER);
    expect(getWizardStepOrder("bridge_piles")).toEqual(BRIDGE_PILES_WIZARD_STEP_ORDER);
    expect(getWizardStepOrder("fbs")).toEqual(FBS_WIZARD_STEP_ORDER);
  });

  it("maps piles, steps, marches, bridge_piles, and fbs step ids", () => {
    expect(mapLegacyWizardStep("piles")).toBe("piles");
    expect(getProductInputStep("piles")).toBe("piles");
    expect(mapLegacyWizardStep("steps")).toBe("steps");
    expect(getProductInputStep("steps")).toBe("steps");
    expect(mapLegacyWizardStep("marches")).toBe("marches");
    expect(getProductInputStep("marches")).toBe("marches");
    expect(mapLegacyWizardStep("bridge_piles")).toBe("bridge_piles");
    expect(getProductInputStep("bridge_piles")).toBe("bridge_piles");
    expect(mapLegacyWizardStep("fbs")).toBe("fbs");
    expect(getProductInputStep("fbs")).toBe("fbs");
  });

  it("marks piles, steps, marches, bridge_piles, and fbs as simple KP product types", () => {
    expect(isSimpleKpProductType("piles")).toBe(true);
    expect(isSimpleKpProductType("steps")).toBe(true);
    expect(isSimpleKpProductType("marches")).toBe(true);
    expect(isSimpleKpProductType("bridge_piles")).toBe(true);
    expect(isSimpleKpProductType("fbs")).toBe(true);
    expect(isSimpleKpProductType("plates")).toBe(false);
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
