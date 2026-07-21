import { describe, expect, it } from "vitest";
import { mapLegacyWizardStep, WIZARD_STEP_ORDER } from "@/features/commercial-offer/lib/wizardStepOrder";

describe("wizardStepOrder", () => {
  it("defines 3 canonical wizard steps", () => {
    expect(WIZARD_STEP_ORDER).toEqual(["plates", "client", "result"]);
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
