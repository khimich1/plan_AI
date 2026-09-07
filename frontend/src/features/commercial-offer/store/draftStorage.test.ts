import { afterEach, describe, expect, it, vi } from "vitest";

import { draftStorage } from "@/features/commercial-offer/store/draftStorage";
import type { WizardStoreState } from "@/features/commercial-offer/types/commercialOffer";

const makeState = (): WizardStoreState => ({
  productType: "plates",
  draftId: "draft-1",
  currentStep: "plates",
  sourceText: "ПБ 78-12-8п 2",
  selectedImageName: null,
  normalizedText: "",
  batchReviewText: "",
  pendingBatchReview: false,
  confirmedBatchCount: 0,
  lastPlateMode: "replace",
  managerId: null,
  clientName: "",
  discountPercent: 0,
  conditionsMode: "standard",
  deliveryConditions: "",
  paymentConditions: "",
  executionTermsInput: "",
  widePlateActions: {},
  unpricedPlateActions: {},
  invalidWidthActions: {},
  lastDraft: null,
  lastSaveResult: null,
  isPickingProductType: false,
});

afterEach(() => {
  vi.restoreAllMocks();
  window.sessionStorage.clear();
});

describe("draftStorage", () => {
  it("round-trips state through sessionStorage", () => {
    const state = makeState();
    draftStorage.save(state);
    expect(draftStorage.load()).toEqual(state);
  });

  it("load returns null when nothing is stored", () => {
    expect(draftStorage.load()).toBeNull();
  });

  it("load returns null on corrupt JSON instead of throwing", () => {
    window.sessionStorage.setItem("commercial-offer-wizard:v1", "{not-json");
    expect(draftStorage.load()).toBeNull();
  });

  it("save swallows QuotaExceededError instead of crashing the wizard", () => {
    vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => {
      throw new DOMException("quota exceeded", "QuotaExceededError");
    });
    expect(() => draftStorage.save(makeState())).not.toThrow();
  });

  it("clear swallows storage failures", () => {
    vi.spyOn(Storage.prototype, "removeItem").mockImplementation(() => {
      throw new DOMException("storage unavailable", "SecurityError");
    });
    expect(() => draftStorage.clear()).not.toThrow();
  });
});
