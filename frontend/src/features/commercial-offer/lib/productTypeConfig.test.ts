import { describe, expect, it } from "vitest";
import {
  getProductTypeConfig,
  INGEST_REQUIRED_MESSAGE,
  PRODUCT_TYPE_CONFIG,
  RESOLVE_GATE_MESSAGES,
} from "@/features/commercial-offer/lib/productTypeConfig";
import type { ProductType } from "@/features/commercial-offer/types/commercialOffer";

const ALL_PRODUCT_TYPES: ProductType[] = ["plates", "piles", "steps", "marches", "bridge_piles", "fbs"];

describe("PRODUCT_TYPE_CONFIG", () => {
  it("covers every ProductType exactly once", () => {
    expect(Object.keys(PRODUCT_TYPE_CONFIG).sort()).toEqual([...ALL_PRODUCT_TYPES].sort());
    for (const type of ALL_PRODUCT_TYPES) {
      expect(PRODUCT_TYPE_CONFIG[type].productType).toBe(type);
    }
  });

  it("maps each product to its REST endpoint segment", () => {
    expect(PRODUCT_TYPE_CONFIG.plates.endpointSegment).toBe("plates");
    expect(PRODUCT_TYPE_CONFIG.piles.endpointSegment).toBe("piles");
    expect(PRODUCT_TYPE_CONFIG.steps.endpointSegment).toBe("steps");
    expect(PRODUCT_TYPE_CONFIG.marches.endpointSegment).toBe("marches");
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.endpointSegment).toBe("bridge-piles");
    expect(PRODUCT_TYPE_CONFIG.fbs.endpointSegment).toBe("fbs");
  });

  it("uses the product type itself as the wizard input step", () => {
    for (const type of ALL_PRODUCT_TYPES) {
      expect(PRODUCT_TYPE_CONFIG[type].inputStep).toBe(type);
    }
  });

  it("marks grade-supporting products (piles, marches, bridge_piles, fbs)", () => {
    expect(PRODUCT_TYPE_CONFIG.plates.supportsGrades).toBe(false);
    expect(PRODUCT_TYPE_CONFIG.steps.supportsGrades).toBe(false);
    expect(PRODUCT_TYPE_CONFIG.piles.supportsGrades).toBe(true);
    expect(PRODUCT_TYPE_CONFIG.marches.supportsGrades).toBe(true);
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.supportsGrades).toBe(true);
    expect(PRODUCT_TYPE_CONFIG.fbs.supportsGrades).toBe(true);
  });

  it("marks every product except plates as a simple KP flow", () => {
    expect(PRODUCT_TYPE_CONFIG.plates.isSimpleKp).toBe(false);
    expect(PRODUCT_TYPE_CONFIG.piles.isSimpleKp).toBe(true);
    expect(PRODUCT_TYPE_CONFIG.steps.isSimpleKp).toBe(true);
    expect(PRODUCT_TYPE_CONFIG.marches.isSimpleKp).toBe(true);
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.isSimpleKp).toBe(true);
    expect(PRODUCT_TYPE_CONFIG.fbs.isSimpleKp).toBe(true);
  });

  it("maps each product to its ingest action", () => {
    expect(PRODUCT_TYPE_CONFIG.plates.ingestAction).toBe("ingest_plates");
    expect(PRODUCT_TYPE_CONFIG.piles.ingestAction).toBe("ingest_piles");
    expect(PRODUCT_TYPE_CONFIG.steps.ingestAction).toBe("ingest_steps");
    expect(PRODUCT_TYPE_CONFIG.marches.ingestAction).toBe("ingest_marches");
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.ingestAction).toBe("ingest_bridge_piles");
    expect(PRODUCT_TYPE_CONFIG.fbs.ingestAction).toBe("ingest_fbs");
  });

  it("maps each product to its metadata batches field", () => {
    expect(PRODUCT_TYPE_CONFIG.plates.batchesField).toBe("plate_batches");
    expect(PRODUCT_TYPE_CONFIG.piles.batchesField).toBe("pile_batches");
    expect(PRODUCT_TYPE_CONFIG.steps.batchesField).toBe("step_batches");
    expect(PRODUCT_TYPE_CONFIG.marches.batchesField).toBe("march_batches");
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.batchesField).toBe("bridge_pile_batches");
    expect(PRODUCT_TYPE_CONFIG.fbs.batchesField).toBe("fbs_batches");
  });

  it("provides display labels for every product", () => {
    expect(PRODUCT_TYPE_CONFIG.plates.labels).toEqual({ nounPlural: "Плиты", nounGenitivePlural: "плит" });
    expect(PRODUCT_TYPE_CONFIG.piles.labels).toEqual({ nounPlural: "Сваи", nounGenitivePlural: "свай" });
    expect(PRODUCT_TYPE_CONFIG.steps.labels).toEqual({ nounPlural: "Ступени", nounGenitivePlural: "ступеней" });
    expect(PRODUCT_TYPE_CONFIG.marches.labels).toEqual({ nounPlural: "Марши", nounGenitivePlural: "маршей" });
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.labels).toEqual({
      nounPlural: "Мостовые сваи",
      nounGenitivePlural: "мостовых свай",
    });
    expect(PRODUCT_TYPE_CONFIG.fbs.labels).toEqual({ nounPlural: "ФБС", nounGenitivePlural: "ФБС" });
  });
});

describe("getProductTypeConfig", () => {
  it("returns the config of the given product type", () => {
    expect(getProductTypeConfig("bridge_piles").endpointSegment).toBe("bridge-piles");
    expect(getProductTypeConfig("fbs").productType).toBe("fbs");
  });

  it("falls back to plates for null and undefined (legacy drafts)", () => {
    expect(getProductTypeConfig(undefined).productType).toBe("plates");
    expect(getProductTypeConfig(null).productType).toBe("plates");
  });

  it("falls back to plates for unknown runtime values", () => {
    expect(getProductTypeConfig("not-a-product" as ProductType).productType).toBe("plates");
  });

  it("falls back to plates for prototype-chain keys", () => {
    expect(getProductTypeConfig("constructor" as ProductType).productType).toBe("plates");
    expect(getProductTypeConfig("hasOwnProperty" as ProductType).productType).toBe("plates");
  });
});

describe("wizard finish messages", () => {
  it("maps every product to a distinct ingest action", () => {
    const actions = ALL_PRODUCT_TYPES.map((type) => PRODUCT_TYPE_CONFIG[type].ingestAction);
    expect(new Set(actions).size).toBe(ALL_PRODUCT_TYPES.length);
    for (const action of actions) {
      expect(action).toMatch(/^ingest_/);
    }
  });

  it("shares one ingest-required text", () => {
    expect(INGEST_REQUIRED_MESSAGE).toBe(
      "Сначала распознайте и получите хотя бы одну позицию в заказе.",
    );
  });

  it("keeps the three plates resolve-gate messages", () => {
    expect(Object.keys(RESOLVE_GATE_MESSAGES).sort()).toEqual([
      "resolve_invalid_widths",
      "resolve_unpriced_plates",
      "resolve_wide_plates",
    ]);
    expect(RESOLVE_GATE_MESSAGES.resolve_wide_plates).toBe(
      "Сначала примите решение по позициям шире стандартной.",
    );
    expect(RESOLVE_GATE_MESSAGES.resolve_invalid_widths).toBe(
      "Нестандартная ширина: замените на заводской рез или исключите позицию.",
    );
    expect(RESOLVE_GATE_MESSAGES.resolve_unpriced_plates).toBe(
      "Сначала примите решение по позициям без цены в прайсе.",
    );
  });
});
