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

  it("provides noun labels for every product", () => {
    expect(PRODUCT_TYPE_CONFIG.plates.labels.nounPlural).toBe("Плиты");
    expect(PRODUCT_TYPE_CONFIG.plates.labels.nounGenitivePlural).toBe("плит");
    expect(PRODUCT_TYPE_CONFIG.piles.labels.nounPlural).toBe("Сваи");
    expect(PRODUCT_TYPE_CONFIG.piles.labels.nounGenitivePlural).toBe("свай");
    expect(PRODUCT_TYPE_CONFIG.steps.labels.nounPlural).toBe("Ступени");
    expect(PRODUCT_TYPE_CONFIG.steps.labels.nounGenitivePlural).toBe("ступеней");
    expect(PRODUCT_TYPE_CONFIG.marches.labels.nounPlural).toBe("Марши");
    expect(PRODUCT_TYPE_CONFIG.marches.labels.nounGenitivePlural).toBe("маршей");
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.labels.nounPlural).toBe("Мостовые сваи");
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.labels.nounGenitivePlural).toBe("мостовых свай");
    expect(PRODUCT_TYPE_CONFIG.fbs.labels.nounPlural).toBe("ФБС");
    expect(PRODUCT_TYPE_CONFIG.fbs.labels.nounGenitivePlural).toBe("ФБС");
  });

  it("fills every input-step label for every product (no empty copy)", () => {
    for (const type of ALL_PRODUCT_TYPES) {
      const labels = PRODUCT_TYPE_CONFIG[type].labels;
      expect(labels.stepTitle, type).toBe(`Шаг 1. ${labels.nounPlural}`);
      expect(labels.listLabel, type).toBe(`Список ${labels.nounGenitivePlural}`);
      expect(labels.reviewListTitle, type).toBe(`${labels.listLabel} для расчёта`);
      expect(labels.emptySubtitle, type).toBe(
        `Вставьте текст списка ${labels.nounGenitivePlural} или загрузите фото таблицы.`,
      );
      expect(labels.aiHint, type).toBe(
        `Редкий сценарий: опишите, что сделать со списком ${labels.nounGenitivePlural}.`,
      );
      expect(labels.initialDescription, type).toBe(
        `Загрузите фото или вставьте список ${labels.nounGenitivePlural} для расчёта.`,
      );
      expect(labels.previewChangedMessage, type).toBe(
        `Изменён список ${labels.nounGenitivePlural} — нажмите «Список верен» для пересчёта состава.`,
      );
      expect(labels.placeholder.length, type).toBeGreaterThan(0);
      expect(labels.aiPlaceholder, type).toMatch(/^Например: /);
      expect(labels.addMoreDescription, type).toMatch(/^Добавьте ещё .+ или перейдите к оформлению клиента\.$/);
      expect(labels.previewEmptyMessage, type).toMatch(/^Список пуст — распознайте .+\.$/);
    }
  });

  it("keeps the exact current list placeholders", () => {
    expect(PRODUCT_TYPE_CONFIG.plates.labels.placeholder).toBe("ПБ 78-12-8п 2\n71-12-8 3\nПБ 66-12-8п 4");
    expect(PRODUCT_TYPE_CONFIG.piles.labels.placeholder).toBe("С120.35-12 B25 5\nС120.35-13и 3");
    expect(PRODUCT_TYPE_CONFIG.steps.labels.placeholder).toBe("ЛС11 10\nЛС14-1лев 5\nЛС11-Б-1 2");
    expect(PRODUCT_TYPE_CONFIG.marches.labels.placeholder).toBe("1ЛМ 27-11-14-4 B25 5\nЛМ 2,8 3");
    // Deliberate copy-paste leftovers from the pile step — preserved verbatim until the
    // customer confirms the canonical examples (see plan 2026-09-08 §5).
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.labels.placeholder).toBe("С120.35-12 B25 5\nС120.35-13и 3");
    expect(PRODUCT_TYPE_CONFIG.fbs.labels.placeholder).toBe("С120.35-12 B25 5\nС120.35-13и 3");
  });

  it("keeps the exact accusative-dependent copy (addMore / previewEmpty)", () => {
    expect(PRODUCT_TYPE_CONFIG.plates.labels.addMoreDescription).toBe(
      "Добавьте ещё плиты или перейдите к оформлению клиента.",
    );
    expect(PRODUCT_TYPE_CONFIG.piles.labels.addMoreDescription).toBe(
      "Добавьте ещё сваи или перейдите к оформлению клиента.",
    );
    expect(PRODUCT_TYPE_CONFIG.steps.labels.addMoreDescription).toBe(
      "Добавьте ещё ступени или перейдите к оформлению клиента.",
    );
    expect(PRODUCT_TYPE_CONFIG.marches.labels.addMoreDescription).toBe(
      "Добавьте ещё марши или перейдите к оформлению клиента.",
    );
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.labels.addMoreDescription).toBe(
      "Добавьте ещё мостовые сваи или перейдите к оформлению клиента.",
    );
    expect(PRODUCT_TYPE_CONFIG.fbs.labels.addMoreDescription).toBe(
      "Добавьте ещё ФБС или перейдите к оформлению клиента.",
    );
    expect(PRODUCT_TYPE_CONFIG.plates.labels.previewEmptyMessage).toBe("Список пуст — распознайте плиты.");
    expect(PRODUCT_TYPE_CONFIG.piles.labels.previewEmptyMessage).toBe("Список пуст — распознайте сваи.");
    expect(PRODUCT_TYPE_CONFIG.steps.labels.previewEmptyMessage).toBe("Список пуст — распознайте ступени.");
    expect(PRODUCT_TYPE_CONFIG.marches.labels.previewEmptyMessage).toBe("Список пуст — распознайте марши.");
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.labels.previewEmptyMessage).toBe(
      "Список пуст — распознайте мостовые сваи.",
    );
    expect(PRODUCT_TYPE_CONFIG.fbs.labels.previewEmptyMessage).toBe("Список пуст — распознайте ФБС.");
  });

  it("keeps the exact AI placeholders", () => {
    expect(PRODUCT_TYPE_CONFIG.plates.labels.aiPlaceholder).toBe("Например: убери строки с 6п");
    expect(PRODUCT_TYPE_CONFIG.piles.labels.aiPlaceholder).toBe("Например: убери строки с B15");
    expect(PRODUCT_TYPE_CONFIG.steps.labels.aiPlaceholder).toBe("Например: убери строки с ЛС11");
    expect(PRODUCT_TYPE_CONFIG.marches.labels.aiPlaceholder).toBe("Например: убери строки с B15");
    expect(PRODUCT_TYPE_CONFIG.bridge_piles.labels.aiPlaceholder).toBe("Например: убери строки с B15");
    expect(PRODUCT_TYPE_CONFIG.fbs.labels.aiPlaceholder).toBe("Например: убери строки с B15");
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
